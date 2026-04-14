/**
 * 为 HTML 文本推断适合写入 PPTX 的字体分类结果
 *
 * 支持 4 类：
 * - latin   : 拉丁/西文主字体
 * - ea      : East Asian，中日韩及全角标点主字体
 * - cs      : Complex Script，阿拉伯/希伯来/印地语/泰语等复杂脚本主字体
 * - symbol  : 数学/符号/箭头/公式等更适合单独归类的字体
 *
 * 设计原则：
 * 1. 不试图精确还原浏览器底层最终 fallback 字体文件
 * 2. 而是为 PPTX 选出"最接近当前渲染效果"的代表字体
 * 3. 优先保留作者在 HTML/CSS 中表达的字体意图
 * 4. 当本地 font-family 无法较好解释渲染时，再逐层 fallback
 *
 * 候选池分 3 层：
 * - 第 1 层：当前元素 computed font-family
 * - 第 2 层：祖先元素 computed font-family
 * - 第 3 层：Office/PPT 友好的 fallback 字体池
 *
 * 核心输出：
 * {
 *   latin: 'Arial',
 *   ea: 'Microsoft YaHei',
 *   cs: 'Arial',
 *   symbol: 'Cambria Math',
 *   details: { ... }
 * }
 *
 * 适用场景：
 * - HTML 转 PPTX
 * - 网页内容导出为 PowerPoint
 * - 对多语言文本按 script 选择代表字体
 *
 * 注意：
 * - 这是"为 PPTX 选代表字体"的启发式算法
 * - 不是浏览器真实字形 fallback 的官方 API
 *
 * @param {HTMLElement} element 要检测的 DOM 元素
 * @param {string|null} customText 可选，自定义文本；不传则使用 element.textContent
 * @param {object} options 可选配置
 * @returns {{
 *   latin: string,
 *   ea: string,
 *   cs: string,
 *   symbol: string,
 *   details: object
 * }}
 */
function detectFontsForPPTX(element, customText = null, options = {}) {
  if (!(element instanceof HTMLElement)) {
    throw new TypeError("detectFontsForPPTX: element must be an HTMLElement");
  }

  const config = {
    /**
     * 每类脚本最多抽取多少字符作为样本。
     * 样本越长越稳，但也会略微增加 measureText 成本。
     */
    maxSampleLength: options.maxSampleLength ?? 32,

    /**
     * 当前元素字体栈的匹配阈值。
     * 分数越低说明越像当前渲染。
     * 如果当前元素候选池已足够好，就不再向上/向外扩展。
     */
    currentThreshold: options.currentThreshold ?? 0.35,

    /**
     * 祖先候选池的匹配阈值。
     * 若当前层不够好，则引入祖先候选池。
     * 若祖先层足够好，就不再进入 Office fallback 层。
     */
    ancestorThreshold: options.ancestorThreshold ?? 0.75,

    /**
     * 是否将 generic family / 某些不稳定字体名 规范化为更适合 PPTX 的字体名。
     * 建议保持 true。
     */
    preferOfficeFonts: options.preferOfficeFonts ?? true,

    /**
     * symbol 内容占比达到该阈值时，才认为它是"显著 symbol 内容"。
     * 否则 symbol 最终可退化为 latin。
     */
    symbolRatioThreshold: options.symbolRatioThreshold ?? 0.12,

    /**
     * 当字体栈中出现明显 math/symbol 字体时，也可触发 symbol 独立检测。
     */
    forceSymbolByFontFamily: options.forceSymbolByFontFamily ?? true,
  };

  const style = window.getComputedStyle(element);
  const text = String(customText ?? element.textContent ?? "").trim();

  // Office/PPT 友好的 fallback 池。
  // 注意：这些不是"浏览器真实系统 fallback"，而是面向 PPT 输出更稳定的代表字体池。
  const OFFICE_LATIN_FONTS = getOfficeLatinFonts();
  const OFFICE_EA_FONTS = getOfficeEAFonts();
  const OFFICE_CS_FONTS = getOfficeCSFonts();
  const OFFICE_SYMBOL_FONTS = getOfficeSymbolFonts();

  // 没有文本时直接返回一组稳妥缺省值。
  if (!text) {
    return {
      latin: "Arial",
      ea: "Microsoft YaHei",
      cs: "Arial",
      symbol: "Cambria Math",
      details: {
        text: "",
        samples: {
          latin: "",
          ea: "",
          cs: "",
          symbol: "",
        },
        currentFontStack: normalizeFontList(style.fontFamily),
        ancestorFontStack: [],
        scores: {
          latin: 0,
          ea: 0,
          cs: 0,
          symbol: 0,
        },
        stages: {
          latin: "office",
          ea: "office",
          cs: "office",
          symbol: "office",
        },
        scriptCounts: {
          latin: 0,
          ea: 0,
          cs: 0,
          symbol: 0,
          total: 0,
        },
        hasSignificantSymbolContent: false,
      },
    };
  }

  // 允许外部传入 Canvas 2D context 以便批量复用。
  const ctx = options.ctx || getSharedCanvasContext();

  if (!ctx) {
    return {
      latin: "Arial",
      ea: "Microsoft YaHei",
      cs: "Arial",
      symbol: "Cambria Math",
      details: {
        text,
        error: "Canvas 2D context unavailable",
      },
    };
  }

  // 构造 canvas font 前缀，不含 font-family。
  // font-family 会在 measure 时动态拼接。
  const baseStyle = buildCanvasFontBase(style);

  // 第 1 层：当前元素自己的 font-family
  const currentFontStack = normalizeFontList(style.fontFamily);

  // 第 2 层：祖先元素 font-family（去重）
  const ancestorFontStack = collectAncestorFontFamilies(element);

  // 当前 + 祖先，作为第 2 阶段候选池
  const currentPlusAncestors = uniqueFonts([
    ...currentFontStack,
    ...ancestorFontStack,
  ]);

  // 提前分析文本中各类 script 分布，用于：
  // - 生成样本
  // - 判断 symbol 是否值得单独处理
  const scriptCounts = countScripts(text);
  const totalCount = Math.max(scriptCounts.total, 1);

  // 为每个类别抽取样本。
  // 样本是后续 measureText 的依据。
  const latinSample = collectSample(
    text,
    isLatinLikeChar,
    "The quick brown fox jumps 123 ABC xyz",
    config.maxSampleLength,
  );

  const eaSample = collectSample(
    text,
    isEastAsianChar,
    "测试漢字かなカナ한글，。、（）",
    config.maxSampleLength,
  );

  const csSample = collectSample(
    text,
    isComplexScriptChar,
    "العربية עברית हिन्दी ไทย",
    config.maxSampleLength,
  );

  const symbolSample = collectSample(
    text,
    isSymbolLikeChar,
    "∑∫√∞≈≠≤≥→⇒αβγθ",
    config.maxSampleLength,
  );

  // 判断这段文本是否值得单独做 symbol 检测。
  // 触发条件：
  // 1. symbol-like 字符比例达到一定阈值
  // 2. 或 font-family 里出现明显的 math/symbol 字体
  const symbolRatio = scriptCounts.symbol / totalCount;
  const hasMathLikeFontFamily =
    config.forceSymbolByFontFamily &&
    containsMathOrSymbolFontName(currentPlusAncestors);

  const hasSignificantSymbolContent =
    symbolRatio >= config.symbolRatioThreshold || hasMathLikeFontFamily;

  // 分别解析 4 类字体。
  const latinResult = resolveScriptFont({
    ctx,
    baseStyle,
    fullFontFamily: style.fontFamily,
    sample: latinSample,
    currentFonts: currentFontStack,
    ancestorFonts: currentPlusAncestors,
    officeFonts: OFFICE_LATIN_FONTS,
    currentThreshold: config.currentThreshold,
    ancestorThreshold: config.ancestorThreshold,
    normalizeResult: (name) =>
      normalizeToPPTLatinFont(
        name,
        OFFICE_LATIN_FONTS,
        config.preferOfficeFonts,
      ),
  });

  const eaResult = resolveScriptFont({
    ctx,
    baseStyle,
    fullFontFamily: style.fontFamily,
    sample: eaSample,
    currentFonts: currentFontStack,
    ancestorFonts: currentPlusAncestors,
    officeFonts: OFFICE_EA_FONTS,
    currentThreshold: config.currentThreshold,
    ancestorThreshold: config.ancestorThreshold,
    normalizeResult: (name) =>
      normalizeToPPTEAFont(name, OFFICE_EA_FONTS, config.preferOfficeFonts),
    checkGlyphCoverage: true,
  });

  const csResult = resolveScriptFont({
    ctx,
    baseStyle,
    fullFontFamily: style.fontFamily,
    sample: csSample,
    currentFonts: currentFontStack,
    ancestorFonts: currentPlusAncestors,
    officeFonts: OFFICE_CS_FONTS,
    currentThreshold: config.currentThreshold,
    ancestorThreshold: config.ancestorThreshold,
    normalizeResult: (name) =>
      normalizeToPPTCSFont(name, OFFICE_CS_FONTS, config.preferOfficeFonts),
    checkGlyphCoverage: true,
  });

  // symbol 的处理稍微特殊：
  // - 若内容里没有明显 symbol 特征，则不强行单独探测
  // - 这时让 symbol 退化为更稳妥的 latin 结果
  let symbolResult;
  if (hasSignificantSymbolContent) {
    symbolResult = resolveScriptFont({
      ctx,
      baseStyle,
      fullFontFamily: style.fontFamily,
      sample: symbolSample,
      currentFonts: currentFontStack,
      ancestorFonts: currentPlusAncestors,
      officeFonts: OFFICE_SYMBOL_FONTS,
      currentThreshold: config.currentThreshold,
      ancestorThreshold: config.ancestorThreshold,
      normalizeResult: (name) =>
        normalizeToPPTSymbolFont(
          name,
          OFFICE_SYMBOL_FONTS,
          config.preferOfficeFonts,
        ),
      checkGlyphCoverage: true,
    });
  } else {
    symbolResult = {
      font: latinResult.font,
      rawFont: latinResult.rawFont,
      score: latinResult.score,
      stage: "latin-fallback",
      candidates: latinResult.candidates.slice(),
    };
  }

  return {
    latin: latinResult.font,
    ea: eaResult.font,
    cs: csResult.font,
    symbol: symbolResult.font,
    details: {
      text,
      samples: {
        latin: latinSample,
        ea: eaSample,
        cs: csSample,
        symbol: symbolSample,
      },
      currentFontStack,
      ancestorFontStack,
      candidates: {
        latin: latinResult.candidates,
        ea: eaResult.candidates,
        cs: csResult.candidates,
        symbol: symbolResult.candidates,
      },
      scores: {
        latin: latinResult.score,
        ea: eaResult.score,
        cs: csResult.score,
        symbol: symbolResult.score,
      },
      stages: {
        latin: latinResult.stage,
        ea: eaResult.stage,
        cs: csResult.stage,
        symbol: symbolResult.stage,
      },
      rawFonts: {
        latin: latinResult.rawFont,
        ea: eaResult.rawFont,
        cs: csResult.rawFont,
        symbol: symbolResult.rawFont,
      },
      scriptCounts,
      symbolRatio,
      hasMathLikeFontFamily,
      hasSignificantSymbolContent,
    },
  };
}

/* =========================================================
 * 主流程：为某一类 script 解析代表字体
 * ========================================================= */

/**
 * 解析某一类 script 应映射到哪个字体。
 *
 * 过程分 3 阶段：
 * 1. 仅在当前元素 font-family 中找最优解
 * 2. 若不够可信，则扩大到"当前 + 祖先"候选池
 * 3. 若仍不够可信，则引入 Office/PPT fallback 池
 *
 * @param {object} params
 * @returns {{
 *   font: string,
 *   rawFont: string,
 *   score: number,
 *   stage: string,
 *   candidates: string[]
 * }}
 */
function resolveScriptFont({
  ctx,
  baseStyle,
  fullFontFamily,
  sample,
  currentFonts,
  ancestorFonts,
  officeFonts,
  currentThreshold,
  ancestorThreshold,
  normalizeResult,
  checkGlyphCoverage,
}) {
  // "目标渲染"：使用完整 computed font-family 去测样本。
  // 这不是"真实底层字体"，但足以作为候选拟合的参考基准。
  const actual = measureWithStack(ctx, baseStyle, fullFontFamily, sample);

  // 是否启用字形覆盖检测（双 fallback 法）。
  // 目的：当候选字体不含目标 script 的字形时，给予惩罚。
  const glyphCheck = checkGlyphCoverage;

  // 渐进式搜索：每阶段只测量增量候选字体，与上阶段最佳竞争。
  const currentSet = new Set(currentFonts.map((f) => f.toLowerCase()));

  // 第 1 阶段：只看当前元素字体栈。
  const stage1 = detectBestMatch(ctx, baseStyle, sample, actual, currentFonts, glyphCheck);
  if (stage1.font && stage1.score <= currentThreshold) {
    return {
      font: normalizeResult(stage1.font),
      rawFont: stage1.font,
      score: stage1.score,
      stage: "current",
      candidates: currentFonts.slice(),
    };
  }

  // 第 2 阶段：只测量祖先候选池中的增量字体，与 stage1 最佳竞争。
  const ancestorOnly = ancestorFonts.filter(
    (f) => !currentSet.has(f.toLowerCase()),
  );
  const stage2Incremental = detectBestMatch(
    ctx,
    baseStyle,
    sample,
    actual,
    ancestorOnly,
    glyphCheck,
  );
  // 使用 <= 而非 <：当分数相同时，优先选择后续阶段的候选字体。
  // 语义：更靠后的候选池（祖先 > 当前，Office > 祖先）在平局时胜出，
  // 因为 Office 字体池的字体名更适合写入 PPTX，且当 glyph check 惩罚
  // 导致所有候选都获得高分时，能确保 Office 池的字体名被选出而非
  // 透传前序阶段中不合适的字体名（如 Roboto）。
  const stage2 =
    stage2Incremental.score <= stage1.score ? stage2Incremental : stage1;

  if (stage2.font && stage2.score <= ancestorThreshold) {
    return {
      font: normalizeResult(stage2.font),
      rawFont: stage2.font,
      score: stage2.score,
      stage: "ancestor",
      candidates: uniqueFonts([...ancestorFonts]),
    };
  }

  // 第 3 阶段：只测量 Office fallback 池中的增量字体。
  const knownSet = new Set(ancestorFonts.map((f) => f.toLowerCase()));
  const officeOnly = officeFonts.filter(
    (f) => !knownSet.has(f.toLowerCase()),
  );
  const stage3Incremental = detectBestMatch(
    ctx,
    baseStyle,
    sample,
    actual,
    officeOnly,
    glyphCheck,
  );
  const stage3 =
    stage3Incremental.score <= stage2.score ? stage3Incremental : stage2;

  return {
    font: normalizeResult(stage3.font || officeFonts[0]),
    rawFont: stage3.font || officeFonts[0],
    score: stage3.score,
    stage: "office",
    candidates: uniqueFonts([...ancestorFonts, ...officeFonts]),
  };
}

/**
 * 在给定候选字体列表中，找出与"目标渲染"最接近的字体。
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {string} baseStyle
 * @param {string} sample
 * @param {object} actual
 * @param {string[]} fonts
 * @returns {{ font: string, score: number }}
 */
function detectBestMatch(ctx, baseStyle, sample, actual, fonts, glyphCheck) {
  const candidates = fonts.filter(Boolean);
  if (!candidates.length) {
    return {
      font: "",
      score: Number.POSITIVE_INFINITY,
    };
  }

  let bestFont = candidates[candidates.length - 1];
  let bestScore = Number.POSITIVE_INFINITY;

  for (const fontName of candidates) {
    const test = measureSingleFont(ctx, baseStyle, fontName, sample);
    let score = calcMetricScore(test, actual);

    // 字形覆盖检查（双 fallback 法）：
    // 用两个不同的 fallback（monospace 和 serif）分别测量候选字体。
    // - 若字体本身包含目标字形，两个 fallback 不会被触发，度量值一致。
    // - 若字体不含目标字形，浏览器会沿两条不同的 fallback 路径渲染
    //  （monospace → 系统 CJK 兜底字体 vs serif → 另一个系统 CJK 兜底字体），
    //   二者通常会产生不同的度量值（如 macOS 上 PingFang SC vs Songti SC）。
    // 当检测到两条路径不一致时，说明候选字体缺少该字形，给予大幅惩罚。
    if (glyphCheck) {
      if (!fontHasGlyphForSample(ctx, baseStyle, fontName, sample)) {
        score += 1000;
      }
    }

    if (score < bestScore) {
      bestScore = score;
      bestFont = fontName;
    }

    // 极小分数说明已非常接近，可以提前结束。
    if (score <= 0.0001) break;
  }

  return {
    font: bestFont,
    score: bestScore,
  };
}

/* =========================================================
 * 文本样本与 script 识别
 * ========================================================= */

/**
 * 从原文本中抽取某一类字符，组成样本字符串。
 * 若原文中没有该类字符，则使用 fallback 样本。
 *
 * @param {string} text
 * @param {(ch: string) => boolean} predicate
 * @param {string} fallback
 * @param {number} maxLen
 * @returns {string}
 */
function collectSample(text, predicate, fallback, maxLen) {
  let result = "";
  let count = 0;

  for (const ch of text) {
    // 跳过空白/不可见字符，避免污染测量样本（与 countScripts 保持一致）。
    if (/\s/.test(ch)) continue;

    if (predicate(ch)) {
      result += ch;
      count += 1;
      // 按 codepoint 计数而非 UTF-16 code unit，正确处理 SMP 字符。
      if (count >= maxLen) break;
    }
  }

  return result || fallback;
}

/**
 * 统计文本中 4 类脚本的大致数量。
 * 注意这里是工程上的粗分，不是严格 Unicode Script 规范分类。
 *
 * @param {string} text
 * @returns {{
 *   latin: number,
 *   ea: number,
 *   cs: number,
 *   symbol: number,
 *   other: number,
 *   total: number
 * }}
 */
function countScripts(text) {
  const counts = {
    latin: 0,
    ea: 0,
    cs: 0,
    symbol: 0,
    other: 0,
    total: 0,
  };

  for (const ch of text) {
    // 忽略纯空白字符，不计入 total。
    if (/\s/.test(ch)) continue;

    counts.total += 1;

    if (isEastAsianChar(ch)) {
      counts.ea += 1;
      continue;
    }

    if (isComplexScriptChar(ch)) {
      counts.cs += 1;
      continue;
    }

    if (isLatinLikeChar(ch)) {
      counts.latin += 1;
      continue;
    }

    if (isSymbolLikeChar(ch)) {
      counts.symbol += 1;
      continue;
    }

    counts.other += 1;
  }

  return counts;
}

/**
 * Latin-like：
 * - 基本拉丁字母 + 数字
 * - Latin-1 Supplement / Extended-A / Extended-B
 * - Latin Extended Additional（越南语重音字母等）
 * - General Punctuation（em-dash、引号、省略号等排版标点，
 *   在 PPTX 中由 latin 字体渲染）
 *
 * 这里故意不把 ASCII 标点都算进来，
 * 因为很多基础标点在视觉上区分力不足，作为样本意义有限。
 */
function isLatinLikeChar(ch) {
  return /[A-Za-z0-9\u00C0-\u024F\u1E00-\u1EFF\u2000-\u206F]/.test(ch);
}

/**
 * East Asian：
 * - CJK Unified Ideographs（含 BMP 及所有扩展区块 B~H）
 * - CJK 标点
 * - 全角字符
 * - 日文平假名/片假名
 * - 韩文 Hangul
 *
 * 使用 Unicode property escapes 自动覆盖 SMP 中的 CJK Extension B~H，
 * 无需手动枚举不断增长的区块范围。
 */
function isEastAsianChar(ch) {
  // BMP 快速路径：CJK 标点 + 全角 + 假名 + Hangul
  if (
    /[\u3000-\u303F\uFF00-\uFFEF\u3040-\u30FF\uAC00-\uD7AF]/.test(ch)
  ) {
    return true;
  }
  // Han 脚本覆盖所有 CJK Unified Ideographs（BMP + SMP 扩展）
  return /\p{Script=Han}/u.test(ch);
}

/**
 * Complex Script：
 * 覆盖 OOXML 规范中属于 cs 的主要脚本范围：
 * - 亚美尼亚 \u0530-\u058F
 * - 希伯来   \u0590-\u05FF
 * - 阿拉伯   \u0600-\u06FF, \u0750-\u077F, \u08A0-\u08FF
 * - 天城文   \u0900-\u097F
 * - 孟加拉   \u0980-\u09FF
 * - 古木基   \u0A00-\u0A7F
 * - 古吉拉特 \u0A80-\u0AFF
 * - 奥里亚   \u0B00-\u0B7F
 * - 泰米尔   \u0B80-\u0BFF
 * - 泰卢固   \u0C00-\u0C7F
 * - 卡纳达   \u0C80-\u0CFF
 * - 马拉雅拉姆 \u0D00-\u0D7F
 * - 僧伽罗   \u0D80-\u0DFF
 * - 泰文     \u0E00-\u0E7F
 * - 老挝     \u0E80-\u0EFF
 * - 藏文     \u0F00-\u0FFF
 * - 缅甸     \u1000-\u109F
 * - 格鲁吉亚 \u10A0-\u10FF
 * - 埃塞俄比亚 \u1200-\u137F
 */
function isComplexScriptChar(ch) {
  return /[\u0530-\u058F\u0590-\u05FF\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\u0900-\u097F\u0980-\u09FF\u0A00-\u0A7F\u0A80-\u0AFF\u0B00-\u0B7F\u0B80-\u0BFF\u0C00-\u0C7F\u0C80-\u0CFF\u0D00-\u0D7F\u0D80-\u0DFF\u0E00-\u0E7F\u0E80-\u0EFF\u0F00-\u0FFF\u1000-\u109F\u10A0-\u10FF\u1200-\u137F]/.test(
    ch,
  );
}

/**
 * Symbol-like：
 * 包括常见：
 * - Greek and Coptic（公式中常见）
 * - Letterlike Symbols
 * - Arrows
 * - Math Operators
 * - Misc Technical
 * - Box Drawing
 * - Geometric Shapes
 * - Misc Symbols
 * - Dingbats
 * - Supplemental Math Operators
 *
 * 注意：
 * - 刻意排除 General Punctuation (\u2000-\u206F)，
 *   其中 em-dash/引号/省略号等在 PPTX 中由 latin 字体渲染，
 *   归入 symbol 会错误触发数学字体检测。
 * - 希腊字母并不总是"symbol"，但在公式/LaTeX 场景中很常见。
 */
function isSymbolLikeChar(ch) {
  // BMP 范围
  if (
    /[\u0370-\u03FF\u2100-\u214F\u2190-\u21FF\u2200-\u22FF\u2300-\u23FF\u2500-\u257F\u25A0-\u25FF\u2600-\u26FF\u2700-\u27BF\u27C0-\u27EF]/.test(
      ch,
    )
  ) {
    return true;
  }
  // SMP：Mathematical Alphanumeric Symbols (U+1D400-1D7FF)、
  //       Emoji 常用区块
  return /[\u{1D400}-\u{1D7FF}\u{1F600}-\u{1F64F}\u{1F300}-\u{1F5FF}\u{1F680}-\u{1F6FF}\u{1F900}-\u{1F9FF}]/u.test(
    ch,
  );
}

/**
 * 判断某个字体候选池中是否出现明显的数学/符号字体名。
 * 这对 LaTeX/MathJax/KaTeX 场景很有帮助。
 */
function containsMathOrSymbolFontName(fonts) {
  const patterns = [
    /cambria math/i,
    /stix/i,
    /latin modern math/i,
    /xits/i,
    /\bmath\b/i,
    /\bsymbol\b/i,
    /wingdings/i,
    /webdings/i,
    /mt extra/i,
    /seg(oe)? ui symbol/i,
    /katex/i,
  ];

  return fonts.some((font) => patterns.some((p) => p.test(font)));
}

/* =========================================================
 * 字体候选池收集
 * ========================================================= */

/**
 * 收集祖先节点的 font-family。
 * 注意：
 * - 从 parentElement 开始，不重复包含当前元素
 * - 返回结果去重
 * - 顺序保留"越近的祖先越靠前"的信息
 */
function collectAncestorFontFamilies(element) {
  const fonts = [];
  const seen = new Set();

  let node = element.parentElement;
  while (node && node.nodeType === 1) {
    const style = window.getComputedStyle(node);
    const list = normalizeFontList(style.fontFamily);

    for (const font of list) {
      const key = font.toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        fonts.push(font);
      }
    }

    node = node.parentElement;
  }

  return fonts;
}

/**
 * 将 CSS font-family 字符串拆成字体数组，并做基础清洗：
 * - split(',')
 * - trim
 * - 去掉首尾引号
 * - 去重
 */
function normalizeFontList(fontFamily) {
  if (!fontFamily || typeof fontFamily !== "string") return [];

  return uniqueFonts(
    fontFamily
      .split(",")
      .map((f) => f.trim())
      .map(stripQuotes)
      .filter(Boolean),
  );
}

/**
 * 大小写不敏感去重，但保留首次出现的原始形式。
 */
function uniqueFonts(list) {
  const result = [];
  const seen = new Set();

  for (const item of list) {
    const value = String(item || "").trim();
    if (!value) continue;

    const key = value.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      result.push(value);
    }
  }

  return result;
}

/**
 * 去掉字体名前后的单/双引号。
 */
function stripQuotes(name) {
  return String(name || "")
    .replace(/^["']|["']$/g, "")
    .trim();
}

/* =========================================================
 * Canvas 文本测量
 * ========================================================= */

/**
 * 懒加载的共享 Canvas 2D context 单例。
 * 避免每次调用 detectFontsForPPTX 都创建新 canvas。
 */
let _sharedCtx = null;
function getSharedCanvasContext() {
  if (_sharedCtx) return _sharedCtx;
  if (typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  _sharedCtx = canvas.getContext("2d", { alpha: false });
  return _sharedCtx;
}

/**
 * 构造 canvas font 前缀，保留：
 * - font-style
 * - font-variant
 * - font-weight
 * - font-size / line-height
 *
 * 注意：canvas ctx.font 只接受标准 CSS font shorthand，
 * 不支持 font-stretch 等非简写属性。
 * fontSize 和 lineHeight 必须拼为 "16px/24px" 无空格。
 *
 * 不包含 font-family，方便后续动态拼接候选字体。
 */
function buildCanvasFontBase(style) {
  const sizeFragment =
    style.lineHeight && style.lineHeight !== "normal"
      ? `${style.fontSize}/${style.lineHeight}`
      : style.fontSize;

  return [style.fontStyle, style.fontVariant, style.fontWeight, sizeFragment]
    .filter(Boolean)
    .join(" ");
}

/**
 * 使用完整 font-family 栈测量样本。
 * 这作为"目标渲染"的近似。
 */
function measureWithStack(ctx, baseStyle, fullFontFamily, sample) {
  ctx.font = `${baseStyle} ${fullFontFamily}`;
  return readTextMetrics(ctx.measureText(sample));
}

/**
 * 用单个字体名测量样本。
 *
 * 关键细节：附加 monospace 作为"哨兵 fallback"。
 * 当 fontName 未安装时，浏览器会回退到 monospace，
 * 其宽度特征与多数 serif/sans-serif 差异显著，
 * 从而在评分中自然惩罚未安装的字体。
 */
function measureSingleFont(ctx, baseStyle, fontName, sample) {
  const family = quoteFontFamily(fontName);
  ctx.font = `${baseStyle} ${family}, monospace`;
  return readTextMetrics(ctx.measureText(sample));
}

/**
 * 检测候选字体是否真正包含目标样本的字形（单哨兵法）。
 *
 * 原理：
 * 比较 "candidate, monospace" 与裸 "monospace" 的测量结果。
 * - 若候选字体包含目标字形，添加它会改变渲染结果，两者度量不同。
 * - 若候选字体不含目标字形，浏览器的 fallback 路径完全一致：
 *   "candidate, monospace" → candidate 缺字形 → monospace → 系统兜底
 *   "monospace"            → monospace → 系统兜底
 *   二者经过完全相同的 fallback 链，度量必然一致。
 *
 * 例外：当候选字体恰好 IS 系统 CJK 兜底字体（如 macOS 上的 PingFang SC）时，
 * 两者度量也相同，产生假阴性。但这在实践中可接受，因为这些字体会被
 * normalizeToPPTEAFont 映射到 Office 字体（如 Microsoft YaHei）。
 *
 * 为何不用双 fallback 法（对比 monospace/serif 两条路径）：
 * Chrome 的 per-glyph CJK fallback 在不同版本/平台上表现不稳定。
 * 某些环境下 "Roboto, serif" 和 "Roboto, monospace" 的 CJK fallback 都指向
 * 同一个 PingFang SC（Chrome 绕过 serif sentinel 直接走系统 CJK 回退),
 * 导致双 fallback 法误判 Roboto "有 CJK 字形"。
 * 单哨兵法不依赖两条 fallback 路径的分叉，因此跨平台更稳定。
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {string} baseStyle
 * @param {string} fontName
 * @param {string} sample
 * @returns {boolean} true 表示字体很可能包含目标字形
 */
function fontHasGlyphForSample(ctx, baseStyle, fontName, sample) {
  const family = quoteFontFamily(fontName);

  // 带候选字体测量
  ctx.font = `${baseStyle} ${family}, monospace`;
  const withCandidate = readTextMetrics(ctx.measureText(sample));

  // 裸 monospace 基线
  ctx.font = `${baseStyle} monospace`;
  const bareMono = readTextMetrics(ctx.measureText(sample));

  // 若候选字体改变了渲染结果（度量差异 ≥ 0.01），说明它提供了字形。
  return calcMetricScore(withCandidate, bareMono) >= 0.01;
}

/**
 * 为 canvas font-family 做安全引用。
 * - 简单单词可直接返回
 * - 含空格/特殊字符则加双引号
 */
function quoteFontFamily(name) {
  const n = stripQuotes(name);
  if (/^[a-zA-Z0-9_-]+$/.test(n)) return n;
  return `"${n.replace(/"/g, '\\"')}"`;
}

/**
 * 读取 measureText 返回的核心 metrics。
 */
function readTextMetrics(m) {
  return {
    width: m.width || 0,
    ascent: m.actualBoundingBoxAscent || 0,
    descent: m.actualBoundingBoxDescent || 0,
    left: m.actualBoundingBoxLeft || 0,
    right: m.actualBoundingBoxRight || 0,
  };
}

/**
 * 计算候选字体与"目标渲染"的距离分数。
 *
 * 分数越小表示越接近。
 * 所有维度均按各自 baseline 归一化为比例偏差，
 * 使权重系数真正反映各维度的预期重要性。
 *
 * 权重语义：
 * - width  (10)：整体宽度是最强区分信号
 * - ascent  (4)：上缘高度区分字体家族
 * - descent (4)：下缘深度同上
 * - left    (2)：左侧 bearing 辅助区分
 * - right   (2)：右侧 bearing 辅助区分
 */
function calcMetricScore(test, actual) {
  const safeDiv = (diff, base) => Math.abs(diff) / Math.max(Math.abs(base), 1);

  const widthDiff = safeDiv(test.width - actual.width, actual.width);
  const ascentDiff = safeDiv(test.ascent - actual.ascent, actual.ascent);
  const descentDiff = safeDiv(test.descent - actual.descent, actual.descent);
  const leftDiff = safeDiv(test.left - actual.left, actual.left);
  const rightDiff = safeDiv(test.right - actual.right, actual.right);

  return (
    widthDiff * 10 +
    ascentDiff * 4 +
    descentDiff * 4 +
    leftDiff * 2 +
    rightDiff * 2
  );
}

/* =========================================================
 * Office / PPT 友好 fallback 字体池
 * ========================================================= */

/**
 * Latin fallback 字体池。
 */
function getOfficeLatinFonts() {
  return [
    "Calibri",
    "Arial",
    "Times New Roman",
    "Verdana",
    "Tahoma",
    "Segoe UI",
  ];
}

/**
 * East Asian fallback 字体池。
 *
 * 重要：此池中的字体必须是 PowerPoint 实际认识并可以嵌入/渲染的字体。
 * Apple 专属字体（PingFang SC、Hiragino Sans GB 等）不在此列——
 * 它们在 macOS 上有效，但 PowerPoint 字体选择器中不会出现，
 * 写入 .pptx 后在 Windows 端打开会 fallback 到默认字体。
 *
 * 这些字体由 Office 安装包自带，即使在 macOS 上安装 Office 后也可用。
 */
function getOfficeEAFonts() {
  return [
    "Microsoft YaHei",
    "SimSun",
    "SimHei",
    "DengXian",
    "KaiTi",
    "FangSong",
    "Meiryo",
    "MS Gothic",
    "MS Mincho",
    "Malgun Gothic",
    "Batang",
  ];
}

/**
 * Complex Script fallback 字体池。
 */
function getOfficeCSFonts() {
  return [
    "Arial",
    "Times New Roman",
    "Tahoma",
    "Segoe UI",
    "Nirmala UI",
    "Mangal",
    "Aparajita",
    "Kokila",
    "Shruti",
    "Raavi",
    "Gautami",
    "Vrinda",
    "Leelawadee UI",
    "David",
  ];
}

/**
 * Symbol / Math fallback 字体池。
 */
function getOfficeSymbolFonts() {
  return [
    "Cambria Math",
    "Segoe UI Symbol",
    "Arial Unicode MS",
    "Symbol",
    "MT Extra",
    "Times New Roman",
    "Arial",
    "Microsoft YaHei",
  ];
}

/* =========================================================
 * 输出到 PPTX 前的字体规范化
 * ========================================================= */

/**
 * Latin 字体规范化：
 * - generic family -> 稳定 Office 字体
 * - 某些苹果系/系统 UI 字体 -> Arial / Segoe UI 等更稳字体
 * - 其他字体尽量保留原名，尊重设计意图
 */
function normalizeToPPTLatinFont(name, officeFonts, preferOfficeFonts = true) {
  const font = stripQuotes(String(name || "").trim());
  if (!font) return officeFonts[0] || "Arial";

  const lower = font.toLowerCase();

  const aliasMap = {
    "system-ui": "Arial",
    "sans-serif": "Arial",
    serif: "Times New Roman",
    monospace: "Courier New",
    "ui-sans-serif": "Arial",
    "ui-serif": "Times New Roman",
    "ui-monospace": "Courier New",
    "helvetica neue": "Arial",
    helvetica: "Arial",
    arialmt: "Arial",
    "sf pro text": "Arial",
    "sf pro display": "Arial",
    "segoe ui": "Segoe UI",
  };

  if (aliasMap[lower]) {
    return aliasMap[lower];
  }

  // 若不强制偏 Office，则保留原字体名。
  if (!preferOfficeFonts) {
    return font;
  }

  // 默认仍保留原名。
  // 这里只规范化那些明显不稳定 / 泛化的字体名。
  return font;
}

/**
 * East Asian 字体规范化。
 *
 * 所有映射的目标字体必须是 PowerPoint 实际认识的字体。
 * macOS/Linux 专属 CJK 字体（PingFang SC、Hiragino、Noto Sans CJK 等）
 * 虽然浏览器能渲染，但必须映射到 Office 对应字体，
 * 否则 .pptx 在 Windows 上打开会丢失字体。
 *
 * 映射参考：
 * - 现代黑体类（PingFang SC、Heiti SC、Noto Sans CJK）→ Microsoft YaHei
 * - 宋体类（Songti SC、STSong）→ SimSun
 * - 楷体类（STKaiti）→ KaiTi
 * - 日文专属   → Meiryo / MS Gothic
 * - 韩文专属   → Malgun Gothic
 */
function normalizeToPPTEAFont(name, officeFonts, preferOfficeFonts = true) {
  const font = stripQuotes(String(name || "").trim());
  if (!font) return officeFonts[0] || "Microsoft YaHei";

  const lower = font.toLowerCase();

  const aliasMap = {
    // Generic families
    "system-ui": officeFonts[0] || "Microsoft YaHei",
    "sans-serif": officeFonts[0] || "Microsoft YaHei",
    serif: "SimSun",
    "ui-sans-serif": officeFonts[0] || "Microsoft YaHei",

    // macOS 中文字体 -> Office 等价字体
    "pingfang sc": "Microsoft YaHei",
    "pingfang tc": "Microsoft JhengHei",
    ".pingfang sc": "Microsoft YaHei",
    "hiragino sans gb": "Microsoft YaHei",
    "hiragino sans": "Meiryo",
    "heiti sc": "SimHei",
    "heiti tc": "Microsoft JhengHei",
    "songti sc": "SimSun",
    "songti tc": "PMingLiU",
    "stheiti": "SimHei",
    "stsong": "SimSun",
    "stkaiti": "KaiTi",
    "stfangsong": "FangSong",
    "apple lihei pro": "Microsoft YaHei",

    // Windows 别名 / 变体
    "microsoft yahei ui": "Microsoft YaHei",
    "ms yahei": "Microsoft YaHei",
    "ms gothic": "MS Gothic",
    "ms mincho": "MS Mincho",
    "yu gothic": "Meiryo",
    "yu gothic ui": "Meiryo",
    "yu mincho": "MS Mincho",

    // Linux CJK 字体
    "noto sans cjk sc": "Microsoft YaHei",
    "noto sans cjk tc": "Microsoft JhengHei",
    "noto sans cjk jp": "Meiryo",
    "noto sans cjk kr": "Malgun Gothic",
    "noto serif cjk sc": "SimSun",
    "source han sans sc": "Microsoft YaHei",
    "source han sans tc": "Microsoft JhengHei",
    "source han serif sc": "SimSun",
    "wenquanyi micro hei": "Microsoft YaHei",
    "wenquanyi zen hei": "Microsoft YaHei",
    "wenquanyi micro hei mono": "Microsoft YaHei",

    // 韩文
    "apple sd gothic neo": "Malgun Gothic",
    "nanum gothic": "Malgun Gothic",
    "nanumgothic": "Malgun Gothic",
    "batang": "Batang",
    "gulim": "Malgun Gothic",
  };

  if (aliasMap[lower]) {
    return aliasMap[lower];
  }

  if (!preferOfficeFonts) {
    return font;
  }

  return font;
}

/**
 * Complex Script 字体规范化。
 * 这里尽量保守：
 * - generic family 映射到更稳妥的字体
 * - 其他保留原名，避免误伤
 */
function normalizeToPPTCSFont(name, officeFonts, preferOfficeFonts = true) {
  const font = stripQuotes(String(name || "").trim());
  if (!font) return officeFonts[0] || "Arial";

  const lower = font.toLowerCase();

  const aliasMap = {
    "system-ui": officeFonts[0] || "Arial",
    "sans-serif": officeFonts[0] || "Arial",
    serif: "Times New Roman",
    "ui-sans-serif": officeFonts[0] || "Arial",
    "ui-serif": "Times New Roman",
  };

  if (aliasMap[lower]) {
    return aliasMap[lower];
  }

  if (!preferOfficeFonts) {
    return font;
  }

  return font;
}

/**
 * Symbol 字体规范化。
 * 重点照顾：
 * - generic family
 * - math / symbol 相关字体的稳定表达
 */
function normalizeToPPTSymbolFont(name, officeFonts, preferOfficeFonts = true) {
  const font = stripQuotes(String(name || "").trim());
  if (!font) return officeFonts[0] || "Cambria Math";

  const lower = font.toLowerCase();

  const aliasMap = {
    "system-ui": "Segoe UI Symbol",
    "sans-serif": "Segoe UI Symbol",
    serif: "Times New Roman",
    "cambria math": "Cambria Math",
    "segoe ui symbol": "Segoe UI Symbol",
    "stix two math": "STIX Two Math",
    symbol: "Symbol",
    "mt extra": "MT Extra",
    "arial unicode ms": "Arial Unicode MS",
  };

  if (aliasMap[lower]) {
    return aliasMap[lower];
  }

  if (!preferOfficeFonts) {
    return font;
  }

  return font;
}

/* =========================================================
 * 模块导出
 * ========================================================= */

// UMD：支持 CommonJS / AMD / 浏览器全局
if (typeof module !== "undefined" && module.exports) {
  module.exports = detectFontsForPPTX;
} else if (typeof define === "function" && define.amd) {
  define(function () {
    return detectFontsForPPTX;
  });
} else if (typeof window !== "undefined") {
  window.detectFontsForPPTX = detectFontsForPPTX;
}
