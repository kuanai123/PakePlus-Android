/**
 * parser.js（浏览器版，由 parser-core.js 自动转换，业务逻辑一致） — 题库文档解析核心（Excel / Word → 统一题目结构）
 *
 * 独立于微信云 SDK，可在 Node 环境直接 require 进行测试。
 * 统一题目结构：
 * {
 *   type: 'single' | 'multi' | 'judge' | 'fill' | 'qa',
 *   stem: 题干（保留 <--NRC--> 填空标记）,
 *   options: [{key:'A', text:'...'}],   // 单选/多选
 *   answer: 'A' | ['A','C'] | 'Y'|'N' | '文本' | ['词1','词2'] | 长文本,
 *   answerRule: 'exact' | 'set' | 'yn' | 'text' | 'multiword' | 'self',
 *   analysis, chapter, tags: [],
 *   sourceRow,          // 原始行号（Excel）或段序号（Word）
 *   problem: null | 'no-stem' | 'no-answer' | 'few-options' | 'bad-answer'
 * }
 */

var XLSX = window.XLSX;


var FILL_TOKEN_RE = /<--NRC-->/g;

var TYPE_MAP = {
  '单选题': 'single', '单选': 'single', '单项选择题': 'single',
  '多选题': 'multi', '多选': 'multi', '多项选择题': 'multi',
  '判断题': 'judge', '判断': 'judge', '判断正误题': 'judge',
  '填空题': 'fill', '填空': 'fill',
  '简答题': 'qa', '简答': 'qa', '问答': 'qa', '问答题': 'qa', '论述题': 'qa'
};

var COLUMN_ALIASES = {
  serial: ['序号', '题号', '编号', 'no', 'no.'],
  type: ['试题类型', '题型', '类型', '题目类型', '试题种类'],
  stem: ['试题内容', '题干', '题目', '题目内容', '试题', '题面'],
  optA: ['供选项a', '选项a', '选项 a', 'a项', 'a'],
  optB: ['供选项b', '选项b', '选项 b', 'b项', 'b'],
  optC: ['供选项c', '选项c', '选项 c', 'c项', 'c'],
  optD: ['供选项d', '选项d', '选项 d', 'd项', 'd'],
  answer: ['标准答案', '答案', '正确答案', '参考答案'],
  tags: ['标签', '分类', '专业分类'],
  chapter: ['出处条文', '出处', '章节', '知识点', '来源', '条文'],
  analysis: ['试题解析', '解析', '答案解析', '题目解析', '解释']
};

var JUDGE_TRUE = ['y', '对', '正确', '√', 't', 'true', '是', 'yes'];
var JUDGE_FALSE = ['n', '错', '错误', '×', 'x', 'f', 'false', '否', 'no'];

/* ---------- 基础工具 ---------- */

function normText(v) {
  if (v === undefined || v === null) return '';
  return String(v).trim();
}

function normKey(v) {
  return normText(v)
    .toLowerCase()
    .replace(/[\s()（）\[\]【】]/g, '');
}

/** 判断题答案归一为 'Y' / 'N'，无法识别返回 '' */
function normalizeJudgeAnswer(raw) {
  var k = normKey(raw);
  if (!k) return '';
  if (JUDGE_TRUE.indexOf(k) >= 0) return 'Y';
  if (JUDGE_FALSE.indexOf(k) >= 0) return 'N';
  return '';
}

/** 多选答案 'ABC' / 'A、C' / 'A C' → ['A','C'] */
function normalizeMultiAnswer(raw) {
  var s = normText(raw).toUpperCase().replace(/[^A-D]/g, '');
  var out = [];
  for (var i = 0; i < s.length; i++) {
    if (out.indexOf(s[i]) < 0) out.push(s[i]);
  }
  return out.sort();
}

function stemHash(stem) {
  var s = normText(stem).replace(/\s+/g, '');
  var h = 5381;
  for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0;
  return 'h' + h.toString(16);
}

function hasAnyOption(options) {
  return !!options && ['A', 'B', 'C', 'D'].some(function (k) {
    return normText(options[k]);
  });
}

function countBlanks(stem) {
  var m = normText(stem).match(FILL_TOKEN_RE);
  return m ? m.length : 0;
}

/* ---------- 通用题目组装 ---------- */

/**
 * raw: { type?, stem, options{}, answerRaw, analysis, chapter, tags[], sourceRow }
 * 返回统一题目结构；识别不了答案/题干时打 problem 标记（默认不导入）
 */
function buildQuestion(raw) {
  var q = {
    type: raw.type || '',
    stem: normText(raw.stem),
    options: [],
    answer: '',
    answerRule: 'text',
    analysis: normText(raw.analysis),
    chapter: normText(raw.chapter),
    tags: (raw.tags || []).filter(Boolean),
    sourceRow: raw.sourceRow || 0,
    problem: null
  };

  // 题型推断（无类型列或类型为空时按形状判断）
  if (!q.type) {
    q.type = inferType(q.stem, raw.options, raw.answerRaw);
  }

  // 显式题型与答案形状矛盾时信任答案（Word 分组标题可能漏写/写错）：
  // 声明单选/多选但没有选项行，且答案是 Y/N → 判错题；含填空标记 → 填空题
  if ((q.type === 'single' || q.type === 'multi') && !hasAnyOption(raw.options)) {
    if (normalizeJudgeAnswer(raw.answerRaw === undefined || raw.answerRaw === null ? '' : raw.answerRaw)) {
      q.type = 'judge';
    } else if (countBlanks(q.stem) > 0) {
      q.type = 'fill';
    }
  }

  // 选项组装（单选/多选）
  if (q.type === 'single' || q.type === 'multi') {
    ['A', 'B', 'C', 'D'].forEach(function (key) {
      var text = normText(raw.options && raw.options[key]);
      // 兼容"0 占位"：删除清单里无选项单元格写 0，视为空
      if (text === '0' || text === '0.0') text = '';
      if (text) q.options.push({ key: key, text: text });
    });
  }

  // 答案按题型分支规范化
  var ansRaw = raw.answerRaw;
  if (ansRaw === undefined || ansRaw === null || normText(ansRaw) === '') {
    q.problem = 'no-answer';
    return q;
  }
  switch (q.type) {
    case 'single':
      var a = normText(ansRaw).toUpperCase();
      if (/^[A-D]$/.test(a)) {
        q.answer = a;
        q.answerRule = 'exact';
      } else {
        q.problem = 'bad-answer';
      }
      break;
    case 'multi':
      var m = normalizeMultiAnswer(ansRaw);
      if (m.length >= 2) {
        q.answer = m;
        q.answerRule = 'set';
      } else if (m.length === 1) {
        // 退化：只选一个的多选，仍按集合处理
        q.answer = m;
        q.answerRule = 'set';
      } else {
        q.problem = 'bad-answer';
      }
      break;
    case 'judge':
      var j = normalizeJudgeAnswer(ansRaw);
      if (j) {
        q.answer = j;
        q.answerRule = 'yn';
      } else {
        q.problem = 'bad-answer';
      }
      break;
    case 'fill':
      q.answer = fillAnswers(ansRaw, q.stem);
      q.answerRule = 'multiword';
      break;
    default: // qa 及其他
      q.answer = normText(ansRaw);
      q.answerRule = 'self';
      break;
  }

  // 选项数量校验
  if ((q.type === 'single' || q.type === 'multi') && q.options.length < 2 && !q.problem) {
    q.problem = 'few-options';
  }
  return q;
}

/** 填空答案：单空整串；多空按常见分隔符拆成数组 */
function fillAnswers(ansRaw, stem) {
  var text = normText(ansRaw);
  var blanks = countBlanks(stem);
  if (blanks >= 2 && /[,;；、|\s]/.test(text)) {
    var parts = text.split(/[,;；、|\s]+/).filter(Boolean);
    if (parts.length === blanks) return parts;
  }
  return [text];
}

/** 无题型列时按内容形状推断 */
function inferType(stem, options, answerRaw) {
  var hasOptions = options && ['A', 'B', 'C', 'D'].some(function (k) {
    return normText(options[k]);
  });
  if (hasOptions) {
    var m = normalizeMultiAnswer(answerRaw);
    return m.length >= 2 ? 'multi' : 'single';
  }
  var j = normalizeJudgeAnswer(answerRaw);
  if (j && normText(stem).length > 8) return 'judge';
  if (countBlanks(stem) > 0) return 'fill';
  return 'qa';
}

/* ---------- Excel 解析 ---------- */

function isEmptyCell(v) {
  return v === undefined || v === null || String(v).trim() === '';
}

/** 在前 10 行内寻找表头行（匹配 >=3 个已知列别名） */
function findHeaderRow(rows) {
  for (var i = 0; i < Math.min(rows.length, 10); i++) {
    var row = rows[i] || [];
    var hits = 0;
    for (var c = 0; c < row.length; c++) {
      if (matchColumn(normKey(row[c]))) hits++;
    }
    if (hits >= 3) return i;
  }
  return -1;
}

function matchColumn(key) {
  if (!key) return null;
  var names = Object.keys(COLUMN_ALIASES);
  for (var i = 0; i < names.length; i++) {
    if (COLUMN_ALIASES[names[i]].indexOf(key) >= 0) return names[i];
  }
  return null;
}

/** 表头 → 字段列索引映射 */
function mapColumns(headerRow) {
  var map = {};
  (headerRow || []).forEach(function (cell, idx) {
    var field = matchColumn(normKey(cell));
    if (field && map[field] === undefined) map[field] = idx;
  });
  return map;
}

/**
 * 计算 sheet 的真实数据范围，防御格式残留导致的超大 !ref（如 A1:XFD2057 整表宽）
 * 不裁剪的话 sheet_to_json 会按全宽逐格处理，2056 行能拖到 15 秒
 */
function realRange(ws) {
  var ref = ws && ws['!ref'];
  if (!ref) return null;
  var range = XLSX.utils.decode_range(ref);
  var maxC = range.s.c;
  var maxR = range.s.r;
  var re = /^([A-Z]+)(\d+)$/;
  Object.keys(ws).forEach(function (k) {
    var m = k.match(re);
    if (!m) return;
    var c = XLSX.utils.decode_col(m[1]);
    var r = parseInt(m[2], 10) - 1;
    if (c > maxC) maxC = c;
    if (r > maxR) maxR = r;
  });
  // 留 2 列余量，避免边界数据被裁掉
  return { s: { r: range.s.r, c: range.s.c }, e: { r: maxR, c: maxC + 2 } };
}

function parseExcelBuffer(buffer) {
  var wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: false });
  return wb.SheetNames.map(function (name) {
    var ws = wb.Sheets[name];
    var rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: '',
      raw: true,
      blankrows: false,
      range: realRange(ws) || 'A1'
    });
    var headerIdx = findHeaderRow(rows);
    if (headerIdx < 0) {
      return { name: name, skipped: true, reason: '未识别到表头', items: [], stats: emptyStats() };
    }
    var colMap = mapColumns(rows[headerIdx]);
    // 岗位适用列（表头含"值班员"且值为 √ 的列）
    var postCols = [];
    (rows[headerIdx] || []).forEach(function (cell, idx) {
      var t = normText(cell);
      if (/值班员/.test(t)) postCols.push({ idx: idx, name: t });
    });

    var items = [];
    for (var r = headerIdx + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var get = function (field) {
        var ci = colMap[field];
        return ci === undefined ? '' : row[ci];
      };
      if (!colMap.stem || isEmptyCell(get('stem'))) continue; // 无题干列或空行
      var options = {};
      ['A', 'B', 'C', 'D'].forEach(function (k) {
        options[k] = get('opt' + k);
      });
      // 单列"选项"兼容：单元格内多行 "A.xxx\nB.xxx"
      if (!options.A && colMap.optA === undefined) {
        var merged = get('stem') && findMergedOptions(row, colMap);
        if (merged) options = merged;
      }
      var tags = normText(get('tags'))
        .split(/[,;；、\s]+/)
        .filter(Boolean);
      postCols.forEach(function (pc) {
        if (normText(row[pc.idx]).indexOf('√') >= 0) tags.push(pc.name);
      });
      var item = buildQuestion({
        type: TYPE_MAP[normKey(get('type'))] || '',
        stem: get('stem'),
        options: options,
        answerRaw: get('answer'),
        analysis: get('analysis') === 0 || get('analysis') === '0' ? '' : get('analysis'),
        chapter: get('chapter'),
        tags: tags,
        sourceRow: r + 1
      });
      if (!item.stem) item.problem = item.problem || 'no-stem';
      items.push(item);
    }
    return {
      name: name,
      skipped: false,
      items: items,
      stats: calcStats(items)
    };
  });
}

function emptyStats() {
  return { total: 0, ok: 0, byProblem: {} };
}

function calcStats(items) {
  var st = emptyStats();
  st.total = items.length;
  items.forEach(function (it) {
    if (it.problem) {
      st.byProblem[it.problem] = (st.byProblem[it.problem] || 0) + 1;
    } else {
      st.ok++;
    }
  });
  return st;
}

/* ---------- Word(docx) 解析 ---------- */

var TYPE_HEADER_RE = /^[【\[(（]?\s*(单选题|多选题|判断题|填空题|简答题|单项选择题|多项选择题|判断说明题)\s*[】\])）]?/;
var ANSWER_LINE_RE = /^(标准答案|正确答案|参考答案|答案)\s*[:：]?\s*(.+)$/;
var ANALYSIS_LINE_RE = /^(试题解析|答案解析|解析|解释)\s*[:：]?\s*(.+)$/;
var OPTION_LINE_RE = /^([A-Da-d])\s*[\.、．:：)）]\s*(.+)$/;
var SERIAL_LINE_RE = /^(\d{1,3})\s*[\.、．]\s*(.+)$/;
var GROUP_HEADER_RE = /^[一二三四五六七八九十]+\s*[、．.]\s*(单选题|多选题|判断题|填空题|简答题|单项选择题|多项选择题)/;

function parseMarkdownQuestions(md) {
  var lines = String(md || '')
    .split(/\r?\n/)
    .map(function (l) {
      // 兜底去 markdown 转义（convertToMarkdown 路径会残留 1\. A\. <\-\-NRC\-\-> 之类）
      return l.trim().replace(/\\([.!<>\\`*_{}\[\]()#+\-|~])/g, '$1');
    })
    .filter(function (l) { return l.length > 0; });

  var items = [];
  var cur = null;
  var typeHint = '';

  function newQ(type, stemStart) {
    return { type: type || '', stemLines: stemStart ? [stemStart] : [], options: {}, answerRaw: '', analysisLines: [] };
  }
  function settle() {
    if (!cur) return;
    var item = buildQuestion({
      type: cur.type || typeHint || '',
      stem: cur.stemLines.join(''),
      options: cur.options,
      answerRaw: cur.answerRaw,
      analysis: cur.analysisLines.join('\n'),
      chapter: '',
      tags: [],
      sourceRow: items.length + 1
    });
    if (!item.stem) item.problem = 'no-stem';
    items.push(item);
    cur = null;
  }

  lines.forEach(function (line) {
    // 1. 独立题型标记行
    var tm = line.match(TYPE_HEADER_RE);
    if (tm && line.replace(TYPE_HEADER_RE, '').trim().length <= 2) {
      settle();
      typeHint = TYPE_MAP[tm[1]] || '';
      return;
    }
    // 2. 答案行（结算当前题）
    var am = line.match(ANSWER_LINE_RE);
    if (am) {
      if (cur) {
        cur.answerRaw = am[2].trim();
        settle();
      }
      return;
    }
    // 3. 解析行
    var pm = line.match(ANALYSIS_LINE_RE);
    if (pm && cur) {
      cur.analysisLines.push(pm[2].trim());
      return;
    }
    // 4. 选项行
    var om = line.match(OPTION_LINE_RE);
    if (om && cur && cur.stemLines.length > 0) {
      cur.options[om[1].toUpperCase()] = om[2].trim();
      return;
    }
    // 5. 题号行（已有内容则先结算）
    var sm = line.match(SERIAL_LINE_RE);
    if (sm) {
      if (cur && (cur.stemLines.length || cur.answerRaw)) settle();
      cur = newQ(typeHint, sm[2].trim());
      return;
    }
    // 6. 分组标题 "一、单选题"
    var hm = line.match(GROUP_HEADER_RE);
    if (hm) {
      settle();
      typeHint = TYPE_MAP[hm[1]] || '';
      return;
    }
    // 7. 题干内容累积（题型标记与题干同行时剥离标记）
    if (!cur) cur = newQ(typeHint, '');
    var stemLine = tm ? line.replace(TYPE_HEADER_RE, '').trim() : line;
    if (stemLine) cur.stemLines.push(stemLine);
  });
  settle();
  return items;
}

/* ---------- 入口 ---------- */

function extOf(fileName) {
  var m = String(fileName || '').toLowerCase().match(/\.([a-z0-9]+)$/);
  return m ? m[1] : '';
}

/**
 * 解析任意支持的格式，返回：
 * { fileType: 'xlsx'|'xls'|'docx', sheets: [{name, skipped, items, stats}] }
 * 或 { fileType: 'doc'|'unknown', error: '...' }
 */
async function parseAny(buffer, fileName) {
  var ext = extOf(fileName);

  if (ext === 'doc') {
    return { fileType: 'doc', error: 'LEGACY_DOC' };
  }

  if (ext === 'xls' || ext === 'xlsx') {
    try {
      return { fileType: ext, sheets: parseExcelBuffer(buffer) };
    } catch (e) {
      return { fileType: ext, error: '表格解析失败：' + e.message };
    }
  }

  if (ext === 'docx') {
    try {
      var mammoth = window.mammoth;
      // extractRawText 输出纯文本；convertToMarkdown 会把 1. A. <--NRC--> 转义成 1\. A\. 导致匹配失败
      var result = await mammoth.extractRawText({ arrayBuffer: buffer });
      var items = parseMarkdownQuestions(result.value);
      return {
        fileType: 'docx',
        sheets: [{ name: 'Word 文档', skipped: false, items: items, stats: calcStats(items) }]
      };
    } catch (e) {
      return { fileType: 'docx', error: 'Word 解析失败：' + e.message };
    }
  }

  // 扩展名缺失或不符：按内容试探（先表格后 Word）
  try {
    return { fileType: 'xlsx', sheets: parseExcelBuffer(buffer) };
  } catch (e) { /* not a spreadsheet */ }
  try {
    var mammoth2 = window.mammoth;
    var r2 = await mammoth2.extractRawText({ arrayBuffer: buffer });
    var it2 = parseMarkdownQuestions(r2.value);
    return { fileType: 'docx', sheets: [{ name: 'Word 文档', skipped: false, items: it2, stats: calcStats(it2) }] };
  } catch (e2) { /* not a docx */ }
  return { fileType: 'unknown', error: '无法识别的文件格式，仅支持 xls/xlsx/docx' };
}

window.ParserCore = {
  parseAny: parseAny,
  parseExcelBuffer: parseExcelBuffer,
  parseMarkdownQuestions: parseMarkdownQuestions,
  buildQuestion: buildQuestion,
  stemHash: stemHash,
  countBlanks: countBlanks,
  normalizeJudgeAnswer: normalizeJudgeAnswer,
  normalizeMultiAnswer: normalizeMultiAnswer,
  TYPE_MAP: TYPE_MAP
};
