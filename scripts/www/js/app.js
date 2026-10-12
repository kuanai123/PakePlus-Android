/**
 * app.js — 刷题题库（浏览器/安卓版）
 * 全部数据存本地 IndexedDB，无服务器、可离线，安卓浏览器「添加到主屏幕」后像原生 App
 */
(function () {
  var P = window.ParserCore;
  var J = window.Judge;
  var DB = window.DB;

  var TYPE_NAME = { single: '单选', multi: '多选', judge: '判断', fill: '填空', qa: '简答' };
  var TYPE_ORDER = ['single', 'multi', 'judge', 'fill', 'qa'];

  var state = {
    view: 'import',
    parsed: null,
    sheetIdx: 0,
    previewItems: [],
    questions: [],
    wrongs: [],
    batches: [],
    favorites: [],
    session: null,
    editing: null
  };

  /* ---------- 工具 ---------- */
  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s === undefined || s === null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function stemDisplay(stem) {
    return String(stem || '').replace(/<--NRC-->/g, '______');
  }
  function countBlanks(stem) {
    var m = String(stem || '').match(/<--NRC-->/g);
    return m ? m.length : 0;
  }
  function answerText(q) {
    var a = q.answer;
    if (a === undefined || a === null || a === '') return '（未设置）';
    if (q.type === 'judge') return a === 'Y' ? '对' : '错';
    if (Array.isArray(a)) return a.join(q.type === 'multi' ? '' : ' / ');
    return String(a);
  }
  /** Toast：底部居中，2.2 秒自动消失 */
  var toastTimer = null;
  function showToast(msg) {
    var el = $('toast');
    if (!el) return;
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.classList.remove('show'); }, 2200);
  }
  function toast(msg) {
    showToast(msg);
    var s = $('parseStatus');
    if (s && state.view === 'import') s.textContent = msg;
  }
  function newId() { return DB.newId(); }

  /* ---------- 视图切换 ---------- */
  var TITLES = { import: '导入题库', bank: '题库管理', practice: '练习', wrong: '错题本', search: '答案检索' };
  function switchView(v) {
    state.view = v;
    ['import', 'bank', 'practice', 'wrong', 'search'].forEach(function (name) {
      $('view-' + name).hidden = name !== v;
    });
    Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (t) {
      t.classList.toggle('active', t.dataset.view === v);
    });
    $('viewTitle').textContent = TITLES[v];
    if (v === 'bank') renderBank();
    if (v === 'wrong') renderWrongs();
    if (v === 'import') renderBatches();
  }

  /* ---------- 数据加载 ---------- */
  function loadAll() {
    return Promise.all([
      DB.getAll('questions'),
      DB.getAll('wrongs'),
      DB.getAll('batches'),
      DB.getAll('favorites')
    ]).then(function (r) {
      state.questions = r[0].filter(function (q) { return !q.deleted; });
      state.wrongs = r[1];
      state.batches = r[2].filter(function (b) { return !b.deleted; });
      state.favorites = r[3] || [];
      renderStats();
    });
  }
  function renderStats() {
    var due = state.wrongs.filter(function (w) {
      return !w.mastered && new Date(w.nextReviewAt).getTime() <= Date.now();
    }).length;
    var active = state.wrongs.filter(function (w) { return !w.mastered; }).length;
    $('statPill').textContent = '题库 ' + state.questions.length + ' · 错题 ' + active +
      ' · 待复习 ' + due + ' · 收藏 ' + (state.favorites || []).length;
    var bsc = $('bankSubCount');
    if (bsc) bsc.textContent = state.questions.length + ' 题';
  }

  /* ---------- 导入 ---------- */
  var MAX_FILE_SIZE = 20 * 1024 * 1024; // 20MB 上限，防止超大文件拖垮浏览器

  function handleFile(file) {
    if (!file) return;
    if (/\.doc$/i.test(file.name)) { toast('不支持 .doc 旧格式，请用 Word/WPS 另存为 .docx'); return; }
    if (!/\.(xlsx|xls|docx)$/i.test(file.name)) { toast('仅支持 xlsx / xls / docx 格式'); return; }
    if (file.size > MAX_FILE_SIZE) {
      toast('文件过大（' + Math.round(file.size / 1048576) + 'MB），请拆分后再导入');
      return;
    }
    $('parseStatus').textContent = '解析中：' + file.name;
    parseFile(file);
  }

  /** 优先 Worker 解析（不阻塞 UI），失败自动回退主线程 */
  function parseFile(file) {
    file.arrayBuffer().then(function (buf) {
      return new Promise(function (resolve, reject) {
        if (typeof Worker === 'undefined') { reject(new Error('no-worker')); return; }
        var w;
        try { w = new Worker('js/parser.worker.js'); } catch (e) { reject(e); return; }
        var timer = setTimeout(function () {
          try { w.terminate(); } catch (e) {}
          reject(new Error('解析超时'));
        }, 60000);
        w.onmessage = function (e) {
          var d = e.data || {};
          if (d.type === 'progress') { $('parseStatus').textContent = d.text; return; }
          clearTimeout(timer);
          try { w.terminate(); } catch (err) {}
          if (d.ok) resolve(d.parsed);
          else reject(new Error(d.error || '解析失败'));
        };
        w.onerror = function () {
          clearTimeout(timer);
          try { w.terminate(); } catch (err) {}
          reject(new Error('worker error'));
        };
        w.postMessage({ buffer: buf.slice(0), fileName: file.name });
      });
    }).then(function (parsed) { onParsed(parsed, file.name); })
      .catch(function () {
        return file.arrayBuffer()
          .then(function (buf) { return P.parseAny(buf, file.name); })
          .then(function (parsed) { onParsed(parsed, file.name); });
      })
      .catch(function (err) { toast('解析出错：' + (err.message || err)); });
  }

  function onParsed(parsed, fileName) {
    if (parsed.error) {
      toast(parsed.error === 'LEGACY_DOC'
        ? '不支持 .doc 旧格式，请另存为 .docx'
        : '解析失败：' + parsed.error);
      return;
    }
    state.parsed = parsed;
    state.sheetIdx = 0;
    $('parseStatus').textContent = '解析完成：' + fileName + '（' + parsed.fileType + '）';
    renderSheets();
    loadPreview(0);
  }

  $('fileInput').addEventListener('change', function (e) {
    handleFile(e.target.files[0]);
  });

  // 拖拽导入
  (function bindDrop() {
    var zone = $('dropZone');
    if (!zone) return;
    function stop(e) { e.preventDefault(); e.stopPropagation(); }
    ['dragenter', 'dragover'].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { stop(e); zone.classList.add('dragover'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { stop(e); zone.classList.remove('dragover'); });
    });
    zone.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });
  })();

  /* ---------- 滚动加载（IntersectionObserver 哨兵） ---------- */
  /** 每个哨兵元素持有自己的 observer，避免多个列表互相顶掉 */
  function observeMore(el, onHit) {
    if (!el || typeof IntersectionObserver === 'undefined') return;
    if (el._io) { try { el._io.disconnect(); } catch (e) {} }
    var ob = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) { if (en.isIntersecting) onHit(); });
    }, { rootMargin: '200px' });
    ob.observe(el);
    el._io = ob;
  }

  function renderSheets() {
    var box = $('sheetList');
    box.innerHTML = state.parsed.sheets.map(function (s, i) {
      if (s.skipped) return '<div class="q-item muted">' + esc(s.name) + '：' + esc(s.reason) + '</div>';
      var okCount = s.stats.ok;
      return '<div class="q-item"><div class="q-head">' +
        '<span style="font-weight:600">' + esc(s.name) + '</span>' +
        '<span class="muted">' + s.stats.total + ' 条（可用 ' + okCount + '）</span>' +
        '<button class="btn btn-sm" data-sheet="' + i + '">预览并导入</button>' +
        '</div></div>';
    }).join('');
    box.querySelectorAll('[data-sheet]').forEach(function (b) {
      b.onclick = function () { loadPreview(Number(b.dataset.sheet)); };
    });
    $('sheetCard').hidden = false;
  }

  function loadPreview(idx) {
    state.sheetIdx = idx;
    var sheet = state.parsed.sheets[idx];
    state.previewItems = sheet.items.map(function (it, i) {
      return {
        key: i,
        type: it.type, stem: it.stem, options: it.options || [],
        answer: it.answer, answerRule: it.answerRule,
        analysis: it.analysis || '', chapter: it.chapter || '',
        tags: it.tags || [], sourceRow: it.sourceRow,
        problem: it.problem || '', skipped: !!it.problem
      };
    });
    $('previewCard').hidden = false;
    renderPreview();
    $('previewCard').scrollIntoView({ behavior: 'smooth' });
  }

  function renderPreview() {
    var only = $('onlyProblem').checked;
    var list = state.previewItems.filter(function (it) { return only ? !!it.problem : true; });
    var okCount = state.previewItems.filter(function (it) { return !it.problem && !it.skipped; }).length;
    $('previewSummary').textContent = '可导入 ' + okCount + ' / 共 ' + state.previewItems.length;
    $('previewList').innerHTML = list.length ? list.map(function (it) {
      var opts = (it.options || []).map(function (o) {
        return '<div>' + esc(o.key) + '. ' + esc(o.text) + '</div>';
      }).join('');
      var editOpen = state.editing === it.key;
      return '<div class="q-item">' +
        '<div class="q-head">' +
        '<span class="tag tag-' + it.type + '">' + TYPE_NAME[it.type] + '</span>' +
        (it.problem ? '<span class="tag tag-problem">' + problemText(it.problem) + '</span>' : '') +
        (it.chapter ? '<span class="muted">' + esc(it.chapter) + '</span>' : '') +
        '</div>' +
        '<div class="q-stem">' + esc(stemDisplay(it.stem)) + '</div>' +
        (opts ? '<div class="q-opts">' + opts + '</div>' : '') +
        '<div class="q-answer">答案：' + esc(answerText(it)) + '</div>' +
        (editOpen ? editForm(it) : '') +
        '<div class="q-actions">' +
        '<button class="btn btn-sm" data-edit="' + it.key + '">' + (editOpen ? '收起' : '编辑') + '</button>' +
        '<button class="btn btn-sm" data-skip="' + it.key + '">' + (it.skipped ? '恢复' : '跳过') + '</button>' +
        '</div></div>';
    }).join('') : '<div class="empty">没有符合条件的题目</div>';

    $('previewList').querySelectorAll('[data-edit]').forEach(function (b) {
      b.onclick = function () {
        var k = Number(b.dataset.edit);
        state.editing = state.editing === k ? null : k;
        renderPreview();
      };
    });
    $('previewList').querySelectorAll('[data-skip]').forEach(function (b) {
      b.onclick = function () {
        var it = state.previewItems[Number(b.dataset.skip)];
        it.skipped = !it.skipped;
        renderPreview();
      };
    });
    var saveBtn = $('previewList').querySelector('[data-save]');
    if (saveBtn) {
      saveBtn.onclick = function () {
        var it = state.previewItems[Number(saveBtn.dataset.save)];
        it.stem = $('editStem').value.trim();
        it.answer = parseAnswerInput(it.type, $('editAnswer').value);
        it.analysis = $('editAnalysis').value.trim();
        if (it.stem && it.answer !== '' && it.answer !== null) {
          it.problem = '';
          it.skipped = false;
        }
        state.editing = null;
        renderPreview();
      };
    }
  }

  function problemText(p) {
    return { 'no-stem': '无题干', 'no-answer': '无答案', 'few-options': '选项不足', 'bad-answer': '答案异常' }[p] || p;
  }

  function editForm(it) {
    return '<div class="edit-area">' +
      '<label>题干（填空空位用 &lt;--NRC--&gt;）</label>' +
      '<textarea id="editStem" rows="3">' + esc(it.stem) + '</textarea>' +
      '<label>答案</label>' +
      '<input id="editAnswer" value="' + esc(Array.isArray(it.answer) ? it.answer.join(',') : it.answer) + '">' +
      '<label>解析</label>' +
      '<textarea id="editAnalysis" rows="2">' + esc(it.analysis) + '</textarea>' +
      '<button class="btn btn-sm btn-primary" data-save="' + it.key + '">保存修改</button>' +
      '</div>';
  }

  function parseAnswerInput(type, val) {
    var v = String(val || '').trim();
    if (type === 'multi') return v.toUpperCase().replace(/[^A-D]/g, '').split('').sort();
    if (type === 'judge') return (v === 'Y' || v === '对' || v === '正确' || v === '√') ? 'Y' : 'N';
    if (type === 'fill') return /[,，;；、]/.test(v) ? v.split(/[,，;；、]+/).filter(Boolean) : [v];
    return v;
  }

  $('onlyProblem').addEventListener('change', renderPreview);

  $('previewCard').addEventListener('click', function (e) {
    if (e.target.id === 'doImport') doImport();
  });
  // 导入按钮常驻在预览卡底部
  (function addImportBtn() {
    var btn = document.createElement('button');
    btn.className = 'btn btn-primary btn-block';
    btn.id = 'doImport';
    btn.textContent = '确认导入本表';
    $('previewCard').appendChild(btn);
    btn.onclick = doImport;
  })();

  function doImport() {
    var sheet = state.parsed.sheets[state.sheetIdx];
    var items = state.previewItems.filter(function (it) { return !it.skipped && !it.problem; });
    if (!items.length) { toast('没有可导入的题目'); return; }
    var batchId = newId();
    var now = new Date();
    var batch = {
      id: batchId,
      name: (sheet.name || '题库') + ' · ' + now.toLocaleString('zh-CN'),
      fileName: sheet.name,
      importedAt: now,
      totalImported: items.length,
      deleted: false
    };
    var docs = items.map(function (it) {
      return {
        id: newId(),
        batchId: batchId,
        type: it.type,
        stem: it.stem,
        options: it.options || [],
        answer: it.answer,
        answerRule: it.answerRule,
        analysis: it.analysis || '',
        chapter: it.chapter || '',
        tags: it.tags || [],
        deleted: false,
        createdAt: now
      };
    });
    DB.bulkAdd('questions', docs)
      .then(function () { return DB.put('batches', batch); })
      .then(function () {
        state.parsed = null;
        state.previewItems = [];
        $('sheetCard').hidden = true;
        $('previewCard').hidden = true;
        toast('已导入 ' + docs.length + ' 题');
        return loadAll();
      })
      .then(function () { renderBatches(); })
      .catch(function (e) { toast('导入失败：' + e.message); });
  }

  function renderBatches() {
    $('batchList').innerHTML = state.batches.length ? state.batches.map(function (b) {
      return '<div class="q-item"><div class="q-head">' +
        '<span style="flex:1;font-weight:600">' + esc(b.name) + '</span>' +
        '<span class="muted">' + b.totalImported + ' 题</span>' +
        '<button class="btn btn-sm btn-danger" data-delbatch="' + b.id + '">删除整批</button>' +
        '</div></div>';
    }).join('') : '<div class="empty">还没有导入过题库</div>';
    $('batchList').querySelectorAll('[data-delbatch]').forEach(function (btn) {
      btn.onclick = function () { deleteBatch(btn.dataset.delbatch); };
    });
  }

  function deleteBatch(batchId) {
    if (!confirm('删除这个文档？该文档下的所有题目都会被删除。')) return Promise.resolve(false);
    DB.getAll('questions').then(function (all) {
      var updates = all.filter(function (q) { return q.batchId === batchId; })
        .map(function (q) { q.deleted = true; return q; });
      return DB.bulkAdd('questions', updates);
    }).then(function () {
      return DB.get('batches', batchId).then(function (b) {
        if (b) { b.deleted = true; return DB.put('batches', b); }
      });
    }).then(loadAll).then(function () {
      renderBatches(); renderStats();
      return true;
    });
  }

  /* ---------- 题库 ---------- */
  /** 题库管理页：只显示文档列表（名称 + 题数 + 删除），不渲染题目内容 */
  function renderBank() {
    var list = state.batches;
    $('bankList').innerHTML = list.length ? list.map(function (b) {
      var cnt = state.questions.filter(function (q) { return q.batchId === b.id; }).length;
      return '<div class="q-item"><div class="q-head">' +
        '<span class="doc-name">' + esc(b.name) + '</span>' +
        '<span class="muted">' + cnt + ' 题</span>' +
        '<button class="btn btn-sm btn-danger" data-delb="' + b.id + '">删除</button>' +
        '</div></div>';
    }).join('') : '<div class="empty">还没有导入任何文档，先去导入题库</div>';

    // 删除：二次确认（deleteBatch 内）+ 即时刷新列表
    $('bankList').querySelectorAll('[data-delb]').forEach(function (btn) {
      btn.onclick = function () {
        deleteBatch(btn.dataset.delb).then(function (ok) {
          if (ok) renderBank();
        });
      };
    });
  }
  // 题型筛选已移除：题库管理页改为文档列表，不再需要 bankType
  $('exportBank').addEventListener('click', function () { exportData('questions'); });

  /* ---------- 导出 ---------- */
  function exportData(scope) {
    var rows, name;
    if (scope === 'wrongs') {
      rows = [['题型', '章节', '题干', '正确答案', '错误次数', '掌握状态', '记忆阶段']];
      state.wrongs.forEach(function (w) {
        rows.push([
          TYPE_NAME[w.snapshot.type] || w.snapshot.type,
          w.snapshot.chapter || '',
          w.snapshot.stem,
          answerText(w.snapshot),
          w.wrongCount,
          w.mastered ? '已掌握' : '未掌握',
          (w.stage || 0) + '/3'
        ]);
      });
      name = '错题本';
    } else {
      rows = [['题型', '章节', '题干', '选项A', '选项B', '选项C', '选项D', '答案', '解析', '标签']];
      state.questions.forEach(function (q) {
        var opt = function (k) { return ((q.options || []).find(function (o) { return o.key === k; }) || {}).text || ''; };
        rows.push([
          TYPE_NAME[q.type] || q.type, q.chapter || '', q.stem,
          opt('A'), opt('B'), opt('C'), opt('D'),
          answerText(q), q.analysis || '', (q.tags || []).join('、')
        ]);
      });
      name = '题库';
    }
    if (rows.length <= 1) { toast('没有可导出的数据'); return; }
    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 8 }, { wch: 18 }, { wch: 60 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 40 }, { wch: 16 }];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, name);
    XLSX.writeFile(wb, name + '-' + new Date().toISOString().slice(0, 10) + '.xlsx');
  }

  /* ---------- 练习（考试宝式：模式入口 + 答题卡 + 收藏） ---------- */
  var practiceType = '';
  var practiceChapter = '';

  function showPanel(id) {
    ['typePanel', 'chapterPanel', 'wrongPanel'].forEach(function (p) {
      var el = $(p);
      if (el) el.hidden = p !== id;
    });
  }

  // 模式入口卡片
  Array.prototype.forEach.call(document.querySelectorAll('.mode-card'), function (btn) {
    btn.onclick = function () {
      var mode = btn.dataset.mode;
      if (mode === 'type') { showPanel('typePanel'); renderTypePick(); return; }
      if (mode === 'chapter') { showPanel('chapterPanel'); renderChapters(); return; }
      if (mode === 'wrong') { showPanel('wrongPanel'); renderWrongPanel(); return; }
      startPractice(mode);
    };
  });

  function renderTypePick() {
    $('typePick').innerHTML = TYPE_ORDER.map(function (t) {
      return '<button class="type-tab' + (practiceType === t ? ' active' : '') +
        '" data-ptype="' + t + '">' + TYPE_NAME[t] + '</button>';
    }).join('');
    $('typePick').querySelectorAll('[data-ptype]').forEach(function (b) {
      b.onclick = function () { practiceType = b.dataset.ptype; renderTypePick(); };
    });
  }

  function renderChapters() {
    var map = {};
    state.questions.forEach(function (q) {
      var c = q.chapter || '未分类';
      map[c] = (map[c] || 0) + 1;
    });
    var keys = Object.keys(map).sort(function (a, b) { return map[b] - map[a]; });
    $('chapterList').innerHTML = keys.length ? keys.map(function (c) {
      return '<div class="q-item"><div class="q-head">' +
        '<span style="flex:1">' + esc(c) + '</span>' +
        '<span class="muted">' + map[c] + ' 题</span>' +
        '<button class="btn btn-sm" data-chapter="' + esc(c) + '">练习</button></div></div>';
    }).join('') : '<div class="empty">题库没有章节信息</div>';
    $('chapterList').querySelectorAll('[data-chapter]').forEach(function (b) {
      b.onclick = function () { startPractice('chapter', { chapter: b.dataset.chapter }); };
    });
  }

  $('startTypeBtn').addEventListener('click', function () {
    if (!practiceType) { toast('请选择题型'); return; }
    startPractice('type', { type: practiceType });
  });

  /* ---------- 进度持久化（考试宝式：顺序/题型/章节练习自动记录进度） ---------- */
  var PROG_PREFIX = 'stk_prog_';
  function sessionKey(mode, opts) {
    return mode + '|' + ((opts && (opts.type || opts.chapter)) || '');
  }
  function loadProgress(key) {
    try {
      var raw = localStorage.getItem(PROG_PREFIX + key);
      return raw ? JSON.parse(raw) : { idx: 0, answered: {} };
    } catch (e) { return { idx: 0, answered: {} }; }
  }
  function saveProgress(key, prog) {
    try { localStorage.setItem(PROG_PREFIX + key, JSON.stringify(prog)); } catch (e) {}
  }
  function clearProgress(key) {
    try { localStorage.removeItem(PROG_PREFIX + key); } catch (e) {}
  }
  function isProgressMode(mode) {
    return mode === 'normal' || mode === 'type' || mode === 'chapter';
  }

  /** 模式入口卡片上的进度副标题 */
  function renderModeProgress() {
    function setSub(mode, text) {
      var el = document.querySelector('.mode-card[data-mode="' + mode + '"] .mode-sub');
      if (el) el.textContent = text;
    }
    var total = state.questions.length;
    var np = loadProgress(sessionKey('normal', {}));
    var doneCount = Object.keys(np.answered || {}).length;
    setSub('normal', doneCount + '/' + total);
    var lastIdx = np.idx || 0;
    if (doneCount > 0 && lastIdx > 0) setSub('normal', '已做 ' + doneCount + '/' + total);
    var wCount = state.wrongs.filter(function (w) { return !w.mastered; }).length;
    setSub('wrong', wCount + ' 题');
    setSub('fav', state.favorites.length + ' 题');
    setSub('chapter', total > 0 ? '按出处分组' : '');
    setSub('random', '打乱顺序');
    setSub('type', '按题型刷');
  }

  /* ---------- 错题重做面板（按题型细分 + 统计 + 清理已掌握） ---------- */
  function renderWrongPanel() {
    var active = state.wrongs.filter(function (w) { return !w.mastered; });
    var mastered = state.wrongs.filter(function (w) { return w.mastered; });
    $('wsActive').textContent = active.length;
    $('wsMastered').textContent = mastered.length;

    // 正确率：来自练习记录
    DB.getAll('records').then(function (recs) {
      var total = recs.length;
      var right = recs.filter(function (r) { return r.correct; }).length;
      $('wsRate').textContent = total ? Math.round(right / total * 100) + '%' : '--';
    }).catch(function () {});

    // 按题型分组计数
    var map = {};
    active.forEach(function (w) {
      var t = (w.snapshot && w.snapshot.type) || 'qa';
      map[t] = (map[t] || 0) + 1;
    });
    var types = TYPE_ORDER.filter(function (t) { return map[t]; });
    $('wrongTypeList').innerHTML = types.length ? types.map(function (t) {
      return '<div class="q-item"><div class="q-head">' +
        '<span class="tag tag-' + t + '">' + TYPE_NAME[t] + '</span>' +
        '<span class="muted">' + map[t] + ' 题</span>' +
        '<button class="btn btn-sm" data-wtype="' + t + '">重做</button></div></div>';
    }).join('') : '<div class="empty">太棒了，没有未掌握错题</div>';
    $('wrongTypeList').querySelectorAll('[data-wtype]').forEach(function (b) {
      b.onclick = function () { startPractice('wrong', { type: b.dataset.wtype }); };
    });

    // 已掌握列表（可移除）
    $('masteredList').innerHTML = mastered.length ? mastered.slice(0, 50).map(function (w) {
      return '<div class="q-item"><div class="q-head">' +
        '<span class="tag tag-' + ((w.snapshot && w.snapshot.type) || 'qa') + '">' +
        TYPE_NAME[(w.snapshot && w.snapshot.type) || 'qa'] + '</span>' +
        '<span style="flex:1" class="q-stem-mini">' + esc(stemDisplay(w.snapshot && w.snapshot.stem)) + '</span>' +
        '<button class="btn btn-sm btn-danger" data-rmid="' + w.questionId + '">移除</button>' +
        '</div></div>';
    }).join('') : '<div class="empty">还没有已掌握的错题</div>';
    $('masteredList').querySelectorAll('[data-rmid]').forEach(function (b) {
      b.onclick = function () {
        DB.del('wrongs', b.dataset.rmid)
          .then(loadAll)
          .then(function () { renderWrongPanel(); renderStats(); })
          .catch(function (e) { toast('移除失败：' + e.message); });
      };
    });
  }

  $('redoAllBtn').addEventListener('click', function () {
    startPractice('wrong');
  });

  $('clearMasteredBtn').addEventListener('click', function () {
    var mastered = state.wrongs.filter(function (w) { return w.mastered; });
    if (!mastered.length) { toast('没有已掌握错题'); return; }
    if (!confirm('清空 ' + mastered.length + ' 条已掌握错题？')) return;
    Promise.all(mastered.map(function (w) { return DB.del('wrongs', w.questionId); }))
      .then(loadAll)
      .then(function () { renderWrongPanel(); renderStats(); toast('已清空'); })
      .catch(function (e) { toast('操作失败：' + e.message); });
  });

  /* ---------- 返回按钮 ---------- */
  $('quizBack').addEventListener('click', function () {
    $('quizCard').hidden = true;
    $('setupCard').hidden = false;
    state.session = null;
    renderModeProgress();
    // 若当前停在错题面板，同步刷新统计
    var wp = $('wrongPanel');
    if (wp && !wp.hidden) renderWrongPanel();
  });

  $('gotoBank').addEventListener('click', function () {
    switchView('bank');
  });

  $('bankBack').addEventListener('click', function () {
    switchView('import');
  });

  function startPractice(mode, opts) {
    opts = opts || {};
    var pool = state.questions.slice();
    if (opts.type) pool = pool.filter(function (q) { return q.type === opts.type; });
    if (opts.chapter) pool = pool.filter(function (q) { return (q.chapter || '未分类') === opts.chapter; });
    if (mode === 'wrong') {
      var ids = state.wrongs.filter(function (w) { return !w.mastered; }).map(function (w) { return w.questionId; });
      pool = pool.filter(function (q) { return ids.indexOf(q.id) >= 0; });
    }
    if (mode === 'fav') {
      var fids = state.favorites.map(function (f) { return f.questionId; });
      pool = pool.filter(function (q) { return fids.indexOf(q.id) >= 0; });
    }
    if (!pool.length) { toast('没有符合条件的题目'); return; }
    if (mode === 'random') pool.sort(function () { return Math.random() - 0.5; });

    // 无题量概念：一次进入就是全部，流式往下刷
    var key = sessionKey(mode, opts);
    var useProgress = isProgressMode(mode);
    var prog = useProgress ? loadProgress(key) : { idx: 0, answered: {} };
    if (prog.idx >= pool.length) prog.idx = 0;

    state.session = {
      list: pool,
      idx: prog.idx || 0,
      right: 0,
      wrong: 0,
      done: 0,
      mode: mode,
      key: key,
      useProgress: useProgress,
      answeredMap: prog.answered || {}
    };
    // 恢复历史统计（考试宝：打开答题卡可看到过去练习记录）
    state.session.list.forEach(function (q) {
      var v = state.session.answeredMap[q.id];
      if (v === true) state.session.right++;
      else if (v === false) state.session.wrong++;
    });
    state.session.done = state.session.right + state.session.wrong;

    $('setupCard').hidden = true;
    $('quizCard').hidden = false;
    renderQuestion();
  }

  function renderQuestion() {
    var s = state.session;
    var q = s.list[s.idx];
    $('qTypeTag').className = 'tag tag-' + q.type;
    $('qTypeTag').textContent = TYPE_NAME[q.type];
    $('qProgressText').textContent = '第 ' + (s.idx + 1) + '/' + s.list.length +
      ' · 对 ' + s.right + ' · 错 ' + s.wrong + ' · 已做 ' + s.done;
    $('qStem').textContent = stemDisplay(q.stem);
    $('qFeedback').hidden = true;
    $('nextBtn').hidden = true;
    renderSheet();
    renderFavBtn();
    var body = $('qBody');
    body.innerHTML = '';

    if (q.type === 'single') {
      (q.options || []).forEach(function (o) {
        var b = document.createElement('button');
        b.className = 'opt';
        b.textContent = o.key + '. ' + o.text;
        b.onclick = function () { submit(o.key); };
        body.appendChild(b);
      });
    } else if (q.type === 'multi') {
      var checked = [];
      (q.options || []).forEach(function (o) {
        var b = document.createElement('button');
        b.className = 'opt';
        b.textContent = o.key + '. ' + o.text;
        b.onclick = function () {
          var i = checked.indexOf(o.key);
          if (i >= 0) { checked.splice(i, 1); b.classList.remove('selected'); }
          else { checked.push(o.key); b.classList.add('selected'); }
        };
        body.appendChild(b);
      });
      var okBtn = document.createElement('button');
      okBtn.className = 'btn btn-primary btn-block';
      okBtn.textContent = '确认答案';
      okBtn.onclick = function () { submit(checked.sort()); };
      body.appendChild(okBtn);
    } else if (q.type === 'judge') {
      var row = document.createElement('div');
      row.className = 'judge-row';
      ['Y', 'N'].forEach(function (v) {
        var b = document.createElement('button');
        b.className = 'btn';
        b.textContent = v === 'Y' ? '对' : '错';
        b.onclick = function () { submit(v); };
        row.appendChild(b);
      });
      body.appendChild(row);
    } else if (q.type === 'fill') {
      var n = countBlanks(q.stem) || 1;
      var inputs = [];
      for (var i = 0; i < n; i++) {
        var r = document.createElement('div');
        r.className = 'fill-row';
        var inp = document.createElement('input');
        inp.placeholder = '第 ' + (i + 1) + ' 空';
        r.appendChild(inp);
        body.appendChild(r);
        inputs.push(inp);
      }
      var fb = document.createElement('button');
      fb.className = 'btn btn-primary btn-block';
      fb.textContent = '确认答案';
      fb.onclick = function () {
        var vals = inputs.map(function (x) { return x.value.trim(); });
        if (vals.some(function (v) { return !v; })) { toast('请填写所有空'); return; }
        submit(vals);
      };
      body.appendChild(fb);
    } else {
      var showBtn = document.createElement('button');
      showBtn.className = 'btn btn-block';
      showBtn.textContent = '查看参考答案';
      showBtn.onclick = function () {
        var box = document.createElement('div');
        box.className = 'analysis';
        box.textContent = String(q.answer || '');
        body.appendChild(box);
        var row2 = document.createElement('div');
        row2.className = 'judge-row';
        [['known', '掌握了'], ['vague', '模糊'], ['unknown', '没掌握']].forEach(function (pair) {
          var b = document.createElement('button');
          b.className = 'btn';
          b.textContent = pair[1];
          b.onclick = function () { submit(pair[0]); };
          row2.appendChild(b);
        });
        body.appendChild(row2);
        showBtn.hidden = true;
      };
      body.appendChild(showBtn);
    }
  }

  /** 答题卡：题号网格，显示已答/对错，可点击跳题 */
  function renderSheet() {
    var s = state.session;
    var total = s.list.length;
    var win = total > 30 ? 30 : total;
    var start = total > 30 ? Math.max(0, Math.min(s.idx - 8, total - win)) : 0;
    var html = '';
    if (start > 0) html += '<span class="as-ellipsis">…</span>';
    for (var i = start; i < start + win && i < total; i++) {
      var q = s.list[i];
      var a = s.answeredMap[q.id];
      var cls = 'as-item';
      if (i === s.idx) cls += ' as-current';
      else if (a === true) cls += ' as-right';
      else if (a === false) cls += ' as-wrong';
      html += '<button class="' + cls + '" data-jump="' + i + '">' + (i + 1) + '</button>';
    }
    if (start + win < total) html += '<span class="as-ellipsis">…</span>';
    $('answerSheet').innerHTML = html;
    $('answerSheet').querySelectorAll('[data-jump]').forEach(function (b) {
      b.onclick = function () {
        state.session.idx = Number(b.dataset.jump);
        saveProgress(state.session.key, {
          idx: state.session.idx, answered: state.session.answeredMap
        });
        renderQuestion();
      };
    });
  }

  /** 收藏 */
  function snapshotOf(q) {
    return {
      stem: q.stem, type: q.type, options: q.options || [],
      answer: q.answer, answerRule: q.answerRule,
      analysis: q.analysis || '', chapter: q.chapter || ''
    };
  }
  function loadFav() {
    return DB.getAll('favorites').then(function (list) { state.favorites = list || []; });
  }
  function renderFavBtn() {
    var s = state.session;
    if (!s) return;
    var q = s.list[s.idx];
    var on = state.favorites.some(function (f) { return f.questionId === q.id; });
    $('favBtn').textContent = on ? '★ 已收藏' : '☆ 收藏';
    $('favBtn').className = 'btn btn-sm fav-btn' + (on ? ' fav-on' : '');
  }
  $('favBtn').addEventListener('click', function () {
    var q = state.session.list[state.session.idx];
    var exist = state.favorites.some(function (f) { return f.questionId === q.id; });
    var p = exist
      ? DB.del('favorites', q.id)
      : DB.put('favorites', { questionId: q.id, snapshot: snapshotOf(q), createdAt: new Date() });
    p.then(loadFav).then(function () { renderFavBtn(); renderStats(); })
      .catch(function (e) { toast('操作失败：' + e.message); });
  });

  function submit(userAnswer) {
    var s = state.session;
    var q = s.list[s.idx];
    var res = J.judge(q, userAnswer);
    var had = s.answeredMap[q.id];
    if (had !== true && had !== false) s.done++;
    else { if (had === true) s.right--; else s.wrong--; }
    s.answeredMap[q.id] = res.correct;
    if (res.correct) s.right++; else s.wrong++;
    if (s.useProgress) {
      saveProgress(s.key, {
        idx: s.idx + 1 >= s.list.length ? s.idx : s.idx + 1,
        answered: s.answeredMap
      });
    }
    upsertWrong(q, userAnswer, res.correct);
    DB.bulkAdd('records', [{
      questionId: q.id, mode: s.mode,
      userAnswer: userAnswer, correct: res.correct, createdAt: new Date()
    }]);

    var fb = $('qFeedback');
    fb.hidden = false;
    fb.innerHTML =
      '<div class="result ' + (res.correct ? 'result-right' : 'result-wrong') + '">' +
      (res.correct ? '回答正确' : '回答错误') + '</div>' +
      '<div>你的答案：' + esc(userAnswerText(q, userAnswer)) + '</div>' +
      '<div class="q-answer">正确答案：' + esc(answerText(q)) + '</div>' +
      (q.analysis
        ? '<details class="fold"><summary>查看解析</summary><div class="fold-body">' + esc(q.analysis) + '</div></details>'
        : '');

    if (q.type === 'single' || q.type === 'multi') {
      Array.prototype.forEach.call($('qBody').querySelectorAll('.opt'), function (b) {
        var key = b.textContent.split('.')[0].trim();
        var right = q.type === 'single' ? q.answer === key : (q.answer || []).indexOf(key) >= 0;
        if (right) b.classList.add('right');
      });
    }
    renderSheet();
    $('nextBtn').hidden = false;
    var last = s.idx + 1 >= s.list.length;
    $('nextBtn').textContent = last ? '查看结果' : '下一题';
  }

  function userAnswerText(q, ua) {
    if (q.type === 'judge') return ua === 'Y' ? '对' : '错';
    if (q.type === 'fill' && Array.isArray(ua)) return ua.join(' / ');
    if (Array.isArray(ua)) return ua.join('、');
    return ({ known: '掌握了', vague: '模糊', unknown: '没掌握' })[ua] || String(ua);
  }

  $('nextBtn').addEventListener('click', function () {
    var s = state.session;
    if (s.idx + 1 >= s.list.length) {
      var rate = s.done ? Math.round(s.right / s.done * 100) : 0;
      alert('已做到最后一题\n共 ' + s.list.length + ' 题，已做 ' + s.done +
        ' 题\n答对 ' + s.right + ' · 答错 ' + s.wrong + ' · 正确率 ' + rate + '%');
      $('quizCard').hidden = true;
      $('setupCard').hidden = false;
      state.session = null;
      renderModeProgress();
      loadAll();
      return;
    }
    s.idx++;
    if (s.useProgress) {
      saveProgress(s.key, { idx: s.idx, answered: s.answeredMap });
    }
    renderQuestion();
  });

  /** 跳过当前未作答的题目（滑动切题用） */
  function skipQuestion() {
    var s = state.session;
    if (!s) return;
    if (s.idx + 1 >= s.list.length) { toast('已经是最后一题'); return; }
    s.idx++;
    if (s.useProgress) saveProgress(s.key, { idx: s.idx, answered: s.answeredMap });
    renderQuestion();
    showToast('已跳过，可回头再来');
  }

  /** 上一题 */
  function prevQuestion() {
    var s = state.session;
    if (!s) return;
    if (s.idx <= 0) { toast('已经是第一题'); return; }
    s.idx--;
    if (s.useProgress) saveProgress(s.key, { idx: s.idx, answered: s.answeredMap });
    renderQuestion();
  }

  $('prevBtn').addEventListener('click', prevQuestion);

  /** 滑动切题：左滑下一题、右滑上一题 */
  (function bindSwipe() {
    var el = $('quizCard');
    if (!el) return;
    var sx = 0;
    var sy = 0;
    el.addEventListener('touchstart', function (e) {
      var t = e.touches[0];
      sx = t.clientX;
      sy = t.clientY;
    }, { passive: true });
    el.addEventListener('touchend', function (e) {
      var t = e.changedTouches[0];
      var dx = t.clientX - sx;
      var dy = t.clientY - sy;
      if (Math.abs(dx) < 60 || Math.abs(dy) > 40) return;
      if (dx < 0) {
        if ($('nextBtn') && !$('nextBtn').hidden) $('nextBtn').click();
        else skipQuestion();
      } else {
        prevQuestion();
      }
    }, { passive: true });
  })();

  /** 重新练习：清零进度，从头开始 */
  $('restartBtn').addEventListener('click', function () {
    var s = state.session;
    if (!s) return;
    clearProgress(s.key);
    s.answeredMap = {};
    s.idx = 0;
    s.right = 0;
    s.wrong = 0;
    s.done = 0;
    renderQuestion();
    toast('已重新开始');
  });

  /** 错题写入（与云函数逻辑一致：答错归零、答对升 stage，连续答对 3 次掌握） */
  function upsertWrong(q, userAnswer, correct) {
    var now = new Date();
    var existing = null;
    state.wrongs.forEach(function (w) { if (w.questionId === q.id) existing = w; });
    var snapshot = {
      stem: q.stem, type: q.type, options: q.options || [],
      answer: q.answer, answerRule: q.answerRule,
      analysis: q.analysis || '', chapter: q.chapter || ''
    };
    if (!correct) {
      var w = existing || {
        questionId: q.id, wrongCount: 0, stage: 0,
        firstWrongAt: now, mastered: false
      };
      w.wrongCount++;
      w.lastWrongAt = now;
      w.userAnswer = userAnswer;
      w.stage = 0;
      w.nextReviewAt = new Date(now.getTime() + J.REVIEW_INTERVALS[0]);
      w.mastered = false;
      w.snapshot = snapshot;
      DB.put('wrongs', w).then(loadAll);
      return;
    }
    if (existing && !existing.mastered) {
      existing.stage = J.nextStage(existing.stage, true);
      existing.nextReviewAt = new Date(now.getTime() +
        J.REVIEW_INTERVALS[Math.min(existing.stage, J.REVIEW_INTERVALS.length - 1)]);
      if (existing.stage >= J.MASTERED_STAGE) {
        existing.mastered = true;
        existing.masteredAt = now;
      }
      DB.put('wrongs', existing).then(loadAll);
    }
  }

  /* ---------- 错题本 ---------- */
  function renderWrongs() {
    var scope = $('wScope').value;
    var list = state.wrongs.filter(function (w) {
      return scope === 'mastered' ? w.mastered : !w.mastered;
    }).sort(function (a, b) {
      return new Date(b.lastWrongAt) - new Date(a.lastWrongAt);
    });
    $('wrongList').innerHTML = list.length ? list.map(function (w) {
      var due = !w.mastered && new Date(w.nextReviewAt).getTime() <= Date.now();
      return '<div class="q-item"><div class="q-head">' +
        '<span class="tag tag-' + w.snapshot.type + '">' + TYPE_NAME[w.snapshot.type] + '</span>' +
        (due ? '<span class="tag tag-problem">待复习</span>' : '') +
        '<span class="muted">错 ' + w.wrongCount + ' 次 · 进度 ' + w.stage + '/3</span>' +
        '</div>' +
        '<div class="q-stem">' + esc(stemDisplay(w.snapshot.stem)) + '</div>' +
        '<div class="q-answer">答案：' + esc(answerText(w.snapshot)) + '</div>' +
        '<div class="q-actions">' +
        '<button class="btn btn-sm" data-redo="' + w.questionId + '">重做</button>' +
        '<button class="btn btn-sm btn-danger" data-rmw="' + w.questionId + '">移除</button>' +
        '</div></div>';
    }).join('') : '<div class="empty">' + (scope === 'mastered' ? '还没有已掌握的错题' : '太棒了，没有未掌握错题') + '</div>';

    $('wrongList').querySelectorAll('[data-redo]').forEach(function (b) {
      b.onclick = function () {
        var qid = b.dataset.redo;
        var q = state.questions.find(function (x) { return x.id === qid; });
        if (!q) { toast('题目已删除'); return; }
        state.session = { list: [q], idx: 0, right: 0, mode: 'wrong' };
        switchView('practice');
        $('setupCard').hidden = true;
        $('quizCard').hidden = false;
        renderQuestion();
      };
    });
    $('wrongList').querySelectorAll('[data-rmw]').forEach(function (b) {
      b.onclick = function () {
        if (!confirm('移出错题本？')) return;
        DB.del('wrongs', b.dataset.rmw).then(loadAll).then(renderWrongs);
      };
    });
  }
  $('wScope').addEventListener('change', renderWrongs);
  $('exportWrong').addEventListener('click', function () { exportData('wrongs'); });

  /* ---------- 检索 ---------- */
  var searchType = '';

  // 题型筛选 tab（全部 / 单选 / 多选 / 判断 / 填空 / 简答）
  $('typeTabs').addEventListener('click', function (e) {
    var b = e.target.closest('.type-tab');
    if (!b) return;
    searchType = b.dataset.type || '';
    Array.prototype.forEach.call($('typeTabs').querySelectorAll('.type-tab'), function (t) {
      t.classList.toggle('active', t === b);
    });
    if ($('searchInput').value.trim()) doSearch();
  });

  var searchPage = 0;
  var SEARCH_PAGE = 50;
  function doSearch(reset) {
    if (reset) searchPage = 0;
    var raw = $('searchInput').value.trim();
    if (!raw) { toast('请输入关键词'); return; }
    var kws = raw.split(/\s+/).filter(Boolean);
    var hits = state.questions.filter(function (q) {
      var stem = String(q.stem || '');
      return kws.every(function (k) { return stem.indexOf(k) >= 0; });
    });
    if (searchType) {
      hits = hits.filter(function (q) { return q.type === searchType; });
    }
    var typeLabel = { single: '单选题', multi: '多选题', judge: '判断题', fill: '填空题', qa: '简答题' }[searchType] || '';
    $('searchSummary').textContent = '找到 ' + hits.length + ' 条' + (typeLabel ? '（限' + typeLabel + '）' : '');
    var shown = hits.slice(0, (searchPage + 1) * SEARCH_PAGE);
    $('searchList').innerHTML = hits.length ? shown.map(function (q) {
      return '<div class="q-item"><div class="q-head">' +
        '<span class="tag tag-' + q.type + '">' + TYPE_NAME[q.type] + '</span>' +
        (q.chapter ? '<span class="muted">' + esc(q.chapter) + '</span>' : '') +
        '</div>' +
        '<div class="q-stem">' + highlight(stemDisplay(q.stem), kws) + '</div>' +
        renderOptions(q) +
        '<div class="q-answer">答案：' + esc(answerText(q)) +
        '<button class="btn btn-sm copy-btn" data-copy="' + q.id + '">复制</button></div>' +
        (q.analysis ? '<details class="fold"><summary>查看解析</summary><div class="fold-body">' +
          esc(q.analysis) + '</div></details>' : '') +
        '</div>';
    }).join('') + (shown.length < hits.length
      ? '<div class="empty" id="searchMore">下滑加载更多（剩余 ' + (hits.length - shown.length) + ' 条）</div>'
      : '') : '<div class="empty">没有找到相关题目，换个关键词试试</div>';

    $('searchList').querySelectorAll('[data-copy]').forEach(function (b) {
      b.onclick = function () {
        var q = state.questions.find(function (x) { return x.id === b.dataset.copy; });
        if (!q) return;
        copyText(answerText(q));
      };
    });
    var more = $('searchMore');
    if (more) observeMore(more, function () { searchPage++; doSearch(); });
  }

  /** 渲染选项列表（正确答案标绿），让多选题的 ABC 能对上具体内容 */
  function renderOptions(q) {
    if (!q.options || !q.options.length) return '';
    var ansKeys = q.type === 'multi' ? (Array.isArray(q.answer) ? q.answer : []) : [q.answer];
    return '<div class="q-opts">' + q.options.map(function (o) {
      var right = ansKeys.indexOf(o.key) >= 0;
      return '<div class="opt-line' + (right ? ' opt-right' : '') + '">' +
        '<span class="opt-key">' + esc(o.key) + '.</span> ' + esc(o.text) + '</div>';
    }).join('') + '</div>';
  }

  /** 复制到剪贴板（HTTPS 用 clipboard API，失败回退 execCommand） */
  function copyText(text) {
    var done = function () { toast('答案已复制：' + text); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done).catch(function () { fallbackCopy(text, done); });
    } else {
      fallbackCopy(text, done);
    }
  }
  function fallbackCopy(text, done) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); done(); } catch (e) { toast('复制失败，请手动选择'); }
    document.body.removeChild(ta);
  }

  function highlight(text, kws) {
    var out = esc(text);
    kws.forEach(function (k) {
      var safe = esc(k);
      out = out.split(safe).join('<span class="hl">' + safe + '</span>');
    });
    return out;
  }
  $('searchBtn').addEventListener('click', function () { doSearch(true); });
  $('searchInput').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') doSearch(true);
  });
  // 输入防抖自动搜索（300ms）
  var searchTimer = null;
  $('searchInput').addEventListener('input', function () {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(function () {
      if ($('searchInput').value.trim()) doSearch(true);
    }, 300);
  });

  /* ---------- 答题快捷键：1-4 选选项，Enter 下一题 ---------- */
  document.addEventListener('keydown', function (e) {
    var qc = $('quizCard');
    if (!qc || qc.hidden || !state.session) return;
    if (e.key === 'Enter') {
      if ($('nextBtn') && !$('nextBtn').hidden) $('nextBtn').click();
      return;
    }
    var idx = ['1', '2', '3', '4'].indexOf(e.key);
    if (idx < 0) return;
    var q = state.session.list[state.session.idx];
    if (q.type === 'single' && q.options && q.options[idx]) {
      submit(q.options[idx].key);
    } else if (q.type === 'multi' || q.type === 'judge') {
      var btns = $('qBody').querySelectorAll(q.type === 'judge' ? '.judge-row .btn' : '.opt');
      if (btns[idx]) btns[idx].click();
    }
  });

  /* ---------- 初始化 ---------- */
  Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (t) {
    t.onclick = function () { switchView(t.dataset.view); };
  });

  /* ---------- 主题：跟随系统 + 手动切换 + 记忆 ---------- */
  function applyTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    var meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute('content', t === 'dark' ? '#0F172A' : '#4F46E5');
  }
  function initTheme() {
    var saved = null;
    try { saved = localStorage.getItem('stk_theme'); } catch (e) {}
    var sysDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    applyTheme(saved || (sysDark ? 'dark' : 'light'));
    var btn = $('themeBtn');
    if (btn) {
      btn.onclick = function () {
        var cur = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
        applyTheme(cur);
        try { localStorage.setItem('stk_theme', cur); } catch (e) {}
      };
    }
  }
  initTheme();

  /* ---------- 清空所有数据（二次确认） ---------- */
  $('clearAllBtn').addEventListener('click', function () {
    if (!confirm('确定清空所有数据？\n题库、错题、收藏、练习记录和刷题进度都会被删除，且无法恢复。')) return;
    if (!confirm('再次确认：真的要清空全部数据吗？')) return;
    Promise.all([
      DB.clear('questions'), DB.clear('wrongs'), DB.clear('records'),
      DB.clear('batches'), DB.clear('favorites')
    ]).then(function () {
      try {
        Object.keys(localStorage).forEach(function (k) {
          if (k.indexOf('stk_prog_') === 0) localStorage.removeItem(k);
        });
      } catch (e) {}
      return loadAll();
    }).then(function () {
      renderBatches();
      renderBank();
      renderModeProgress();
      renderStats();
      toast('已清空所有数据');
    }).catch(function (e) { toast('清空失败：' + e.message); });
  });

  loadAll().then(function () {
    renderBatches();
    renderModeProgress();
    switchView('practice'); // 默认进入练习页（刷题是主场景）
  });

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('sw.js').catch(function () {});
    });
  }
})();
