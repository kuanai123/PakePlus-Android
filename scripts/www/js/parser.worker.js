/**
 * parser.worker.js — 题库解析 Worker
 * 把 Excel/Word 解析移出主线程，避免大文件（数千题）解析时页面卡死
 * 通过 importScripts 复用已验证的 parser.js 业务逻辑（与主线程完全一致）
 */
/**
 * 加载依赖（路径相对 worker 脚本位置）
 * parser.js 内部使用 window.XLSX / window.mammoth / window.ParserCore，
 * Worker 全局对象是 self，这里做一次别名
 */
self.window = self;
try { importScripts('../lib/xlsx.full.min.js'); } catch (e) { /* Excel 解析将不可用 */ }
try { importScripts('./parser.js'); } catch (e) { /* 解析核心加载失败 */ }
// mammoth 浏览器版在部分环境依赖 DOM，加载失败不影响 Excel 解析
try { importScripts('../lib/mammoth.browser.js'); } catch (e) { /* Word 解析回退主线程 */ }

self.onmessage = function (e) {
  var data = e.data || {};
  var buffer = data.buffer;
  var fileName = data.fileName || '';

  if (!buffer) {
    self.postMessage({ ok: false, error: '没有收到文件内容' });
    return;
  }

  self.postMessage({ type: 'progress', text: '正在解析…' });

  try {
    // 复用主线程同一套解析逻辑，避免两套实现产生差异
    window.ParserCore.parseAny(buffer, fileName).then(function (parsed) {
      if (parsed && parsed.error) {
        self.postMessage({ ok: false, error: parsed.error });
        return;
      }
      self.postMessage({ ok: true, parsed: parsed });
    }).catch(function (err) {
      self.postMessage({ ok: false, error: String(err && err.message || err) });
    });
  } catch (err) {
    self.postMessage({ ok: false, error: String(err && err.message || err) });
  }
};
