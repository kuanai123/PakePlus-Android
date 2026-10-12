/**
 * db.js — IndexedDB 本地存储（安卓/浏览器端，完全离线、无服务器）
 * stores: questions(题库) / wrongs(错题) / records(练习记录) / batches(导入批次)
 */
(function () {
  var DB_NAME = 'shuatikaoshi';
  var VERSION = 2; // v2 新增 favorites（收藏）
  var STORES = ['questions', 'wrongs', 'records', 'batches', 'favorites'];
  var _db = null;

  function open() {
    return new Promise(function (resolve, reject) {
      if (_db) return resolve(_db);
      var req = indexedDB.open(DB_NAME, VERSION);
      req.onupgradeneeded = function (e) {
        var db = e.target.result;
        STORES.forEach(function (s) {
          if (db.objectStoreNames.contains(s)) return;
          if (s === 'records') {
            var os = db.createObjectStore(s, { keyPath: 'id', autoIncrement: true });
            os.createIndex('createdAt', 'createdAt');
          } else if (s === 'wrongs' || s === 'favorites') {
            db.createObjectStore(s, { keyPath: 'questionId' });
          } else {
            var os2 = db.createObjectStore(s, { keyPath: 'id' });
            if (s === 'questions') os2.createIndex('type', 'type');
          }
        });
      };
      req.onsuccess = function (e) { _db = e.target.result; resolve(_db); };
      req.onerror = function (e) { reject(e.target.error); };
    });
  }

  function store(name, mode) {
    return open().then(function (db) {
      return db.transaction(name, mode || 'readonly').objectStore(name);
    });
  }

  function wrap(request) {
    return new Promise(function (resolve, reject) {
      request.onsuccess = function () { resolve(request.result); };
      request.onerror = function () { reject(request.error); };
    });
  }

  function getAll(name) {
    return store(name).then(function (os) { return wrap(os.getAll()); });
  }

  function get(name, id) {
    return store(name).then(function (os) { return wrap(os.get(id)); });
  }

  function put(name, obj) {
    return store(name, 'readwrite').then(function (os) { return wrap(os.put(obj)); });
  }

  function del(name, id) {
    return store(name, 'readwrite').then(function (os) { return wrap(os.delete(id)); });
  }

  function clear(name) {
    return store(name, 'readwrite').then(function (os) { return wrap(os.clear()); });
  }

  function count(name) {
    return store(name).then(function (os) { return wrap(os.count()); });
  }

  /** 批量写入（同一事务，2000 条也能秒级完成） */
  function bulkAdd(name, list) {
    return open().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(name, 'readwrite');
        var os = tx.objectStore(name);
        list.forEach(function (item) { os.put(item); });
        tx.oncomplete = function () { resolve(list.length); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function newId() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  window.DB = {
    open: open,
    getAll: getAll,
    get: get,
    put: put,
    del: del,
    clear: clear,
    count: count,
    bulkAdd: bulkAdd,
    newId: newId
  };
})();
