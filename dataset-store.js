/* =========================================================
   DATASET STORE — penyimpanan data aset aktif (IndexedDB)
   Menyimpan seluruh dataset yang sedang dipakai sehingga
   hasil unggah Excel / tambah barang tetap ada setelah reload.
   ========================================================= */
(function () {
  'use strict';

  var NAME = 'simpati_dataset_v1';
  var STORE = 'dataset';
  var KEY = 'active';
  var openPromise = null;

  function open() {
    if (!window.indexedDB) return Promise.reject(new Error('IndexedDB tidak tersedia'));
    if (openPromise) return openPromise;
    openPromise = new Promise(function (resolve, reject) {
      var req = indexedDB.open(NAME, 1);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error || new Error('Gagal membuka penyimpanan')); };
      req.onblocked = function () { reject(new Error('Penyimpanan sedang dipakai tab lain')); };
    });
    return openPromise;
  }

  function store(mode) {
    return open().then(function (db) {
      return db.transaction(STORE, mode).objectStore(STORE);
    });
  }

  window.DatasetStore = {
    supported: function () { return !!window.indexedDB; },

    get: function () {
      return store('readonly').then(function (s) {
        return new Promise(function (resolve, reject) {
          var req = s.get(KEY);
          req.onsuccess = function () { resolve(req.result || null); };
          req.onerror = function () { reject(req.error); };
        });
      });
    },

    set: function (record) {
      return store('readwrite').then(function (s) {
        return new Promise(function (resolve, reject) {
          var req = s.put(record, KEY);
          req.onsuccess = function () { resolve(true); };
          req.onerror = function () { reject(req.error); };
        });
      });
    },

    clear: function () {
      return store('readwrite').then(function (s) {
        return new Promise(function (resolve, reject) {
          var req = s.delete(KEY);
          req.onsuccess = function () { resolve(true); };
          req.onerror = function () { reject(req.error); };
        });
      });
    },

    usage: function () {
      if (!navigator.storage || !navigator.storage.estimate) return Promise.resolve(null);
      return navigator.storage.estimate().then(function (e) {
        return { usage: e.usage || 0, quota: e.quota || 0 };
      }).catch(function () { return null; });
    }
  };
})();
