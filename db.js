/* Lưu trữ offline bằng IndexedDB (key-value đơn giản). */
(function (root) {
  'use strict';
  const DB = 'kiemke-freshshop', STORE = 'kv';
  let dbp = null;
  function open() {
    if (dbp) return dbp;
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbp;
  }
  async function tx(mode, fn) {
    const db = await open();
    return new Promise((resolve, reject) => {
      const t = db.transaction(STORE, mode), st = t.objectStore(STORE), req = fn(st);
      t.oncomplete = () => resolve(req && req.result);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  }
  root.KKDB = {
    get: key => tx('readonly', st => st.get(key)),
    set: (key, val) => tx('readwrite', st => st.put(val, key)),
    del: key => tx('readwrite', st => st.delete(key))
  };
})(self);
