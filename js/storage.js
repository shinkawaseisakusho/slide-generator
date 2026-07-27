/* 下書きの保存。

   IndexedDB を使う。localStorage は 5MB 前後が上限で、画像を10枚ほど
   入れると保存できなくなるため。IndexedDB が使えない環境（プライベート
   モードなど）では localStorage に落とし、そのときだけ容量の大きい
   動画を保存対象から外す。 */

const LS_KEY = 'slide-generator/v1';
const DB_NAME = 'slide-generator';
const DB_STORE = 'drafts';
const DB_KEY = 'current';

let dbPromise = null;

function openDb() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    if (!globalThis.indexedDB) { reject(new Error('IndexedDB が使えません')); return; }
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(DB_STORE)) req.result.createObjectStore(DB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('IndexedDB を開けません'));
    req.onblocked = () => reject(new Error('IndexedDB がブロックされました'));
  }).catch((err) => { dbPromise = null; throw err; });
  return dbPromise;
}

function run(mode, fn) {
  return openDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, mode);
    const req = fn(tx.objectStore(DB_STORE));
    tx.oncomplete = () => resolve(req && req.result);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('保存が中断されました'));
  }));
}

// localStorage に落とすときは動画を外す（base64のままでは到底入らない）
function withoutVideo(draft) {
  return {
    ...draft,
    slides: (draft.slides || []).map((s) => ({
      ...s,
      media: (s.media || []).map((m) => (m.kind === 'video' ? { ...m, data: null, cover: null } : m)),
    })),
  };
}

/**
 * 下書きを保存する。
 * @returns {Promise<'idb'|'local'|'failed'>} どこに保存できたか
 */
export async function saveDraft(draft) {
  try {
    await run('readwrite', (store) => store.put(draft, DB_KEY));
    return 'idb';
  } catch (err) {
    console.warn('IndexedDB に保存できませんでした:', err);
  }

  try {
    localStorage.setItem(LS_KEY, JSON.stringify(withoutVideo(draft)));
    return 'local';
  } catch (_) {
    return 'failed';
  }
}

/** 下書きを読み出す。無ければ null */
export async function loadDraft() {
  try {
    const data = await run('readonly', (store) => store.get(DB_KEY));
    if (data) return data;
  } catch (err) {
    console.warn('IndexedDB から復元できませんでした:', err);
  }

  // 旧バージョンの下書き、または IndexedDB が使えない環境
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) return JSON.parse(raw);
  } catch (_) { /* noop */ }

  return null;
}

export async function clearDraft() {
  try { await run('readwrite', (store) => store.delete(DB_KEY)); } catch (_) { /* noop */ }
  try { localStorage.removeItem(LS_KEY); } catch (_) { /* noop */ }
}
