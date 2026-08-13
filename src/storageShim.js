// Local stand-in for the Claude artifact `window.storage` API, backed by
// the browser's localStorage. Mirrors the same method signatures so
// App.jsx (copied straight from the Claude artifact) needs no edits.
function k(key) {
  return `ats-storage:${key}`;
}

window.storage = {
  async get(key) {
    const raw = localStorage.getItem(k(key));
    if (raw === null) throw new Error(`Key not found: ${key}`);
    return { key, value: raw, shared: false };
  },
  async set(key, value) {
    localStorage.setItem(k(key), value);
    return { key, value, shared: false };
  },
  async delete(key) {
    const existed = localStorage.getItem(k(key)) !== null;
    localStorage.removeItem(k(key));
    return { key, deleted: existed, shared: false };
  },
  async list(prefix = "") {
    const keys = Object.keys(localStorage)
      .filter((full) => full.startsWith("ats-storage:") && full.slice(12).startsWith(prefix))
      .map((full) => full.slice(12));
    return { keys, prefix, shared: false };
  },
};
