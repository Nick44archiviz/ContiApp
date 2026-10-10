// Backend di prova: salva tutto nel browser (localStorage). Serve per la modalità demo e per i test.
// Con ?as=nome@esempio.it si simulano utenti diversi sullo stesso browser.
const KEY = 'conti.local.v1';
const load = () => { try { return JSON.parse(localStorage.getItem(KEY)) || { users: {} }; } catch (e) { return { users: {} }; } };
const save = db => localStorage.setItem(KEY, JSON.stringify(db));

export function createLocalBackend() {
  const email = (new URLSearchParams(location.search).get('as') || 'demo@locale.test').toLowerCase();
  const uid = 'u_' + email.replace(/[^a-z0-9]/g, '_');
  const listeners = new Set();
  const bump = () => listeners.forEach(f => f());
  window.addEventListener('storage', bump);
  const user = id => { const db = load(); return [db, db.users[id] || (db.users[id] = { profile: null, months: {}, config: null })]; };
  const watch = reader => {
    const run = () => reader(load());
    listeners.add(run); setTimeout(run, 30);
    return () => listeners.delete(run);
  };
  const clone = x => x == null ? x : JSON.parse(JSON.stringify(x));
  return {
    kind: 'local',
    onAuth(cb) { setTimeout(() => cb({ uid, email, name: email.split('@')[0] }), 20); return () => {}; },
    async signIn() {}, async signOut() {},
    async ensureProfile(u) {
      const [db, d] = user(u.uid);
      if (!d.profile) { d.profile = { email: u.email, name: u.name, viewers: [] }; save(db); bump(); }
    },
    watchProfile(id, cb) { return watch(db => cb(clone((db.users[id] || {}).profile))); },
    watchMonths(id, cb) { return watch(db => cb(clone((db.users[id] || {}).months || {}))); },
    watchConfig(id, cb) { return watch(db => cb(clone((db.users[id] || {}).config))); },
    watchShared(mail, cb) {
      return watch(db => cb(Object.entries(db.users)
        .filter(([id, d]) => id !== uid && d.profile && (d.profile.viewers || []).includes(mail))
        .map(([id, d]) => ({ uid: id, email: d.profile.email, name: d.profile.name }))));
    },
    async putRow(id, docId, rid, row) { const [db, d] = user(id); (d.months[docId] = d.months[docId] || { rows: {} }).rows[rid] = row; save(db); bump(); },
    async delRow(id, docId, rid) { const [db, d] = user(id); if (d.months[docId]) delete d.months[docId].rows[rid]; save(db); bump(); },
    async setConfig(id, cfg) { const [db, d] = user(id); d.config = Object.assign({}, d.config || {}, cfg); save(db); bump(); },
    async replaceMonths(id, map) { const [db, d] = user(id); Object.assign(d.months, clone(map)); save(db); bump(); },
    async setViewers(id, list) { const [db, d] = user(id); d.profile = Object.assign({}, d.profile, { viewers: list }); save(db); bump(); }
  };
}
