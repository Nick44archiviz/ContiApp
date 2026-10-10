// Backend Firebase: login Google + Firestore. Struttura dati:
//   users/{uid}                 -> { email, name, viewers: [email di chi può guardare in sola lettura] }
//   users/{uid}/config/main     -> { open, startYear, goal, acc, cats, inc }
//   users/{uid}/months/{YYYY-MM}-> { rows: { <id>: { d, c, t, v:[4 importi], o } } }
import { firebaseConfig } from './firebase-config.js';
const V = '10.12.2';
const base = `https://www.gstatic.com/firebasejs/${V}/`;

export async function createFirebaseBackend() {
  const [appM, authM, fs] = await Promise.all([
    import(base + 'firebase-app.js'), import(base + 'firebase-auth.js'), import(base + 'firebase-firestore.js')
  ]);
  const app = appM.initializeApp(firebaseConfig);
  const auth = authM.getAuth(app);
  let db;
  try {
    db = fs.initializeFirestore(app, { localCache: fs.persistentLocalCache({ tabManager: fs.persistentMultipleTabManager() }) });
  } catch (e) { db = fs.getFirestore(app); }
  const uref = id => fs.doc(db, 'users', id);
  const mref = (id, docId) => fs.doc(db, 'users', id, 'months', docId);
  const cref = id => fs.doc(db, 'users', id, 'config', 'main');
  const noop = () => {};
  return {
    kind: 'firebase',
    onAuth(cb) {
      return authM.onAuthStateChanged(auth, u => cb(u ? { uid: u.uid, email: (u.email || '').toLowerCase(), name: u.displayName || u.email, verified: u.emailVerified } : null));
    },
    async signIn() {
      const p = new authM.GoogleAuthProvider();
      try { await authM.signInWithPopup(auth, p); }
      catch (e) {
        if (['auth/popup-blocked', 'auth/operation-not-supported-in-this-environment', 'auth/web-storage-unsupported'].includes(e.code)) return authM.signInWithRedirect(auth, p);
        if (e.code === 'auth/popup-closed-by-user' || e.code === 'auth/cancelled-popup-request') return;
        throw e;
      }
    },
    signOut() { return authM.signOut(auth); },
    async ensureProfile(u) {
      const s = await fs.getDoc(uref(u.uid));
      if (!s.exists()) await fs.setDoc(uref(u.uid), { email: u.email, name: u.name || u.email, viewers: [] });
      else if (s.data().email !== u.email) await fs.updateDoc(uref(u.uid), { email: u.email });
    },
    watchProfile(id, cb) { return fs.onSnapshot(uref(id), s => cb(s.exists() ? s.data() : null), noop); },
    watchMonths(id, cb, err) {
      return fs.onSnapshot(fs.collection(db, 'users', id, 'months'), s => {
        const m = {}; s.forEach(d => { m[d.id] = d.data(); }); cb(m, s.metadata.fromCache);
      }, err || noop);
    },
    watchConfig(id, cb) { return fs.onSnapshot(cref(id), s => cb(s.exists() ? s.data() : null), noop); },
    watchShared(mail, cb) {
      const q = fs.query(fs.collection(db, 'users'), fs.where('viewers', 'array-contains', mail));
      return fs.onSnapshot(q, s => cb(s.docs.map(d => ({ uid: d.id, email: d.data().email, name: d.data().name || d.data().email }))), noop);
    },
    putRow(id, docId, rid, row) { return fs.setDoc(mref(id, docId), { rows: { [rid]: row } }, { merge: true }); },
    async delRow(id, docId, rid) { try { await fs.updateDoc(mref(id, docId), { ['rows.' + rid]: fs.deleteField() }); } catch (e) { /* mese inesistente: niente da togliere */ } },
    setConfig(id, cfg) { return fs.setDoc(cref(id), cfg, { merge: true }); },
    async replaceMonths(id, map) {
      const b = fs.writeBatch(db);
      for (const k in map) b.set(mref(id, k), map[k]);
      return b.commit();
    },
    setViewers(id, list) { return fs.updateDoc(uref(id), { viewers: list }); }
  };
}
