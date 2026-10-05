/* ชั้นเชื่อมฐานข้อมูล: Firebase (Google Sign-in + Firestore)  หรือ โหมดทดลอง (localStorage) เมื่อยังไม่ได้ใส่ firebase-config */
import { firebaseConfig, ADMIN_EMAILS, ALLOWED_EMAIL_DOMAIN } from '../firebase-config.js';

const FB_VER = '10.14.1';
const CDN = m => `https://www.gstatic.com/firebasejs/${FB_VER}/firebase-${m}.js`;
export const sheetId = (uid, year, month) => `${uid}_${year}-${String(month).padStart(2,'0')}`;

const isConfigured = c => c && c.apiKey && c.projectId && !/YOUR_|xxxx/i.test(c.apiKey);

/* ================= Firebase ================= */
async function createFirebase(){
  const [{initializeApp}, A, F] = await Promise.all([import(CDN('app')), import(CDN('auth')), import(CDN('firestore'))]);
  const app = initializeApp(firebaseConfig);
  const auth = A.getAuth(app), db = F.getFirestore(app);
  const provider = new A.GoogleAuthProvider(); provider.setCustomParameters({prompt:'select_account'});
  const strip = o => JSON.parse(JSON.stringify(o));      // Firestore ไม่รับ undefined
  const users = () => F.collection(db, 'users');

  async function ensureUser(u){
    if (ALLOWED_EMAIL_DOMAIN && !String(u.email||'').toLowerCase().endsWith('@' + ALLOWED_EMAIL_DOMAIN.toLowerCase())){
      await A.signOut(auth); throw new Error(`ใช้ได้เฉพาะอีเมล @${ALLOWED_EMAIL_DOMAIN}`);
    }
    const ref = F.doc(db, 'users', u.uid), snap = await F.getDoc(ref);
    const isOwner = (ADMIN_EMAILS||[]).map(x=>x.toLowerCase()).includes(String(u.email||'').toLowerCase());
    if (snap.exists()){
      const d = snap.data();
      if (isOwner && d.role !== 'admin'){ await F.updateDoc(ref, {role:'admin'}); d.role='admin'; }
      return {uid:u.uid, ...d};
    }
    const prof = {uid:u.uid, email:u.email, name:u.displayName || u.email, photo:u.photoURL || '', division:'', role:isOwner?'admin':'user', createdAt:F.serverTimestamp()};
    await F.setDoc(ref, prof);
    return {...prof, createdAt:null};
  }

  return {
    mode:'firebase',
    onAuth(cb){
      return A.onAuthStateChanged(auth, async u => {
        if (!u) return cb(null);
        try { cb(await ensureUser(u)); } catch (e){ cb(null, e); }
      });
    },
    signIn: () => A.signInWithPopup(auth, provider),
    signOut: () => A.signOut(auth),
    async saveProfile(uid, p){ await F.setDoc(F.doc(db,'users',uid), strip(p), {merge:true}); },
    async listUsers(){ return (await F.getDocs(users())).docs.map(d => ({uid:d.id, ...d.data()})); },
    async setRole(uid, role){ await F.updateDoc(F.doc(db,'users',uid), {role}); },
    async getConfig(){ const s = await F.getDoc(F.doc(db,'config','main')); return s.exists() ? s.data() : null; },
    async saveConfig(cfg){ await F.setDoc(F.doc(db,'config','main'), strip({...cfg, updatedAt:new Date().toISOString()})); },
    async getSheet(uid, year, month){ const s = await F.getDoc(F.doc(db,'timesheets',sheetId(uid,year,month))); return s.exists() ? s.data() : null; },
    async saveSheet(doc){ await F.setDoc(F.doc(db,'timesheets',sheetId(doc.uid,doc.year,doc.month)), strip({...doc, updatedAt:new Date().toISOString()})); },
    async listSheets({year, uid}){
      const c = F.collection(db,'timesheets');
      const q = uid ? F.query(c, F.where('year','==',year), F.where('uid','==',uid)) : F.query(c, F.where('year','==',year));
      return (await F.getDocs(q)).docs.map(d => d.data());
    },
    async deleteSheet(uid, year, month){ await F.deleteDoc(F.doc(db,'timesheets',sheetId(uid,year,month))); },
    async exportAll(){
      const ser = o => JSON.parse(JSON.stringify(o, (k, v) => v && typeof v.toDate === 'function' ? v.toDate().toISOString() : v));
      const [u, t, c] = await Promise.all([F.getDocs(users()), F.getDocs(F.collection(db,'timesheets')), F.getDoc(F.doc(db,'config','main'))]);
      return {users:u.docs.map(d => ser({...d.data(), uid:d.id})), sheets:t.docs.map(d => ser(d.data())), config:c.exists() ? ser(c.data()) : null};
    },
    async importAll({users:us=[], sheets:ss=[], config=null}){
      const ops = [];
      us.forEach(u => ops.push([F.doc(db,'users',u.uid), strip(u)]));
      ss.forEach(d => ops.push([F.doc(db,'timesheets',sheetId(d.uid,d.year,d.month)), strip(d)]));
      if (config) ops.push([F.doc(db,'config','main'), strip(config)]);
      for (let i=0; i<ops.length; i+=400){ const b = F.writeBatch(db); ops.slice(i, i+400).forEach(([r,d]) => b.set(r, d)); await b.commit(); }
      return ops.length;
    }
  };
}

/* ================= โหมดทดลอง (ไม่ต้องมี Firebase) ================= */
function createDemo(){
  const g = (k, d) => { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch(e){ return d; } };
  const s = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch(e){} };
  let cb = null;
  const cur = () => g('demo_current', null);
  const usersMap = () => g('demo_users', {});
  return {
    mode:'demo',
    onAuth(f){ cb = f; const c = cur(); const u = c && usersMap()[c]; setTimeout(() => f(u || null), 0); return () => {}; },
    async signIn(email, name){
      email = (email || 'demo@example.com').trim().toLowerCase(); const us = usersMap();
      let u = Object.values(us).find(x => x.email === email);
      if (!u){ u = {uid:'demo_' + Object.keys(us).length + '_' + email.replace(/\W/g,''), email, name:name || email.split('@')[0], division:'', role:Object.keys(us).length ? 'user' : 'admin', photo:''}; us[u.uid] = u; s('demo_users', us); }
      s('demo_current', u.uid); cb && cb(u);
    },
    async signOut(){ s('demo_current', null); cb && cb(null); },
    async saveProfile(uid, p){ const us = usersMap(); us[uid] = {...us[uid], ...p}; s('demo_users', us); },
    async listUsers(){ return Object.values(usersMap()); },
    async setRole(uid, role){ const us = usersMap(); if (us[uid]) us[uid].role = role; s('demo_users', us); },
    async getConfig(){ return g('demo_config', null); },
    async saveConfig(cfg){ s('demo_config', cfg); },
    async getSheet(uid, y, m){ return g('demo_sheets', {})[sheetId(uid,y,m)] || null; },
    async saveSheet(doc){ const all = g('demo_sheets', {}); all[sheetId(doc.uid,doc.year,doc.month)] = {...doc, updatedAt:new Date().toISOString()}; s('demo_sheets', all); },
    async listSheets({year, uid}){ return Object.values(g('demo_sheets', {})).filter(d => d.year === year && (!uid || d.uid === uid)); },
    async deleteSheet(uid, y, m){ const all = g('demo_sheets', {}); delete all[sheetId(uid,y,m)]; s('demo_sheets', all); },
    async exportAll(){ return {users:Object.values(usersMap()), sheets:Object.values(g('demo_sheets', {})), config:g('demo_config', null)}; },
    async importAll({users:us=[], sheets:ss=[], config=null}){
      const um = usersMap(); us.forEach(u => { um[u.uid] = u; }); s('demo_users', um);
      const all = g('demo_sheets', {}); ss.forEach(d => { all[sheetId(d.uid,d.year,d.month)] = d; }); s('demo_sheets', all);
      if (config) s('demo_config', config);
      return us.length + ss.length + (config ? 1 : 0);
    }
  };
}

export async function createBackend(){
  if (isConfigured(firebaseConfig)) return createFirebase();
  return createDemo();
}
