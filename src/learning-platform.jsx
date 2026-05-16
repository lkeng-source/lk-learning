import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";

const C = {
  navy: "#1B3A5C", navyDark: "#122840", navyLight: "#2A5080",
  gold: "#D4A528", goldLight: "#F0C850", goldPale: "#FFF8E7",
  accent: "#2980B9", bg: "#F5F6FA", card: "#FFFFFF",
  text: "#1E293B", textMid: "#475569", textLight: "#94A3B8",
  border: "#E2E8F0", success: "#16A34A", warning: "#D97706", danger: "#DC2626",
};

const ICONS = ["⚙️","✅","🛡️","🌿","💻","📊","🏭","🔬","📋","🎯","💡","🔧","📦","🧪","🏗️","📐"];
const COLORS = ["#0066CC","#00875A","#E65100","#2E7D32","#5B21B6","#B45309","#DC2626","#0891B2","#7C3AED","#059669"];

const DEFAULT_CATS = [
  { id: "sys1", name: "生產系統", icon: "⚙️", color: "#0066CC", order: 1 },
  { id: "sys2", name: "品質系統", icon: "✅", color: "#00875A", order: 2 },
  { id: "sys3", name: "安全系統", icon: "🛡️", color: "#E65100", order: 3 },
  { id: "sys4", name: "環境系統", icon: "🌿", color: "#2E7D32", order: 4 },
  { id: "sys5", name: "資訊系統", icon: "💻", color: "#5B21B6", order: 5 },
  { id: "mgmt", name: "管理類", icon: "📊", color: "#B45309", order: 6 },
];

// 從 YouTube 各種網址格式擷取影片 ID
const getYouTubeId = (url) => {
  if (!url) return null;
  const m = url.match(/(?:youtu\.be\/|youtube\.com\/(?:watch\?v=|embed\/|v\/|shorts\/))([^?&\s]+)/);
  return m ? m[1] : null;
};

const INIT_COURSES = [
  { id:"c1", title:"5S 管理實務", category:"sys1", instructor:"王大明", duration:45, thumbnail:"📘", description:"學習5S管理方法，提升生產效率與工作環境品質。", views:342, publishDate:"2026-05-10", status:"published", files:[], chapters:[{title:"5S 概論",duration:15,youtubeUrl:"https://www.youtube.com/watch?v=dQw4w9WgXcQ"},{title:"整理與整頓",duration:15,youtubeUrl:""},{title:"清掃、清潔、素養",duration:15,youtubeUrl:""}], quiz:[{q:"5S 中的「整理」指的是什麼？",options:["區分需要與不需要的物品","將物品擺放整齊","打掃環境","維持好習慣"],answer:0},{q:"5S 實施的正確順序為何？",options:["整理→整頓→清掃→清潔→素養","清掃→整理→整頓→清潔→素養","素養→整理→整頓→清掃→清潔","整頓→整理→清掃→清潔→素養"],answer:0}] },
  { id:"c2", title:"ISO 9001 品質管理入門", category:"sys2", instructor:"李美玲", duration:60, thumbnail:"📗", description:"深入了解 ISO 9001 品質管理系統的核心要求與實施方法。", views:528, publishDate:"2026-05-08", status:"published", files:[], chapters:[{title:"ISO 9001 概述",duration:20,youtubeUrl:""},{title:"品質管理原則",duration:20,youtubeUrl:""},{title:"文件化要求",duration:20,youtubeUrl:""}], quiz:[{q:"ISO 9001 最新版本是哪一年發布？",options:["2008","2012","2015","2020"],answer:2}] },
  { id:"c3", title:"職業安全衛生基礎", category:"sys3", instructor:"張志強", duration:50, thumbnail:"📕", description:"了解職場安全衛生法規與防災基本知識。", views:891, publishDate:"2026-05-01", status:"published", files:[], chapters:[{title:"安全衛生法規",duration:20,youtubeUrl:""},{title:"危害辨識",duration:15,youtubeUrl:""},{title:"事故預防",duration:15,youtubeUrl:""}], quiz:[{q:"以下何者屬於物理性危害？",options:["噪音","化學溶劑","病毒","心理壓力"],answer:0}] },
  { id:"c4", title:"環境管理 ISO 14001", category:"sys4", instructor:"陳怡君", duration:55, thumbnail:"📗", description:"學習環境管理系統的規劃與實施。", views:267, publishDate:"2026-04-20", status:"published", files:[], chapters:[{title:"環境管理概論",duration:20,youtubeUrl:""},{title:"環境因素鑑別",duration:20,youtubeUrl:""},{title:"法規遵循",duration:15,youtubeUrl:""}], quiz:[{q:"ISO 14001 屬於哪種管理系統？",options:["品質管理","環境管理","安全管理","資訊管理"],answer:1}] },
  { id:"c5", title:"資訊安全意識訓練", category:"sys5", instructor:"林宗翰", duration:40, thumbnail:"📘", description:"提升資訊安全意識，防範社交工程攻擊。", views:1203, publishDate:"2026-05-12", status:"published", files:[], chapters:[{title:"常見網路威脅",duration:15,youtubeUrl:""},{title:"密碼管理",duration:10,youtubeUrl:""},{title:"釣魚郵件辨識",duration:15,youtubeUrl:""}], quiz:[{q:"強密碼應包含哪些元素？",options:["僅數字","大小寫字母+數字+特殊符號","僅英文字母","生日"],answer:1}] },
  { id:"c6", title:"主管領導力培訓", category:"mgmt", instructor:"黃雅琪", duration:90, thumbnail:"📙", description:"培養中階主管的領導能力與團隊管理技巧。", views:456, publishDate:"2026-05-05", status:"published", files:[], chapters:[{title:"領導力概論",duration:30,youtubeUrl:""},{title:"溝通與激勵",duration:30,youtubeUrl:""},{title:"績效管理",duration:30,youtubeUrl:""}], quiz:[{q:"情境領導理論認為領導風格應依據什麼調整？",options:["領導者個性","部屬成熟度","組織規模","市場環境"],answer:1}] },
  { id:"c7", title:"ERP 系統操作實務", category:"sys5", instructor:"吳建宏", duration:70, thumbnail:"📘", description:"企業資源規劃系統的日常操作與報表查詢。", views:189, publishDate:"2026-04-15", status:"published", files:[], chapters:[{title:"ERP 基礎概念",duration:25,youtubeUrl:""},{title:"模組操作",duration:25,youtubeUrl:""},{title:"報表與分析",duration:20,youtubeUrl:""}], quiz:[{q:"ERP 的全名是什麼？",options:["Enterprise Resource Planning","Enterprise Risk Prevention","Electronic Resource Platform","Efficient Resource Processing"],answer:0}] },
  { id:"c8", title:"消防安全訓練", category:"sys3", instructor:"劉國華", duration:35, thumbnail:"📕", description:"滅火器使用與緊急疏散演練。", views:723, publishDate:"2026-05-11", status:"published", files:[], chapters:[{title:"滅火器種類與使用",duration:15,youtubeUrl:""},{title:"緊急疏散程序",duration:20,youtubeUrl:""}], quiz:[{q:"使用滅火器的口訣是什麼？",options:["拉、瞄、壓、掃","開、瞄、射、收","按、拉、噴、掃","提、拔、握、掃"],answer:0}] },
];

const INIT_USERS = [
  { id:"u1", empNo:"E00001", name:"系統管理員", email:"admin@lkeng.com", password:"LK@dmin2026", role:"admin", department:"資訊部", mustChangePw:false },
  { id:"u2", empNo:"E00002", name:"陳小明", email:"chen@lkeng.com", password:"E00002", role:"user", department:"生產部", mustChangePw:true },
  { id:"u3", empNo:"E00003", name:"林美玲", email:"lin@lkeng.com", password:"E00003", role:"user", department:"品保部", mustChangePw:true },
  { id:"u4", empNo:"E00004", name:"王志偉", email:"wang@lkeng.com", password:"E00004", role:"user", department:"管理部", mustChangePw:true },
];

const inp = { width:"100%", padding:"9px 12px", borderRadius:7, border:`1px solid ${C.border}`, fontSize:13, outline:"none", boxSizing:"border-box", background:"#FFF", color:C.text };

function Btn({children, onClick, variant="primary", disabled, style}) {
  const v = {
    primary: { background:C.navy, color:"#FFF" },
    gold: { background:C.gold, color:"#FFF" },
    outline: { background:"transparent", border:`1px solid ${C.border}`, color:C.textMid },
    danger: { background:"transparent", border:`1px solid ${C.danger}40`, color:C.danger },
    ghost: { background:"transparent", color:C.navy },
  };
  return <button onClick={onClick} disabled={disabled} style={{ padding:"8px 16px", borderRadius:7, border:"none", fontSize:13, fontWeight:500, cursor:disabled?"default":"pointer", opacity:disabled?0.5:1, ...v[variant], ...style }}>{children}</button>;
}

function Field({ label, children }) {
  return <div style={{ marginBottom:10 }}><label style={{ display:"block", color:C.textMid, fontSize:12, marginBottom:4, fontWeight:500 }}>{label}</label>{children}</div>;
}

export default function App() {
  const [categories, setCategories] = useState(DEFAULT_CATS);
  const [courses, setCourses] = useState(INIT_COURSES);
  const [users, setUsers] = useState(INIT_USERS);
  const [currentUser, setCurrentUser] = useState(null);
  const [view, setView] = useState("login");
  const [watchHistory, setWatchHistory] = useState({});
  const [quizResults, setQuizResults] = useState({});

  const handleLogin = (email, pw) => {
    const u = users.find(x => x.email.toLowerCase().trim() === email.toLowerCase().trim() && x.password === pw);
    if (u) { setCurrentUser(u); setView(u.role === "admin" ? "admin" : "front"); return true; }
    return false;
  };
  const handleLogout = () => { setCurrentUser(null); setView("login"); };
  const recordWatch = (cid, chIdx, prog) => {
    if (!currentUser) return;
    setWatchHistory(prev => {
      const k = `${currentUser.id}_${cid}`;
      return { ...prev, [k]: { courseId:cid, chapterIndex:chIdx, progress:prog, lastWatched:new Date().toISOString(), totalTime:(prev[k]?.totalTime||0)+1 } };
    });
    setCourses(prev => prev.map(c => c.id===cid ? {...c, views:c.views+1} : c));
  };
  const saveQuiz = (cid, score, total) => {
    if (!currentUser) return;
    setQuizResults(prev => ({ ...prev, [`${currentUser.id}_${cid}`]: { score, total, date:new Date().toISOString(), userName:currentUser.name } }));
  };
  const updatePassword = (newPw) => {
    setUsers(prev => prev.map(u => u.id===currentUser.id ? {...u, password:newPw, mustChangePw:false} : u));
    setCurrentUser(prev => ({...prev, password:newPw, mustChangePw:false}));
  };

  // 依 order 排序分類（沒有 order 的補在最後）
  const sortedCategories = useMemo(() =>
    [...categories].sort((a,b) => (a.order ?? 9999) - (b.order ?? 9999)),
    [categories]
  );

  if (view==="login") return <Login onLogin={handleLogin} />;
  if (view==="admin") return <Admin {...{categories:sortedCategories,setCategories,courses,setCourses,users,setUsers,watchHistory,quizResults,onLogout:handleLogout,setView,currentUser,updatePassword}} />;
  return <Front {...{categories:sortedCategories,courses:courses.filter(c=>c.status==="published"),currentUser,onLogout:handleLogout,watchHistory,recordWatch,quizResults,saveQuiz,setView,updatePassword}} />;
}

function Login({ onLogin }) {
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = () => {
    setErr("");
    if (!email || !pw) { setErr("請輸入帳號與密碼"); return; }
    setLoading(true);
    setTimeout(() => { if (!onLogin(email, pw)) setErr("帳號或密碼錯誤"); setLoading(false); }, 400);
  };

  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:`linear-gradient(135deg, ${C.navyDark} 0%, ${C.navy} 50%, ${C.navyLight} 100%)`, fontFamily:"'Noto Sans TC',sans-serif", position:"relative", padding:20 }}>
      <div style={{ position:"absolute", top:-100, right:-100, width:350, height:350, borderRadius:"50%", background:`radial-gradient(circle, ${C.gold}18 0%, transparent 70%)` }} />
      <div style={{ position:"relative", width:"100%", maxWidth:400, padding:"36px 30px", borderRadius:18, background:"rgba(255,255,255,0.06)", backdropFilter:"blur(20px)", border:"1px solid rgba(255,255,255,0.1)", boxShadow:"0 24px 48px rgba(0,0,0,0.3)" }}>
        <div style={{ textAlign:"center", marginBottom:24 }}>
          <div style={{ width:56, height:56, borderRadius:14, background:`linear-gradient(135deg, ${C.gold}, ${C.goldLight})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:26, margin:"0 auto 12px", boxShadow:`0 4px 16px ${C.gold}40` }}>🎓</div>
          <h1 style={{ color:"#FFF", fontSize:22, fontWeight:700, margin:0 }}>亞翔學習平台</h1>
          <p style={{ color:"rgba(255,255,255,0.5)", fontSize:12, marginTop:6 }}>L&K Engineering Learning</p>
        </div>

        {/* 第一次登入說明 */}
        <div style={{ marginBottom:18, padding:"10px 12px", background:`${C.gold}15`, borderRadius:8, border:`1px solid ${C.gold}30` }}>
          <p style={{ color:C.goldLight, fontSize:11, margin:0, lineHeight:1.6 }}>
            <strong>💡 首次使用提醒</strong><br />
            預設密碼為您的員工編號（例：E00000）<br />
            登入後請至「個人資料」修改密碼
          </p>
        </div>

        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
          <div>
            <label style={{ display:"block", color:"rgba(255,255,255,0.55)", fontSize:12, marginBottom:4 }}>電子信箱</label>
            <input value={email} onChange={e => setEmail(e.target.value)} onKeyDown={e => e.key==="Enter" && submit()} placeholder="your@lkeng.com" style={{ ...inp, background:"rgba(255,255,255,0.08)", border:"1px solid rgba(255,255,255,0.12)", color:"#FFF" }} />
          </div>
          <div>
            <label style={{ display:"block", color:"rgba(255,255,255,0.55)", fontSize:12, marginBottom:4 }}>密碼</label>
            <input type="password" value={pw} onChange={e => setPw(e.target.value)} onKeyDown={e => e.key==="Enter" && submit()} placeholder="••••••••" style={{ ...inp, background:"rgba(255,255,255,0.08)", border:"1px solid rgba(255,255,255,0.12)", color:"#FFF" }} />
          </div>
          {err && <div style={{ color:"#FCA5A5", fontSize:12, textAlign:"center", padding:7, background:"rgba(220,38,38,0.12)", borderRadius:7 }}>{err}</div>}
          <button onClick={submit} disabled={loading} style={{ width:"100%", padding:12, borderRadius:9, border:"none", background:`linear-gradient(135deg, ${C.gold}, ${C.goldLight})`, color:C.navyDark, fontSize:14, fontWeight:700, cursor:"pointer", marginTop:4, opacity:loading?0.7:1, boxShadow:`0 4px 12px ${C.gold}30` }}>
            {loading ? "登入中..." : "登入"}
          </button>
        </div>
        <div style={{ marginTop:18, padding:10, borderRadius:8, background:"rgba(255,255,255,0.04)" }}>
          <p style={{ color:"rgba(255,255,255,0.4)", fontSize:10, margin:0, textAlign:"center", fontWeight:500 }}>測試帳號（首次登入會被要求改密碼）</p>
          <p style={{ color:"rgba(255,255,255,0.55)", fontSize:10, margin:"4px 0 0", textAlign:"center" }}>管理員：admin@lkeng.com / LK@dmin2026</p>
          <p style={{ color:"rgba(255,255,255,0.55)", fontSize:10, margin:"2px 0 0", textAlign:"center" }}>同仁：chen@lkeng.com / E00002（員工編號）</p>
        </div>
      </div>
    </div>
  );
}

/* ════ 修改密碼彈窗 ════ */
function ChangePasswordModal({ currentUser, updatePassword, onClose, force }) {
  const [oldPw, setOldPw] = useState("");
  const [newPw, setNewPw] = useState("");
  const [confirmPw, setConfirmPw] = useState("");
  const [err, setErr] = useState("");
  const [success, setSuccess] = useState(false);

  const submit = () => {
    setErr("");
    if (oldPw !== currentUser.password) { setErr("舊密碼錯誤"); return; }
    if (newPw.length < 6) { setErr("新密碼至少 6 個字元"); return; }
    if (newPw !== confirmPw) { setErr("兩次輸入的密碼不一致"); return; }
    if (newPw === currentUser.empNo) { setErr("新密碼不可與員工編號相同"); return; }
    updatePassword(newPw);
    setSuccess(true);
    setTimeout(() => onClose(), 1500);
  };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:20 }}>
      <div style={{ background:"#FFF", borderRadius:12, padding:28, width:"100%", maxWidth:380, boxShadow:"0 24px 48px rgba(0,0,0,0.3)" }}>
        <h3 style={{ fontSize:18, fontWeight:700, color:C.text, margin:"0 0 6px" }}>🔒 修改密碼</h3>
        {force ? (
          <p style={{ fontSize:12, color:C.warning, marginBottom:16, padding:"8px 10px", background:`${C.warning}10`, borderRadius:6 }}>
            ⚠️ 您目前使用的是預設密碼，為了帳號安全，請先修改密碼後再繼續使用。
          </p>
        ) : (
          <p style={{ fontSize:12, color:C.textLight, marginBottom:16 }}>請輸入舊密碼及新密碼</p>
        )}
        {success ? (
          <div style={{ textAlign:"center", padding:20 }}>
            <span style={{ fontSize:42 }}>✅</span>
            <p style={{ fontSize:14, color:C.success, fontWeight:600, marginTop:10 }}>密碼修改成功！</p>
          </div>
        ) : (
          <>
            <Field label="舊密碼"><input type="password" value={oldPw} onChange={e => setOldPw(e.target.value)} style={inp} placeholder="目前的密碼" /></Field>
            <Field label="新密碼（至少 6 字元）"><input type="password" value={newPw} onChange={e => setNewPw(e.target.value)} style={inp} placeholder="新密碼" /></Field>
            <Field label="確認新密碼"><input type="password" value={confirmPw} onChange={e => setConfirmPw(e.target.value)} onKeyDown={e => e.key==="Enter" && submit()} style={inp} placeholder="再次輸入新密碼" /></Field>
            {err && <p style={{ color:C.danger, fontSize:12, margin:"6px 0", textAlign:"center" }}>{err}</p>}
            <div style={{ display:"flex", gap:8, marginTop:14, justifyContent:"flex-end" }}>
              {!force && <Btn onClick={onClose} variant="outline">取消</Btn>}
              <Btn onClick={submit} variant="gold">確認修改</Btn>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function Front({ categories, courses, currentUser, onLogout, watchHistory, recordWatch, quizResults, saveQuiz, setView, updatePassword }) {
  const [page, setPage] = useState("home");
  const [selectedCourse, setSelectedCourse] = useState(null);
  const [search, setSearch] = useState("");
  const [filterCat, setFilterCat] = useState("all");
  const [showChat, setShowChat] = useState(false);
  const [showPwModal, setShowPwModal] = useState(currentUser.mustChangePw || false);

  const userHistory = useMemo(() => Object.entries(watchHistory).filter(([k]) => k.startsWith(currentUser.id+"_")).map(([,v]) => v), [watchHistory, currentUser]);
  const userQuizzes = useMemo(() => Object.entries(quizResults).filter(([k]) => k.startsWith(currentUser.id+"_")).map(([k,v]) => ({...v, courseId:k.split("_")[1]})), [quizResults, currentUser]);
  const filtered = useMemo(() => {
    let l = courses;
    if (filterCat !== "all") l = l.filter(c => c.category===filterCat);
    if (search) l = l.filter(c => c.title.includes(search) || c.description.includes(search) || c.instructor.includes(search));
    return l;
  }, [courses, filterCat, search]);
  const newest = useMemo(() => [...courses].sort((a,b) => b.publishDate.localeCompare(a.publishDate)).slice(0,4), [courses]);
  const popular = useMemo(() => [...courses].sort((a,b) => b.views-a.views).slice(0,4), [courses]);
  const inProgress = useMemo(() => userHistory.filter(h => h.progress<100).map(h => ({...h, course:courses.find(c => c.id===h.courseId)})).filter(h => h.course), [userHistory, courses]);

  const Card = ({ course }) => {
    const cat = categories.find(c => c.id===course.category);
    return (
      <div onClick={() => { setSelectedCourse(course); setPage("course"); }} style={{ background:"#FFF", borderRadius:10, overflow:"hidden", cursor:"pointer", border:`1px solid ${C.border}`, transition:"all 0.2s" }}
        onMouseOver={e => { e.currentTarget.style.transform="translateY(-2px)"; e.currentTarget.style.boxShadow="0 6px 16px rgba(0,0,0,0.08)"; }}
        onMouseOut={e => { e.currentTarget.style.transform="none"; e.currentTarget.style.boxShadow="none"; }}>
        <div style={{ height:90, background:`linear-gradient(135deg, ${cat?.color||C.navy}15, ${cat?.color||C.navy}30)`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:32 }}>{course.thumbnail}</div>
        <div style={{ padding:12 }}>
          <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:5 }}>
            <span style={{ fontSize:10, padding:"2px 6px", borderRadius:4, background:`${cat?.color||C.navy}12`, color:cat?.color||C.navy, fontWeight:500 }}>{cat?.name||"未分類"}</span>
            <span style={{ fontSize:10, color:C.textLight }}>👁 {course.views}</span>
          </div>
          <h3 style={{ margin:0, fontSize:13, fontWeight:600, color:C.text, lineHeight:1.4 }}>{course.title}</h3>
          <p style={{ margin:"4px 0 0", fontSize:11, color:C.textLight }}>{course.instructor} · {course.duration}分</p>
        </div>
      </div>
    );
  };

  return (
    <div style={{ minHeight:"100vh", background:C.bg, fontFamily:"'Noto Sans TC',sans-serif" }}>
      {showPwModal && <ChangePasswordModal currentUser={currentUser} updatePassword={updatePassword} onClose={() => setShowPwModal(false)} force={currentUser.mustChangePw} />}

      <div style={{ background:"#FFF", borderBottom:`2px solid ${C.gold}40`, padding:"0 20px", display:"flex", alignItems:"center", height:56, gap:12, position:"sticky", top:0, zIndex:100, boxShadow:"0 1px 4px rgba(0,0,0,0.04)", flexWrap:"wrap" }}>
        <div onClick={() => { setSelectedCourse(null); setPage("home"); }} style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer" }}>
          <div style={{ width:30, height:30, borderRadius:7, background:`linear-gradient(135deg, ${C.gold}, ${C.goldLight})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:15 }}>🎓</div>
          <span style={{ fontSize:15, fontWeight:700, color:C.navy }}>亞翔學習</span>
        </div>
        <div style={{ flex:1, maxWidth:300, minWidth:120 }}>
          <input value={search} onChange={e => { setSearch(e.target.value); if (page !== "courses") setPage("courses"); }} placeholder="🔍 搜尋..." style={{ width:"100%", padding:"6px 12px", borderRadius:16, border:`1px solid ${C.border}`, fontSize:12, outline:"none", boxSizing:"border-box", background:C.bg }} />
        </div>
        <div style={{ display:"flex", gap:2 }}>
          {[{l:"首頁",p:"home"},{l:"課程",p:"courses"},{l:"我的",p:"profile"}].map(n => (
            <button key={n.p} onClick={() => setPage(n.p)} style={{ padding:"6px 10px", borderRadius:6, border:"none", background:page===n.p?`${C.navy}10`:"transparent", color:page===n.p?C.navy:C.textLight, fontWeight:page===n.p?600:400, fontSize:12, cursor:"pointer" }}>{n.l}</button>
          ))}
        </div>
        {currentUser.role==="admin" && <Btn onClick={() => setView("admin")} variant="outline" style={{ padding:"4px 10px", fontSize:11 }}>後台</Btn>}
        <div style={{ width:28, height:28, borderRadius:"50%", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, display:"flex", alignItems:"center", justifyContent:"center", color:"#FFF", fontSize:12, fontWeight:600 }}>{currentUser.name[0]}</div>
        <Btn onClick={onLogout} variant="ghost" style={{ fontSize:11, color:C.textLight, padding:"4px 8px" }}>登出</Btn>
      </div>

      {page==="home" && (
        <div>
          <div style={{ background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, padding:"36px 24px", color:"#FFF", position:"relative", overflow:"hidden" }}>
            <div style={{ position:"absolute", right:-30, top:-30, width:180, height:180, borderRadius:"50%", background:`${C.gold}12` }} />
            <h1 style={{ fontSize:22, fontWeight:700, margin:0, position:"relative" }}>歡迎回來，{currentUser.name} 👋</h1>
            <p style={{ fontSize:13, opacity:0.7, marginTop:6, position:"relative" }}>持續學習，提升專業技能</p>
            <div style={{ display:"flex", gap:10, marginTop:18, position:"relative", flexWrap:"wrap" }}>
              {[{l:"已完成",v:userHistory.filter(h=>h.progress>=100).length},{l:"進行中",v:inProgress.length},{l:"學習時數",v:`${userHistory.reduce((s,h)=>s+(h.totalTime||0),0)}分`}].map(s => (
                <div key={s.l} style={{ padding:"8px 14px", borderRadius:8, background:"rgba(255,255,255,0.1)" }}>
                  <span style={{ fontSize:11, opacity:0.6 }}>{s.l}</span>
                  <span style={{ display:"block", fontSize:18, fontWeight:700 }}>{s.v}</span>
                </div>
              ))}
            </div>
          </div>
          <div style={{ padding:"24px 20px" }}>
            {inProgress.length > 0 && <div style={{ marginBottom:28 }}>
              <h2 style={{ fontSize:16, fontWeight:700, color:C.text, marginBottom:12 }}>▶ 繼續觀看</h2>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(220px,1fr))", gap:10 }}>
                {inProgress.map(h => (
                  <div key={h.courseId} onClick={() => { setSelectedCourse(h.course); setPage("course"); }} style={{ background:"#FFF", borderRadius:9, padding:12, cursor:"pointer", border:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:10 }}>
                    <span style={{ fontSize:24 }}>{h.course.thumbnail}</span>
                    <div style={{ flex:1, minWidth:0 }}>
                      <p style={{ margin:0, fontSize:12, fontWeight:600, color:C.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{h.course.title}</p>
                      <div style={{ marginTop:5, height:3, borderRadius:2, background:C.border }}><div style={{ height:"100%", borderRadius:2, background:`linear-gradient(90deg, ${C.navy}, ${C.gold})`, width:`${h.progress}%` }} /></div>
                      <p style={{ margin:"3px 0 0", fontSize:10, color:C.textLight }}>{h.progress}%</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>}
            <div style={{ marginBottom:28 }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                <h2 style={{ fontSize:16, fontWeight:700, color:C.text, margin:0 }}>🆕 最新上架</h2>
                <button onClick={() => setPage("courses")} style={{ border:"none", background:"none", color:C.navy, fontSize:12, cursor:"pointer", fontWeight:500 }}>查看全部 →</button>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(170px,1fr))", gap:12 }}>{newest.map(c => <Card key={c.id} course={c} />)}</div>
            </div>
            <div>
              <h2 style={{ fontSize:16, fontWeight:700, color:C.text, marginBottom:12 }}>🔥 最熱門課程</h2>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(170px,1fr))", gap:12 }}>{popular.map(c => <Card key={c.id} course={c} />)}</div>
            </div>
          </div>
        </div>
      )}

      {page==="courses" && (
        <div style={{ padding:"24px 20px" }}>
          <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:14 }}>所有課程</h2>
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:18 }}>
            <button onClick={() => setFilterCat("all")} style={{ padding:"4px 12px", borderRadius:14, border:`1px solid ${filterCat==="all"?C.navy:C.border}`, background:filterCat==="all"?`${C.navy}10`:"#FFF", color:filterCat==="all"?C.navy:C.textMid, fontSize:11, cursor:"pointer", fontWeight:500 }}>全部</button>
            {categories.map(cat => (
              <button key={cat.id} onClick={() => setFilterCat(cat.id)} style={{ padding:"4px 12px", borderRadius:14, border:`1px solid ${filterCat===cat.id?cat.color:C.border}`, background:filterCat===cat.id?`${cat.color}10`:"#FFF", color:filterCat===cat.id?cat.color:C.textMid, fontSize:11, cursor:"pointer", fontWeight:500 }}>{cat.icon} {cat.name}</button>
            ))}
          </div>
          {filtered.length===0 ? <div style={{ textAlign:"center", padding:36, color:C.textLight }}><p style={{ fontSize:30 }}>🔍</p><p>沒有找到符合條件的課程</p></div> : <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(170px,1fr))", gap:12 }}>{filtered.map(c => <Card key={c.id} course={c} />)}</div>}
        </div>
      )}

      {page==="profile" && (
        <div style={{ padding:"24px 20px" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:18, flexWrap:"wrap", gap:10 }}>
            <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>我的學習</h2>
            <Btn onClick={() => setShowPwModal(true)} variant="outline" style={{ fontSize:12 }}>🔒 修改密碼</Btn>
          </div>

          {/* 個人資料卡 */}
          <div style={{ background:"#FFF", borderRadius:10, padding:16, border:`1px solid ${C.border}`, marginBottom:16, display:"flex", alignItems:"center", gap:14, flexWrap:"wrap" }}>
            <div style={{ width:54, height:54, borderRadius:"50%", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, display:"flex", alignItems:"center", justifyContent:"center", color:"#FFF", fontSize:22, fontWeight:600 }}>{currentUser.name[0]}</div>
            <div style={{ flex:1, minWidth:200 }}>
              <p style={{ margin:0, fontSize:15, fontWeight:600, color:C.text }}>{currentUser.name}</p>
              <p style={{ margin:"2px 0 0", fontSize:12, color:C.textLight }}>{currentUser.empNo} · {currentUser.department} · {currentUser.email}</p>
            </div>
          </div>

          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(160px,1fr))", gap:12, marginBottom:24 }}>
            {[{l:"觀看課程",v:userHistory.length,i:"📚"},{l:"學習時數",v:`${userHistory.reduce((s,h)=>s+(h.totalTime||0),0)} 分`,i:"⏱️"},{l:"測驗完成",v:userQuizzes.length,i:"📝"}].map(s => (
              <div key={s.l} style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:10 }}>
                <span style={{ fontSize:24 }}>{s.i}</span>
                <div><p style={{ margin:0, fontSize:11, color:C.textLight }}>{s.l}</p><p style={{ margin:"3px 0 0", fontSize:18, fontWeight:700, color:C.text }}>{s.v}</p></div>
              </div>
            ))}
          </div>
          <h3 style={{ fontSize:14, fontWeight:600, color:C.text, marginBottom:8 }}>瀏覽紀錄</h3>
          {userHistory.length===0 ? <p style={{ color:C.textLight, fontSize:12 }}>尚無觀看紀錄</p> : userHistory.map(h => {
            const c = courses.find(cc => cc.id===h.courseId); if (!c) return null;
            return <div key={h.courseId} onClick={() => { setSelectedCourse(c); setPage("course"); }} style={{ background:"#FFF", borderRadius:8, padding:10, border:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:10, cursor:"pointer", marginBottom:6 }}>
              <span style={{ fontSize:22 }}>{c.thumbnail}</span>
              <div style={{ flex:1 }}><p style={{ margin:0, fontSize:12, fontWeight:600 }}>{c.title}</p><p style={{ margin:"2px 0 0", fontSize:10, color:C.textLight }}>學習 {h.totalTime} 分 · 進度 {h.progress}%</p></div>
              <span style={{ fontSize:11, color:h.progress>=100?C.success:C.navy, fontWeight:500 }}>{h.progress>=100?"✅":"▶"}</span>
            </div>;
          })}
          {userQuizzes.length > 0 && <>
            <h3 style={{ fontSize:14, fontWeight:600, color:C.text, margin:"20px 0 8px" }}>測驗紀錄</h3>
            {userQuizzes.map((q,i) => { const c = courses.find(cc=>cc.id===q.courseId); const pct = Math.round(q.score/q.total*100);
              return <div key={i} style={{ background:"#FFF", borderRadius:8, padding:10, border:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:10, marginBottom:6 }}>
                <span style={{ fontSize:18 }}>{pct>=60?"🎉":"📖"}</span>
                <div style={{ flex:1 }}><p style={{ margin:0, fontSize:12, fontWeight:600 }}>{c?.title||"未知"}</p><p style={{ margin:"2px 0 0", fontSize:10, color:C.textLight }}>{q.score}/{q.total}（{pct}%）</p></div>
                <span style={{ fontSize:11, padding:"2px 7px", borderRadius:5, background:pct>=60?`${C.success}12`:`${C.danger}12`, color:pct>=60?C.success:C.danger, fontWeight:500 }}>{pct>=60?"通過":"未通過"}</span>
              </div>;
            })}
          </>}
        </div>
      )}

      {page==="course" && selectedCourse && <CoursePage {...{categories,course:selectedCourse,goBack:()=>{setSelectedCourse(null);setPage("home");},watchHistory,currentUser,recordWatch,saveQuiz}} />}

      <button onClick={() => setShowChat(!showChat)} style={{ position:"fixed", bottom:18, right:18, width:48, height:48, borderRadius:"50%", border:"none", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, color:C.goldLight, fontSize:20, cursor:"pointer", boxShadow:`0 4px 14px ${C.navy}40`, zIndex:200 }}>{showChat?"✕":"💬"}</button>
      {showChat && <ChatBot onClose={() => setShowChat(false)} courses={courses} categories={categories} />}
    </div>
  );
}

function CoursePage({ categories, course, goBack, watchHistory, currentUser, recordWatch, saveQuiz }) {
  const [activeCh, setActiveCh] = useState(0);
  const [showQuiz, setShowQuiz] = useState(false);
  const progress = watchHistory[`${currentUser.id}_${course.id}`]?.progress || 0;
  const cat = categories.find(c => c.id===course.category);
  const currentChapter = course.chapters[activeCh];
  const videoId = getYouTubeId(currentChapter?.youtubeUrl);

  // 模擬觀看進度（章節切換時記錄）
  useEffect(() => {
    if (videoId) {
      const op = Math.min(100, Math.round((activeCh + 1)/course.chapters.length*100));
      recordWatch(course.id, activeCh, op);
    }
  }, [activeCh, videoId]);

  if (showQuiz) return <Quiz course={course} goBack={() => setShowQuiz(false)} saveQuiz={saveQuiz} />;

  return (
    <div style={{ padding:"24px 20px" }}>
      <button onClick={goBack} style={{ border:"none", background:"none", color:C.navy, fontSize:12, cursor:"pointer", padding:0, fontWeight:500, marginBottom:12 }}>← 返回</button>
      <div style={{ display:"flex", gap:18, flexWrap:"wrap" }}>
        <div style={{ flex:"1 1 380px", minWidth:0 }}>
          {/* YouTube 嵌入 */}
          <div style={{ background:"#000", borderRadius:10, aspectRatio:"16/9", overflow:"hidden", position:"relative" }}>
            {videoId ? (
              <iframe
                key={videoId}
                src={`https://www.youtube.com/embed/${videoId}?rel=0&modestbranding=1`}
                title={currentChapter.title}
                style={{ width:"100%", height:"100%", border:"none" }}
                allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
                allowFullScreen
              />
            ) : (
              <div style={{ width:"100%", height:"100%", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", color:"#FFF", background:"#1a1a1a" }}>
                <span style={{ fontSize:42 }}>📹</span>
                <p style={{ fontSize:13, opacity:0.7, margin:"8px 0 0" }}>本章節尚未提供影片</p>
                <p style={{ fontSize:11, opacity:0.4, margin:"4px 0 0" }}>請聯絡管理員上傳影片</p>
              </div>
            )}
          </div>
          {videoId && (
            <div style={{ marginTop:8, padding:"8px 12px", background:`${C.gold}10`, borderRadius:6, fontSize:11, color:C.textMid, lineHeight:1.5 }}>
              💡 若影片無法播放（錯誤 153）：請在 YouTube 後台確認影片設為「<strong>公開</strong>」或「<strong>不公開</strong>」，並於影片設定中勾選「<strong>允許嵌入</strong>」。
              不公開影片需透過 https:// 網址才能嵌入，本地檔案開啟時可能無法播放，部署到伺服器後即可正常運作。
              <a href={`https://www.youtube.com/watch?v=${videoId}`} target="_blank" rel="noopener noreferrer" style={{ color:C.accent, marginLeft:6 }}>→ 直接到 YouTube 觀看</a>
            </div>
          )}
          <div style={{ marginTop:14 }}>
            <span style={{ fontSize:10, padding:"3px 7px", borderRadius:4, background:`${cat?.color||C.navy}12`, color:cat?.color||C.navy, fontWeight:500 }}>{cat?.name||"未分類"}</span>
            <span style={{ fontSize:11, color:C.textLight, marginLeft:8 }}>👁 {course.views}</span>
            <h1 style={{ fontSize:20, fontWeight:700, color:C.text, margin:"8px 0 4px" }}>{course.title}</h1>
            <p style={{ fontSize:13, color:C.navy, margin:"4px 0 0", fontWeight:500 }}>目前章節：{currentChapter?.title}</p>
            <p style={{ fontSize:12, color:C.textMid, margin:"4px 0 0" }}>講師：{course.instructor} · {course.duration} 分鐘</p>
            <p style={{ fontSize:13, color:C.text, marginTop:10, lineHeight:1.7 }}>{course.description}</p>
            {course.files?.length > 0 && (
              <div style={{ marginTop:12, padding:12, background:C.goldPale, borderRadius:8, border:`1px solid ${C.gold}30` }}>
                <p style={{ fontSize:12, fontWeight:600, color:C.navy, margin:"0 0 6px" }}>📎 課程附件</p>
                {course.files.map((f,i) => <div key={i} style={{ fontSize:11, color:C.accent, padding:"3px 0", cursor:"pointer" }}>📄 {f.name} <span style={{ color:C.textLight }}>({f.size})</span></div>)}
              </div>
            )}
          </div>
        </div>
        <div style={{ flex:"0 0 250px" }}>
          <div style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}`, marginBottom:10 }}>
            <p style={{ margin:0, fontSize:12, fontWeight:600 }}>學習進度</p>
            <div style={{ marginTop:7, height:5, borderRadius:3, background:C.border }}><div style={{ height:"100%", borderRadius:3, background:`linear-gradient(90deg, ${C.navy}, ${C.gold})`, width:`${progress}%`, transition:"width 0.5s" }} /></div>
            <p style={{ margin:"3px 0 0", fontSize:11, color:C.textLight, textAlign:"right" }}>{progress}%</p>
          </div>
          <div style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}`, marginBottom:10 }}>
            <p style={{ margin:"0 0 8px", fontSize:12, fontWeight:600 }}>課程章節</p>
            {course.chapters.map((ch,i) => {
              const hasVideo = !!getYouTubeId(ch.youtubeUrl);
              return (
                <div key={i} onClick={() => setActiveCh(i)} style={{ padding:"6px 8px", borderRadius:5, cursor:"pointer", display:"flex", alignItems:"center", gap:7, background:i===activeCh?`${C.navy}08`:"transparent", marginBottom:2 }}>
                  <span style={{ width:20, height:20, borderRadius:"50%", background:i===activeCh?C.navy:C.border, color:i===activeCh?"#FFF":C.textLight, display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, fontWeight:600, flexShrink:0 }}>{i+1}</span>
                  <div style={{ flex:1, minWidth:0 }}>
                    <p style={{ margin:0, fontSize:11, fontWeight:i===activeCh?600:400, color:i===activeCh?C.navy:C.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{ch.title}</p>
                    <p style={{ margin:"1px 0 0", fontSize:9, color:C.textLight }}>{ch.duration} 分鐘 {hasVideo ? "🎬" : "⏸"}</p>
                  </div>
                </div>
              );
            })}
          </div>
          {course.quiz?.length > 0 && (
            <button onClick={() => setShowQuiz(true)} style={{ width:"100%", padding:11, borderRadius:9, border:"none", background:`linear-gradient(135deg, ${C.gold}, ${C.goldLight})`, color:C.navyDark, fontSize:13, fontWeight:600, cursor:"pointer", boxShadow:`0 2px 8px ${C.gold}30` }}>
              📝 開始課後測驗（{course.quiz.length} 題）
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function Quiz({ course, goBack, saveQuiz }) {
  const [answers, setAnswers] = useState({});
  const [submitted, setSubmitted] = useState(false);
  const score = submitted ? course.quiz.reduce((s,q,i) => s + (answers[i]===q.answer?1:0), 0) : 0;
  const pct = Math.round(score/course.quiz.length*100);

  const submit = () => { let s=0; course.quiz.forEach((q,i) => { if(answers[i]===q.answer) s++; }); saveQuiz(course.id,s,course.quiz.length); setSubmitted(true); };

  return (
    <div style={{ padding:"24px 20px", maxWidth:640, margin:"0 auto" }}>
      <button onClick={goBack} style={{ border:"none", background:"none", color:C.navy, fontSize:12, cursor:"pointer", padding:0, fontWeight:500, marginBottom:12 }}>← 返回課程</button>
      <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:4 }}>📝 {course.title} — 課後測驗</h2>
      <p style={{ color:C.textLight, fontSize:12, marginBottom:18 }}>共 {course.quiz.length} 題，及格 60%</p>
      {submitted ? (
        <div style={{ textAlign:"center", padding:32, background:"#FFF", borderRadius:12, border:`1px solid ${C.border}` }}>
          <span style={{ fontSize:42 }}>{pct>=60?"🎉":"📖"}</span>
          <h3 style={{ fontSize:20, fontWeight:700, color:C.text, marginTop:12 }}>{pct>=60?"恭喜通過！":"再接再厲！"}</h3>
          <p style={{ fontSize:15, color:C.textMid, marginTop:4 }}>得分：{score}/{course.quiz.length}（{pct}%）</p>
          <div style={{ display:"flex", gap:8, justifyContent:"center", marginTop:16 }}>
            <Btn onClick={() => { setAnswers({}); setSubmitted(false); }} variant="outline">重新測驗</Btn>
            <Btn onClick={goBack}>返回課程</Btn>
          </div>
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          {course.quiz.map((q,qi) => (
            <div key={qi} style={{ background:"#FFF", borderRadius:9, padding:16, border:`1px solid ${C.border}` }}>
              <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>{qi+1}. {q.q}</p>
              <div style={{ display:"flex", flexDirection:"column", gap:5, marginTop:8 }}>
                {q.options.map((opt,oi) => (
                  <label key={oi} onClick={() => setAnswers(p=>({...p,[qi]:oi}))} style={{ display:"flex", alignItems:"center", gap:7, padding:"7px 10px", borderRadius:5, cursor:"pointer", border:`1px solid ${answers[qi]===oi?C.navy:C.border}`, background:answers[qi]===oi?`${C.navy}06`:"transparent" }}>
                    <div style={{ width:14, height:14, borderRadius:"50%", border:`2px solid ${answers[qi]===oi?C.navy:C.border}`, display:"flex", alignItems:"center", justifyContent:"center" }}>{answers[qi]===oi && <div style={{ width:6, height:6, borderRadius:"50%", background:C.navy }} />}</div>
                    <span style={{ fontSize:12, color:C.text }}>{opt}</span>
                  </label>
                ))}
              </div>
            </div>
          ))}
          <Btn onClick={submit} disabled={Object.keys(answers).length<course.quiz.length} style={{ alignSelf:"center", padding:"10px 28px" }}>提交答案</Btn>
        </div>
      )}
    </div>
  );
}

function ChatBot({ onClose, courses, categories }) {
  const [msgs, setMsgs] = useState([{ role:"assistant", content:"你好！我是亞翔學習助手 🤖\n有任何關於課程的問題都可以問我！" }]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const btmRef = useRef(null);
  useEffect(() => { btmRef.current?.scrollIntoView({ behavior:"smooth" }); }, [msgs]);

  const send = async () => {
    if (!input.trim()||loading) return;
    const msg = input.trim(); setInput(""); setMsgs(p => [...p, {role:"user",content:msg}]); setLoading(true);
    const ctx = courses.map(c => `[${c.title}] 分類:${categories.find(cat=>cat.id===c.category)?.name} 講師:${c.instructor}`).join("\n");
    try {
      const r = await fetch("https://api.anthropic.com/v1/messages", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ model:"claude-sonnet-4-20250514", max_tokens:800, system:`你是亞翔工程企業內部學習平台的AI助手。課程：\n${ctx}\n\n簡潔友善，使用繁體中文。`, messages:[{role:"user",content:msg}] }) });
      const d = await r.json();
      setMsgs(p => [...p, {role:"assistant", content: d.content?.map(b=>b.text||"").join("")||"抱歉，暫時無法回應。"}]);
    } catch {
      const matched = courses.filter(c => c.title.includes(msg)||c.description.includes(msg));
      const reply = matched.length > 0 ? matched.map(c=>`📘 ${c.title}（${c.instructor}）`).join("\n") : `目前平台共有 ${courses.length} 門課程，涵蓋${categories.map(c=>c.name).join("、")}等類別。`;
      setMsgs(p => [...p, {role:"assistant", content:reply}]);
    }
    setLoading(false);
  };

  return (
    <div style={{ position:"fixed", bottom:74, right:18, width:340, height:450, borderRadius:12, background:"#FFF", boxShadow:"0 10px 32px rgba(0,0,0,0.16)", display:"flex", flexDirection:"column", zIndex:200, overflow:"hidden", border:`1px solid ${C.border}` }}>
      <div style={{ padding:"10px 14px", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, color:"#FFF", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
        <span style={{ fontSize:13, fontWeight:600 }}>🤖 學習助手</span>
        <button onClick={onClose} style={{ border:"none", background:"none", color:"#FFF", fontSize:15, cursor:"pointer" }}>✕</button>
      </div>
      <div style={{ flex:1, overflowY:"auto", padding:10, display:"flex", flexDirection:"column", gap:7 }}>
        {msgs.map((m,i) => <div key={i} style={{ alignSelf:m.role==="user"?"flex-end":"flex-start", maxWidth:"80%", padding:"7px 11px", borderRadius:9, background:m.role==="user"?C.navy:C.bg, color:m.role==="user"?"#FFF":C.text, fontSize:12, lineHeight:1.6, whiteSpace:"pre-wrap" }}>{m.content}</div>)}
        {loading && <div style={{ alignSelf:"flex-start", padding:"7px 11px", borderRadius:9, background:C.bg, color:C.textLight, fontSize:12 }}>思考中...</div>}
        <div ref={btmRef} />
      </div>
      <div style={{ padding:8, borderTop:`1px solid ${C.border}`, display:"flex", gap:5 }}>
        <input value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => e.key==="Enter"&&send()} placeholder="輸入訊息..." style={{ flex:1, padding:"7px 10px", borderRadius:7, border:`1px solid ${C.border}`, fontSize:12, outline:"none" }} />
        <Btn onClick={send} disabled={loading} style={{ padding:"7px 12px", fontSize:12 }}>發送</Btn>
      </div>
    </div>
  );
}

/* ════ ADMIN PANEL ════ */
function Admin({ categories, setCategories, courses, setCourses, users, setUsers, watchHistory, quizResults, onLogout, setView, currentUser, updatePassword }) {
  const [tab, setTab] = useState("dashboard");
  const [showPwModal, setShowPwModal] = useState(false);

  const tabs = [
    { id:"dashboard", label:"總覽", icon:"📊" },
    { id:"courses", label:"課程管理", icon:"📚" },
    { id:"categories", label:"分類管理", icon:"🏷️" },
    { id:"users", label:"使用者管理", icon:"👥" },
    { id:"analytics", label:"學習分析", icon:"📈" },
    { id:"quizzes", label:"測驗紀錄", icon:"📝" },
  ];

  return (
    <div style={{ display:"flex", minHeight:"100vh", background:C.bg, fontFamily:"'Noto Sans TC',sans-serif" }}>
      {showPwModal && <ChangePasswordModal currentUser={currentUser} updatePassword={updatePassword} onClose={() => setShowPwModal(false)} />}
      <div style={{ width:200, background:C.navy, color:"#FFF", display:"flex", flexDirection:"column", flexShrink:0 }}>
        <div style={{ padding:"16px 14px", display:"flex", alignItems:"center", gap:8, borderBottom:"1px solid rgba(255,255,255,0.08)" }}>
          <div style={{ width:30, height:30, borderRadius:7, background:`linear-gradient(135deg, ${C.gold}, ${C.goldLight})`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:15 }}>🎓</div>
          <div><p style={{ margin:0, fontSize:13, fontWeight:600 }}>管理後台</p><p style={{ margin:0, fontSize:9, opacity:0.4 }}>L&K Admin</p></div>
        </div>
        <div style={{ flex:1, padding:"8px 6px" }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{ width:"100%", padding:"8px 10px", borderRadius:5, border:"none", background:tab===t.id?"rgba(255,255,255,0.12)":"transparent", color:tab===t.id?"#FFF":"rgba(255,255,255,0.5)", fontSize:12, cursor:"pointer", textAlign:"left", display:"flex", alignItems:"center", gap:7, marginBottom:2, fontWeight:tab===t.id?500:400 }}>
              <span style={{ fontSize:13 }}>{t.icon}</span>{t.label}
            </button>
          ))}
        </div>
        <div style={{ padding:"10px 8px", borderTop:"1px solid rgba(255,255,255,0.08)" }}>
          <button onClick={() => setShowPwModal(true)} style={{ width:"100%", padding:"6px 8px", borderRadius:5, border:"1px solid rgba(255,255,255,0.15)", background:"transparent", color:"rgba(255,255,255,0.6)", fontSize:11, cursor:"pointer", marginBottom:4 }}>🔒 修改密碼</button>
          <button onClick={() => setView("front")} style={{ width:"100%", padding:"6px 8px", borderRadius:5, border:"1px solid rgba(255,255,255,0.15)", background:"transparent", color:"rgba(255,255,255,0.6)", fontSize:11, cursor:"pointer", marginBottom:4 }}>前往前台 →</button>
          <button onClick={onLogout} style={{ width:"100%", padding:"6px 8px", borderRadius:5, border:"none", background:"rgba(220,38,38,0.12)", color:"#FCA5A5", fontSize:11, cursor:"pointer" }}>登出</button>
        </div>
      </div>

      <div style={{ flex:1, padding:"20px 24px", overflowY:"auto", minWidth:0 }}>
        {tab==="dashboard" && <Dashboard {...{courses,users,categories}} />}
        {tab==="courses" && <CourseAdmin {...{categories,courses,setCourses}} />}
        {tab==="categories" && <CategoryAdmin {...{categories,setCategories,courses}} />}
        {tab==="users" && <UserAdmin {...{users,setUsers}} />}
        {tab==="analytics" && <Analytics courses={courses} />}
        {tab==="quizzes" && <QuizRecords {...{quizResults,users,courses}} />}
      </div>
    </div>
  );
}

function Dashboard({ courses, users, categories }) {
  const totalViews = courses.reduce((s,c) => s+c.views, 0);
  const published = courses.filter(c => c.status==="published").length;
  return (
    <div>
      <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:16 }}>儀表板總覽</h2>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(160px,1fr))", gap:12, marginBottom:20 }}>
        {[{l:"總課程數",v:courses.length,i:"📚",c:C.navy},{l:"已上架",v:published,i:"✅",c:C.success},{l:"使用者",v:users.length,i:"👥",c:C.accent},{l:"總瀏覽",v:totalViews,i:"👁",c:C.gold}].map(s => (
          <div key={s.l} style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}` }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}><span style={{ fontSize:22 }}>{s.i}</span><span style={{ fontSize:10, color:s.c, padding:"2px 6px", borderRadius:4, background:`${s.c}12` }}>{s.l}</span></div>
            <p style={{ fontSize:22, fontWeight:700, color:C.text, margin:"6px 0 0" }}>{s.v}</p>
          </div>
        ))}
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(300px,1fr))", gap:12 }}>
        <div style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}` }}>
          <h3 style={{ color:C.text, fontSize:14, fontWeight:600, marginBottom:10 }}>🔥 熱門課程 TOP 5</h3>
          {[...courses].sort((a,b)=>b.views-a.views).slice(0,5).map((c,i) => (
            <div key={c.id} style={{ display:"flex", alignItems:"center", gap:7, padding:"5px 0", borderBottom:i<4?`1px solid ${C.border}`:"none" }}>
              <span style={{ width:18, height:18, borderRadius:"50%", background:i<3?C.navy:`${C.navy}20`, color:i<3?"#FFF":C.navy, display:"flex", alignItems:"center", justifyContent:"center", fontSize:9, fontWeight:600 }}>{i+1}</span>
              <span style={{ flex:1, fontSize:12, color:C.text }}>{c.title}</span>
              <span style={{ fontSize:11, color:C.textLight }}>{c.views}</span>
            </div>
          ))}
        </div>
        <div style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}` }}>
          <h3 style={{ color:C.text, fontSize:14, fontWeight:600, marginBottom:10 }}>📂 各系統課程數</h3>
          {categories.map(cat => {
            const n = courses.filter(c => c.category===cat.id).length;
            return <div key={cat.id} style={{ display:"flex", alignItems:"center", gap:7, padding:"4px 0" }}>
              <span style={{ fontSize:13 }}>{cat.icon}</span>
              <span style={{ flex:1, fontSize:12, color:C.text }}>{cat.name}</span>
              <div style={{ width:80, height:4, borderRadius:2, background:`${C.navy}10` }}><div style={{ height:"100%", borderRadius:2, background:cat.color, width:`${n/Math.max(courses.length,1)*100}%` }} /></div>
              <span style={{ fontSize:11, color:C.textLight, width:18, textAlign:"right" }}>{n}</span>
            </div>;
          })}
        </div>
      </div>
    </div>
  );
}

function CourseAdmin({ categories, courses, setCourses }) {
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [title, setTitle] = useState("");
  const [category, setCategory] = useState("sys1");
  const [instructor, setInstructor] = useState("");
  const [description, setDescription] = useState("");
  const [chapters, setChapters] = useState([{ title:"第一章", duration:15, youtubeUrl:"" }]);

  const reset = () => {
    setTitle(""); setCategory("sys1"); setInstructor(""); setDescription("");
    setChapters([{ title:"第一章", duration:15, youtubeUrl:"" }]);
    setEditing(null); setShowForm(false);
  };
  const startEdit = (c) => {
    setTitle(c.title); setCategory(c.category); setInstructor(c.instructor); setDescription(c.description);
    setChapters(c.chapters.map(ch => ({ ...ch, youtubeUrl: ch.youtubeUrl || "" })));
    setEditing(c.id); setShowForm(true);
  };
  const save = (publishNow = false) => {
    if (!title.trim()) { alert("請輸入課程名稱"); return; }
    const totalDuration = chapters.reduce((s,c) => s + (+c.duration||0), 0);
    if (editing) {
      setCourses(prev => prev.map(c => c.id===editing ? { ...c, title, category, instructor, description, chapters, duration: totalDuration } : c));
    } else {
      setCourses(prev => [...prev, { id:`c${Date.now()}`, title, category, instructor, description, duration: totalDuration, thumbnail:"📘", views:0, publishDate:new Date().toISOString().split("T")[0], status: publishNow ? "published" : "draft", files:[], chapters, quiz:[] }]);
    }
    reset();
  };

  const updateChapter = (idx, key, val) => setChapters(prev => prev.map((c,i) => i===idx ? { ...c, [key]: val } : c));
  const addChapter = () => setChapters(prev => [...prev, { title:`第${prev.length+1}章`, duration:15, youtubeUrl:"" }]);
  const removeChapter = (idx) => setChapters(prev => prev.length > 1 ? prev.filter((_,i) => i!==idx) : prev);

  const uploadFile = (cid) => {
    const inp = document.createElement("input");
    inp.type = "file"; inp.multiple = true;
    inp.accept = ".pdf,.pptx,.ppt,.doc,.docx,.zip,.xlsx,.xls";
    inp.onchange = (e) => {
      const arr = Array.from(e.target.files).map(f => ({ name:f.name, size:f.size>1048576?`${(f.size/1048576).toFixed(1)} MB`:`${(f.size/1024).toFixed(0)} KB`, type:f.type }));
      setCourses(prev => prev.map(c => c.id===cid ? {...c, files:[...(c.files||[]), ...arr]} : c));
    };
    inp.click();
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:10 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>課程管理</h2>
        <Btn onClick={() => { reset(); setShowForm(true); }}>+ 新增課程</Btn>
      </div>

      {showForm && (
        <div style={{ background:"#FFF", borderRadius:9, padding:16, border:`1px solid ${C.border}`, marginBottom:14 }}>
          <h3 style={{ color:C.text, fontSize:14, fontWeight:600, margin:"0 0 12px" }}>{editing?"編輯課程":"新增課程"}</h3>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(180px,1fr))", gap:10 }}>
            <Field label="課程名稱"><input value={title} onChange={e => setTitle(e.target.value)} style={inp} placeholder="例：5S 管理實務" /></Field>
            <Field label="講師"><input value={instructor} onChange={e => setInstructor(e.target.value)} style={inp} placeholder="講師姓名" /></Field>
            <Field label="分類">
              <select value={category} onChange={e => setCategory(e.target.value)} style={inp}>
                {categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </Field>
          </div>
          <Field label="課程說明"><textarea value={description} onChange={e => setDescription(e.target.value)} rows={2} style={{ ...inp, resize:"vertical" }} placeholder="課程內容說明..." /></Field>

          {/* 章節編輯 */}
          <div style={{ marginTop:14, padding:12, background:C.bg, borderRadius:8 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
              <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>📋 章節設定（每章節對應一支 YouTube 影片）</p>
              <Btn onClick={addChapter} variant="outline" style={{ padding:"4px 10px", fontSize:11 }}>+ 新增章節</Btn>
            </div>
            <div style={{ padding:"10px 12px", background:`${C.gold}10`, borderRadius:6, fontSize:11, color:C.navy, marginBottom:10, lineHeight:1.7 }}>
              💡 <strong>YouTube 影片設定步驟：</strong><br />
              <strong>1.</strong> 登入 YouTube → 上傳影片<br />
              <strong>2.</strong> 隱私權選「<strong>不公開</strong>」（拿到連結的人才能看）<br />
              <strong>3.</strong> 進階設定 → 確認「<strong>允許嵌入</strong>」已勾選 ✅<br />
              <strong>4.</strong> 複製影片網址貼到下方欄位（支援 youtu.be/xxx 或 youtube.com/watch?v=xxx）<br />
              <span style={{ color:C.warning }}>⚠️ <strong>注意</strong>：直接在電腦雙擊本檔案測試時，「不公開」影片可能因 YouTube 防盜連機制無法嵌入（錯誤 153）。<u>正式部署到伺服器後就會正常</u>。測試階段可以先設為「公開」。</span>
            </div>
            {chapters.map((ch, idx) => (
              <div key={idx} style={{ background:"#FFF", borderRadius:7, padding:12, marginBottom:8, border:`1px solid ${C.border}` }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                  <span style={{ fontSize:12, fontWeight:600, color:C.navy }}>第 {idx+1} 章</span>
                  {chapters.length > 1 && <button onClick={() => removeChapter(idx)} style={{ border:"none", background:"none", color:C.danger, fontSize:11, cursor:"pointer" }}>移除</button>}
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr", gap:8, marginBottom:8 }}>
                  <input value={ch.title} onChange={e => updateChapter(idx, "title", e.target.value)} placeholder="章節名稱" style={inp} />
                  <input type="number" value={ch.duration} onChange={e => updateChapter(idx, "duration", +e.target.value||0)} placeholder="時長(分)" style={inp} />
                </div>
                <input value={ch.youtubeUrl} onChange={e => updateChapter(idx, "youtubeUrl", e.target.value)} placeholder="🎬 YouTube 影片網址（例：https://youtu.be/xxxxx）" style={inp} />
                {ch.youtubeUrl && (
                  getYouTubeId(ch.youtubeUrl)
                    ? <p style={{ fontSize:10, color:C.success, margin:"4px 0 0" }}>✅ 影片 ID 已偵測：{getYouTubeId(ch.youtubeUrl)}</p>
                    : <p style={{ fontSize:10, color:C.danger, margin:"4px 0 0" }}>⚠️ 網址格式無法識別</p>
                )}
              </div>
            ))}
          </div>

          <div style={{ marginTop:14, padding:10, background:`${C.warning}10`, borderRadius:7, fontSize:11, color:C.warning, lineHeight:1.6 }}>
            💡 <strong>提醒</strong>：建立課程後預設為「草稿」狀態，<u>必須點「上架」才會出現在前台</u>。
          </div>
          <div style={{ display:"flex", gap:6, marginTop:10, flexWrap:"wrap" }}>
            {editing ? (
              <Btn onClick={() => save(false)} variant="gold">儲存變更</Btn>
            ) : (
              <>
                <Btn onClick={() => save(true)} variant="gold">✅ 建立並直接上架</Btn>
                <Btn onClick={() => save(false)} variant="outline">建立草稿（之後再上架）</Btn>
              </>
            )}
            <Btn onClick={reset} variant="outline">取消</Btn>
          </div>
        </div>
      )}

      <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:700 }}>
          <thead><tr style={{ borderBottom:`1px solid ${C.border}` }}>
            {["課程名稱","分類","講師","章節/影片","瀏覽","附件","狀態","操作"].map(h => <th key={h} style={{ padding:"9px 10px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
          </tr></thead>
          <tbody>
            {courses.map(c => {
              const cat = categories.find(cc => cc.id===c.category);
              const videoCount = c.chapters.filter(ch => getYouTubeId(ch.youtubeUrl)).length;
              return (
                <tr key={c.id} style={{ borderBottom:`1px solid ${C.border}` }}>
                  <td style={{ padding:"8px 10px", fontSize:12, color:C.text }}>{c.title}</td>
                  <td style={{ padding:"8px 10px" }}><span style={{ fontSize:10, padding:"2px 6px", borderRadius:4, background:`${cat?.color||C.navy}12`, color:cat?.color||C.navy }}>{cat?.name||"未分類"}</span></td>
                  <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>{c.instructor}</td>
                  <td style={{ padding:"8px 10px", fontSize:11, color:C.textMid }}>{videoCount}/{c.chapters.length} 🎬</td>
                  <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>{c.views}</td>
                  <td style={{ padding:"8px 10px" }}>
                    <button onClick={() => uploadFile(c.id)} style={{ border:`1px dashed ${C.gold}`, background:C.goldPale, color:C.navy, fontSize:10, cursor:"pointer", padding:"2px 7px", borderRadius:4, fontWeight:500 }}>📎 ({(c.files||[]).length})</button>
                  </td>
                  <td style={{ padding:"8px 10px" }}>
                    <span style={{ fontSize:10, padding:"3px 7px", borderRadius:7, background:c.status==="published"?`${C.success}12`:`${C.warning}12`, color:c.status==="published"?C.success:C.warning }}>{c.status==="published"?"已上架":"草稿"}</span>
                  </td>
                  <td style={{ padding:"8px 10px" }}>
                    <div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>
                      <Btn onClick={() => startEdit(c)} variant="outline" style={{ padding:"3px 7px", fontSize:10 }}>編輯</Btn>
                      <Btn onClick={() => setCourses(p => p.map(cc => cc.id===c.id?{...cc, status:cc.status==="published"?"draft":"published"}:cc))} variant="outline" style={{ padding:"3px 7px", fontSize:10 }}>{c.status==="published"?"下架":"上架"}</Btn>
                      <Btn onClick={() => { if(confirm("確定刪除此課程？")) setCourses(p => p.filter(cc => cc.id!==c.id)) }} variant="danger" style={{ padding:"3px 7px", fontSize:10 }}>刪除</Btn>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function CategoryAdmin({ categories, setCategories, courses }) {
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState(null);
  const [name, setName] = useState("");
  const [icon, setIcon] = useState("⚙️");
  const [color, setColor] = useState("#0066CC");
  const [viewMode, setViewMode] = useState("card"); // "card" or "list"

  const reset = () => { setName(""); setIcon("⚙️"); setColor("#0066CC"); setEditId(null); setShowForm(false); };
  const startEdit = (cat) => { setName(cat.name); setIcon(cat.icon); setColor(cat.color); setEditId(cat.id); setShowForm(true); };
  const save = () => {
    if (!name.trim()) return;
    if (editId) {
      setCategories(prev => prev.map(c => c.id===editId ? {...c, name, icon, color} : c));
    } else {
      // 新增時 order 設為目前最大 + 1
      const maxOrder = Math.max(0, ...categories.map(c => c.order ?? 0));
      setCategories(prev => [...prev, { id:`cat${Date.now()}`, name, icon, color, order: maxOrder + 1 }]);
    }
    reset();
  };
  const remove = (id) => {
    if (courses.some(c => c.category === id)) { alert("此分類下仍有課程，請先移動或刪除課程。"); return; }
    if (!confirm("確定要刪除此分類嗎？")) return;
    setCategories(prev => prev.filter(c => c.id !== id));
  };

  // 上下移動（交換 order）
  const move = (id, direction) => {
    setCategories(prev => {
      const sorted = [...prev].sort((a,b) => (a.order ?? 9999) - (b.order ?? 9999));
      const idx = sorted.findIndex(c => c.id === id);
      const target = direction === "up" ? idx - 1 : idx + 1;
      if (target < 0 || target >= sorted.length) return prev;
      // 交換兩個項目的 order
      const orderA = sorted[idx].order ?? idx + 1;
      const orderB = sorted[target].order ?? target + 1;
      return prev.map(c => {
        if (c.id === sorted[idx].id) return { ...c, order: orderB };
        if (c.id === sorted[target].id) return { ...c, order: orderA };
        return c;
      });
    });
  };

  // 直接編輯排序數字
  const updateOrder = (id, newOrder) => {
    setCategories(prev => prev.map(c => c.id === id ? { ...c, order: +newOrder || 0 } : c));
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>分類管理</h2>
        <div style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap" }}>
          {/* 檢視模式切換 */}
          <div style={{ display:"inline-flex", border:`1px solid ${C.border}`, borderRadius:7, overflow:"hidden", background:"#FFF" }}>
            <button onClick={() => setViewMode("card")} style={{ padding:"6px 12px", border:"none", background: viewMode==="card"?C.navy:"transparent", color: viewMode==="card"?"#FFF":C.textMid, fontSize:12, cursor:"pointer", fontWeight:500 }}>🔲 卡片</button>
            <button onClick={() => setViewMode("list")} style={{ padding:"6px 12px", border:"none", background: viewMode==="list"?C.navy:"transparent", color: viewMode==="list"?"#FFF":C.textMid, fontSize:12, cursor:"pointer", fontWeight:500 }}>📋 列表</button>
          </div>
          <Btn onClick={() => { reset(); setShowForm(true); }}>+ 新增分類</Btn>
        </div>
      </div>

      <div style={{ marginBottom:14, padding:"8px 12px", background:`${C.gold}10`, borderRadius:7, fontSize:11, color:C.navy, lineHeight:1.6 }}>
        💡 <strong>排序提示</strong>：分類顯示順序會依「排序」欄位由小到大排列，可以用上下箭頭調整，或直接修改排序數字。
      </div>

      {showForm && (
        <div style={{ background:"#FFF", borderRadius:9, padding:16, border:`1px solid ${C.border}`, marginBottom:14 }}>
          <h3 style={{ color:C.text, fontSize:14, fontWeight:600, margin:"0 0 10px" }}>{editId?"編輯分類":"新增分類"}</h3>
          <Field label="分類名稱"><input value={name} onChange={e => setName(e.target.value)} style={inp} placeholder="例：研發系統" /></Field>
          <Field label="圖示">
            <div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>
              {ICONS.map(ic => (
                <button key={ic} onClick={() => setIcon(ic)} style={{ width:30, height:30, borderRadius:5, border:`2px solid ${icon===ic?C.navy:C.border}`, background:icon===ic?`${C.navy}10`:"#FFF", fontSize:15, cursor:"pointer" }}>{ic}</button>
              ))}
            </div>
          </Field>
          <Field label="顏色">
            <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
              {COLORS.map(co => (
                <button key={co} onClick={() => setColor(co)} style={{ width:26, height:26, borderRadius:"50%", border:`3px solid ${color===co?"#333":"transparent"}`, background:co, cursor:"pointer" }} />
              ))}
            </div>
          </Field>
          <div style={{ display:"flex", gap:6, marginTop:10 }}>
            <Btn onClick={save} variant="gold">{editId?"儲存":"建立"}</Btn>
            <Btn onClick={reset} variant="outline">取消</Btn>
          </div>
        </div>
      )}

      {/* 卡片檢視 */}
      {viewMode === "card" && (
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(240px,1fr))", gap:10 }}>
          {categories.map((cat, idx) => {
            const count = courses.filter(c => c.category===cat.id).length;
            return (
              <div key={cat.id} style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}`, position:"relative" }}>
                <div style={{ position:"absolute", top:8, right:8, display:"flex", flexDirection:"column", gap:2 }}>
                  <button onClick={() => move(cat.id, "up")} disabled={idx===0} style={{ width:22, height:18, border:`1px solid ${C.border}`, background:"#FFF", borderRadius:4, cursor: idx===0?"default":"pointer", fontSize:10, opacity: idx===0?0.3:1 }}>▲</button>
                  <button onClick={() => move(cat.id, "down")} disabled={idx===categories.length-1} style={{ width:22, height:18, border:`1px solid ${C.border}`, background:"#FFF", borderRadius:4, cursor: idx===categories.length-1?"default":"pointer", fontSize:10, opacity: idx===categories.length-1?0.3:1 }}>▼</button>
                </div>
                <div style={{ display:"flex", alignItems:"center", gap:10, paddingRight:30 }}>
                  <div style={{ width:40, height:40, borderRadius:9, background:`${cat.color}15`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20, flexShrink:0 }}>{cat.icon}</div>
                  <div style={{ flex:1, minWidth:0 }}>
                    <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>{cat.name}</p>
                    <p style={{ margin:"2px 0 0", fontSize:11, color:C.textLight }}>排序 #{cat.order ?? "—"} · {count} 門課程</p>
                  </div>
                </div>
                <div style={{ display:"flex", gap:4, marginTop:10, justifyContent:"flex-end" }}>
                  <Btn onClick={() => startEdit(cat)} variant="outline" style={{ padding:"3px 8px", fontSize:10 }}>編輯</Btn>
                  <Btn onClick={() => remove(cat.id)} variant="danger" style={{ padding:"3px 8px", fontSize:10 }}>刪除</Btn>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* 列表檢視 */}
      {viewMode === "list" && (
        <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:560 }}>
            <thead>
              <tr style={{ borderBottom:`1px solid ${C.border}`, background:C.bg }}>
                <th style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500, width:80 }}>排序</th>
                <th style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500, width:50 }}>圖示</th>
                <th style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>名稱</th>
                <th style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>顏色</th>
                <th style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>課程數</th>
                <th style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>操作</th>
              </tr>
            </thead>
            <tbody>
              {categories.map((cat, idx) => {
                const count = courses.filter(c => c.category===cat.id).length;
                return (
                  <tr key={cat.id} style={{ borderBottom:`1px solid ${C.border}` }}>
                    <td style={{ padding:"8px 12px" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                        <input
                          type="number"
                          value={cat.order ?? ""}
                          onChange={e => updateOrder(cat.id, e.target.value)}
                          style={{ width:42, padding:"3px 6px", borderRadius:5, border:`1px solid ${C.border}`, fontSize:12, textAlign:"center", outline:"none" }}
                        />
                        <button onClick={() => move(cat.id, "up")} disabled={idx===0} style={{ padding:"2px 5px", border:`1px solid ${C.border}`, background:"#FFF", borderRadius:4, cursor: idx===0?"default":"pointer", fontSize:9, opacity: idx===0?0.3:1 }}>▲</button>
                        <button onClick={() => move(cat.id, "down")} disabled={idx===categories.length-1} style={{ padding:"2px 5px", border:`1px solid ${C.border}`, background:"#FFF", borderRadius:4, cursor: idx===categories.length-1?"default":"pointer", fontSize:9, opacity: idx===categories.length-1?0.3:1 }}>▼</button>
                      </div>
                    </td>
                    <td style={{ padding:"8px 12px", fontSize:18 }}>{cat.icon}</td>
                    <td style={{ padding:"8px 12px", fontSize:13, color:C.text, fontWeight:500 }}>{cat.name}</td>
                    <td style={{ padding:"8px 12px" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                        <div style={{ width:18, height:18, borderRadius:"50%", background:cat.color, border:`1px solid ${C.border}` }} />
                        <span style={{ fontSize:11, color:C.textLight, fontFamily:"monospace" }}>{cat.color}</span>
                      </div>
                    </td>
                    <td style={{ padding:"8px 12px", fontSize:12, color:C.textMid }}>{count}</td>
                    <td style={{ padding:"8px 12px" }}>
                      <div style={{ display:"flex", gap:4 }}>
                        <Btn onClick={() => startEdit(cat)} variant="outline" style={{ padding:"3px 8px", fontSize:10 }}>編輯</Btn>
                        <Btn onClick={() => remove(cat.id)} variant="danger" style={{ padding:"3px 8px", fontSize:10 }}>刪除</Btn>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function UserAdmin({ users, setUsers }) {
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState(null);
  const [empNo, setEmpNo] = useState("");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [department, setDepartment] = useState("");
  const [role, setRole] = useState("user");
  const [importResult, setImportResult] = useState(null);

  const reset = () => { setEmpNo(""); setName(""); setEmail(""); setPassword(""); setDepartment(""); setRole("user"); setEditId(null); setShowForm(false); };
  const startEdit = (u) => { setEmpNo(u.empNo||""); setName(u.name); setEmail(u.email); setPassword(u.password); setDepartment(u.department); setRole(u.role); setEditId(u.id); setShowForm(true); };
  const save = () => {
    if (!name.trim() || !email.trim() || !empNo.trim()) { alert("員工編號、姓名、Email 為必填"); return; }
    const finalPw = password.trim() || empNo;
    if (editId) setUsers(prev => prev.map(u => u.id===editId ? {...u, empNo, name, email, password:finalPw, department, role} : u));
    else setUsers(prev => [...prev, { id:`u${Date.now()}`, empNo, name, email, password:finalPw, department, role, mustChangePw:!password.trim() }]);
    reset();
  };

  // 下載 Excel 範本
  const downloadTemplate = () => {
    const data = [
      { "員工編號": "E00001", "姓名": "王小明", "Email": "wang@lkeng.com", "部門": "生產部", "角色": "user", "密碼（留空則預設員工編號）": "" },
      { "員工編號": "E00002", "姓名": "陳大華", "Email": "chen@lkeng.com", "部門": "品保部", "角色": "user", "密碼（留空則預設員工編號）": "" },
      { "員工編號": "E00003", "姓名": "李經理", "Email": "lee@lkeng.com", "部門": "管理部", "角色": "admin", "密碼（留空則預設員工編號）": "Lee@2026" },
    ];
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{wch:12},{wch:12},{wch:25},{wch:12},{wch:8},{wch:25}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "使用者匯入");
    XLSX.writeFile(wb, "亞翔學習平台_使用者批次匯入範本.xlsx");
  };

  // 處理 Excel 匯入
  const handleImport = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const result = { success:[], failed:[], duplicates:[] };

        const existingEmails = new Set(users.map(u => u.email.toLowerCase()));
        const existingEmpNos = new Set(users.map(u => u.empNo).filter(Boolean));
        const newUsers = [];

        rows.forEach((row, idx) => {
          const empNo = String(row["員工編號"] || "").trim();
          const name = String(row["姓名"] || "").trim();
          const email = String(row["Email"] || "").trim();
          const department = String(row["部門"] || "").trim();
          const role = ["admin","user"].includes(String(row["角色"]||"").trim()) ? String(row["角色"]).trim() : "user";
          const pwField = String(row["密碼（留空則預設員工編號）"] || row["密碼"] || "").trim();

          if (!empNo || !name || !email) {
            result.failed.push(`第 ${idx+2} 列：缺少必填欄位（員工編號/姓名/Email）`);
            return;
          }
          if (existingEmails.has(email.toLowerCase()) || existingEmpNos.has(empNo)) {
            result.duplicates.push(`第 ${idx+2} 列：${name}（${empNo}）已存在`);
            return;
          }
          existingEmails.add(email.toLowerCase());
          existingEmpNos.add(empNo);

          newUsers.push({
            id: `u${Date.now()}_${idx}`,
            empNo, name, email, department, role,
            password: pwField || empNo,
            mustChangePw: !pwField,
          });
          result.success.push(`${name}（${empNo}）`);
        });

        if (newUsers.length > 0) setUsers(prev => [...prev, ...newUsers]);
        setImportResult(result);
      } catch (err) {
        alert("匯入失敗：" + err.message);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>使用者管理</h2>
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          <Btn onClick={() => { reset(); setShowForm(true); }}>+ 新增使用者</Btn>
          <Btn onClick={downloadTemplate} variant="outline">📥 下載範本</Btn>
          <label style={{ cursor:"pointer", display:"inline-flex", alignItems:"center" }}>
            <span style={{ padding:"8px 16px", borderRadius:7, background:C.gold, color:"#FFF", fontSize:13, fontWeight:500, display:"inline-block" }}>📤 批次匯入</span>
            <input type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display:"none" }} />
          </label>
        </div>
      </div>

      {/* 匯入結果 */}
      {importResult && (
        <div style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}`, marginBottom:14 }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
            <h4 style={{ margin:0, fontSize:14, fontWeight:600 }}>📋 匯入結果</h4>
            <button onClick={() => setImportResult(null)} style={{ border:"none", background:"none", color:C.textLight, cursor:"pointer", fontSize:14 }}>✕</button>
          </div>
          <div style={{ display:"flex", gap:14, fontSize:12, flexWrap:"wrap", marginBottom:8 }}>
            <span style={{ color:C.success }}>✅ 成功：{importResult.success.length}</span>
            <span style={{ color:C.warning }}>⚠️ 重複跳過：{importResult.duplicates.length}</span>
            <span style={{ color:C.danger }}>❌ 失敗：{importResult.failed.length}</span>
          </div>
          {importResult.failed.length > 0 && <div style={{ fontSize:11, color:C.danger, marginTop:6 }}>{importResult.failed.map((m,i) => <div key={i}>• {m}</div>)}</div>}
          {importResult.duplicates.length > 0 && <div style={{ fontSize:11, color:C.warning, marginTop:6 }}>{importResult.duplicates.map((m,i) => <div key={i}>• {m}</div>)}</div>}
        </div>
      )}

      {showForm && (
        <div style={{ background:"#FFF", borderRadius:9, padding:16, border:`1px solid ${C.border}`, marginBottom:14 }}>
          <h3 style={{ color:C.text, fontSize:14, fontWeight:600, margin:"0 0 10px" }}>{editId?"編輯使用者":"新增使用者"}</h3>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(180px,1fr))", gap:10 }}>
            <Field label="員工編號 *"><input value={empNo} onChange={e => setEmpNo(e.target.value)} style={inp} placeholder="例：E00001" /></Field>
            <Field label="姓名 *"><input value={name} onChange={e => setName(e.target.value)} style={inp} /></Field>
            <Field label="電子信箱 *"><input value={email} onChange={e => setEmail(e.target.value)} style={inp} placeholder="user@lkeng.com" /></Field>
            <Field label="密碼（留空則預設員工編號）"><input value={password} onChange={e => setPassword(e.target.value)} style={inp} placeholder="留空 = 員工編號" /></Field>
            <Field label="部門"><input value={department} onChange={e => setDepartment(e.target.value)} style={inp} placeholder="例：生產部" /></Field>
            <Field label="角色">
              <select value={role} onChange={e => setRole(e.target.value)} style={inp}>
                <option value="user">一般使用者</option>
                <option value="admin">管理員</option>
              </select>
            </Field>
          </div>
          <div style={{ display:"flex", gap:6, marginTop:10 }}>
            <Btn onClick={save} variant="gold">{editId?"儲存":"建立"}</Btn>
            <Btn onClick={reset} variant="outline">取消</Btn>
          </div>
        </div>
      )}

      <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:650 }}>
          <thead><tr style={{ borderBottom:`1px solid ${C.border}` }}>
            {["員工編號","姓名","信箱","部門","角色","狀態","操作"].map(h => <th key={h} style={{ padding:"9px 10px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
          </tr></thead>
          <tbody>
            {users.map(u => (
              <tr key={u.id} style={{ borderBottom:`1px solid ${C.border}` }}>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.text, fontFamily:"monospace" }}>{u.empNo||"—"}</td>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.text }}>{u.name}</td>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>{u.email}</td>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>{u.department}</td>
                <td style={{ padding:"8px 10px" }}><span style={{ fontSize:10, padding:"3px 7px", borderRadius:7, background:u.role==="admin"?`${C.navy}12`:`${C.accent}12`, color:u.role==="admin"?C.navy:C.accent }}>{u.role==="admin"?"管理員":"使用者"}</span></td>
                <td style={{ padding:"8px 10px" }}>
                  {u.mustChangePw ? <span style={{ fontSize:10, padding:"3px 7px", borderRadius:7, background:`${C.warning}15`, color:C.warning }}>未改密碼</span> : <span style={{ fontSize:10, color:C.success }}>✓</span>}
                </td>
                <td style={{ padding:"8px 10px" }}>
                  <div style={{ display:"flex", gap:4 }}>
                    <Btn onClick={() => startEdit(u)} variant="outline" style={{ padding:"3px 7px", fontSize:10 }}>編輯</Btn>
                    <Btn onClick={() => { if(confirm(`確定刪除 ${u.name}？`)) setUsers(p => p.filter(uu => uu.id!==u.id)) }} variant="danger" style={{ padding:"3px 7px", fontSize:10 }}>刪除</Btn>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Analytics({ courses }) {
  const max = Math.max(...courses.map(c => c.views), 1);
  return (
    <div>
      <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:16 }}>學習分析</h2>
      <div style={{ background:"#FFF", borderRadius:9, padding:16, border:`1px solid ${C.border}` }}>
        <h3 style={{ color:C.text, fontSize:14, fontWeight:600, margin:"0 0 14px" }}>各課程瀏覽次數</h3>
        {[...courses].sort((a,b)=>b.views-a.views).map(c => (
          <div key={c.id} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:9 }}>
            <span style={{ width:130, fontSize:11, color:C.textMid, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", flexShrink:0 }}>{c.title}</span>
            <div style={{ flex:1, height:18, borderRadius:4, background:`${C.navy}08`, overflow:"hidden" }}>
              <div style={{ height:"100%", borderRadius:4, background:`linear-gradient(90deg, ${C.navy}, ${C.gold})`, width:`${c.views/max*100}%`, display:"flex", alignItems:"center", justifyContent:"flex-end", paddingRight:6 }}>
                {c.views>50 && <span style={{ fontSize:9, color:"#FFF", fontWeight:600 }}>{c.views}</span>}
              </div>
            </div>
            {c.views<=50 && <span style={{ fontSize:10, color:C.textLight, width:28, textAlign:"right" }}>{c.views}</span>}
          </div>
        ))}
      </div>
    </div>
  );
}

function QuizRecords({ quizResults, users, courses }) {
  const all = Object.entries(quizResults).map(([k,v]) => {
    const [uid, cid] = k.split("_");
    return { ...v, userName: users.find(u=>u.id===uid)?.name||"未知", courseName: courses.find(c=>c.id===cid)?.title||"未知", pct: Math.round(v.score/v.total*100) };
  });
  return (
    <div>
      <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:16 }}>測驗紀錄</h2>
      {all.length===0 ? <div style={{ textAlign:"center", padding:36, color:C.textLight }}><p style={{ fontSize:30 }}>📝</p><p>尚無測驗紀錄</p></div> : (
        <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:500 }}>
            <thead><tr style={{ borderBottom:`1px solid ${C.border}` }}>
              {["使用者","課程","分數","結果","日期"].map(h => <th key={h} style={{ padding:"9px 10px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
            </tr></thead>
            <tbody>{all.map((r,i) => (
              <tr key={i} style={{ borderBottom:`1px solid ${C.border}` }}>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.text }}>{r.userName}</td>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>{r.courseName}</td>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>{r.score}/{r.total}（{r.pct}%）</td>
                <td style={{ padding:"8px 10px" }}><span style={{ fontSize:10, padding:"3px 7px", borderRadius:7, background:r.pct>=60?`${C.success}12`:`${C.danger}12`, color:r.pct>=60?C.success:C.danger }}>{r.pct>=60?"通過":"未通過"}</span></td>
                <td style={{ padding:"8px 10px", fontSize:11, color:C.textLight }}>{new Date(r.date).toLocaleDateString("zh-TW")}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}
