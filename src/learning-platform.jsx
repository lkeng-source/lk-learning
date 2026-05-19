import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  loginWithEmail, logoutUser, watchAuthState, changePassword,
  getCurrentUserData, getAllUsers, watchAllUsers, createUserAccount, updateUserData, deleteUserData,
  watchCourses, addCourse, updateCourse, deleteCourse, incrementViews,
  watchCategories, addCategory, updateCategory, deleteCategory,
  watchUserHistory, recordWatchProgress, watchAllWatchHistory,
  watchAllQuizResults, saveQuizResult,
  initializeDefaultData
} from "./firebase-data";

/* ─── 配色 ─── */
const C = {
  navy: "#1B3A5C", navyDark: "#122840", navyLight: "#2A5080",
  gold: "#D4A528", goldLight: "#F0C850", goldPale: "#FFF8E7",
  accent: "#2980B9", bg: "#F5F6FA", card: "#FFFFFF",
  text: "#1E293B", textMid: "#475569", textLight: "#94A3B8",
  border: "#E2E8F0", success: "#16A34A", warning: "#D97706", danger: "#DC2626",
};

const ICONS = ["⚙️","✅","🛡️","🌿","💻","📊","🏭","🔬","📋","🎯","💡","🔧","📦","🧪","🏗️","📐"];
const COLORS = [
  // 藍色系
  "#0066CC", "#1E40AF", "#2563EB", "#3B82F6", "#0EA5E9", "#0891B2", "#06B6D4",
  // 綠色系
  "#00875A", "#059669", "#10B981", "#16A34A", "#22C55E", "#65A30D", "#2E7D32",
  // 紅/橘色系
  "#DC2626", "#EF4444", "#E65100", "#F97316", "#EA580C", "#D97706", "#F59E0B",
  // 紫色系
  "#5B21B6", "#6D28D9", "#7C3AED", "#8B5CF6", "#A855F7", "#C026D3", "#DB2777",
  // 棕/金色系
  "#B45309", "#92400E", "#78350F", "#A16207", "#854D0E",
  // 灰/中性色
  "#475569", "#64748B", "#6B7280", "#374151", "#1F2937",
];

const getYouTubeId = (url) => {
  if (!url) return null;
  const m = url.match(/(?:youtu\.be\/|youtube\.com\/(?:watch\?v=|embed\/|v\/|shorts\/))([^?&\s]+)/);
  return m ? m[1] : null;
};

/* ─── L&K Logo 元件 ─── */
function LKLogo({ size = 36, color = "#D4A528", background = "transparent" }) {
  return (
    <div style={{ width: size, height: size, background, borderRadius: size * 0.2, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
      <svg viewBox="0 0 100 50" width={size * 1.8} height={size * 0.9} xmlns="http://www.w3.org/2000/svg" style={{ display: "block" }}>
        <g fill={color} fontFamily="'Times New Roman', serif" fontWeight="bold">
          <text x="0" y="42" fontSize="48" letterSpacing="-2">L</text>
          <text x="34" y="42" fontSize="42" fontStyle="italic">&amp;</text>
          <text x="64" y="42" fontSize="48" letterSpacing="-1">K</text>
          <text x="95" y="14" fontSize="10">®</text>
        </g>
      </svg>
    </div>
  );
}

const inp = { width:"100%", padding:"9px 12px", borderRadius:7, border:`1px solid ${C.border}`, fontSize:13, outline:"none", boxSizing:"border-box", background:"#FFF", color:C.text };

function Btn({children, onClick, variant="primary", disabled, style, type}) {
  const v = {
    primary: { background:C.navy, color:"#FFF" },
    gold: { background:C.gold, color:"#FFF" },
    outline: { background:"transparent", border:`1px solid ${C.border}`, color:C.textMid },
    danger: { background:"transparent", border:`1px solid ${C.danger}40`, color:C.danger },
    ghost: { background:"transparent", color:C.navy },
  };
  return <button type={type||"button"} onClick={onClick} disabled={disabled} style={{ padding:"8px 16px", borderRadius:7, border:"none", fontSize:13, fontWeight:500, cursor:disabled?"default":"pointer", opacity:disabled?0.5:1, ...v[variant], ...style }}>{children}</button>;
}

function Field({ label, children }) {
  return <div style={{ marginBottom:10 }}><label style={{ display:"block", color:C.textMid, fontSize:12, marginBottom:4, fontWeight:500 }}>{label}</label>{children}</div>;
}

/* ══════════════════════════════════════
   主程式 - 用 Firebase 認證狀態決定畫面
   ══════════════════════════════════════ */
export default function App() {
  const [authUser, setAuthUser] = useState(null);  // Firebase Auth 物件
  const [currentUser, setCurrentUser] = useState(null);  // Firestore 中的使用者資料
  const [authLoading, setAuthLoading] = useState(true);
  const [view, setView] = useState("front");
  const [error, setError] = useState(null);

  // 監聽認證狀態
  useEffect(() => {
    const unsub = watchAuthState(async (user) => {
      setAuthUser(user);
      if (user) {
        try {
          const userData = await getCurrentUserData(user.uid);
          if (userData) {
            setCurrentUser(userData);
            setView(userData.role === "admin" ? "admin" : "front");
          } else {
            // Auth 有帳號但 Firestore 沒資料 - 異常狀態
            setError("找不到使用者資料，請聯絡管理員");
            await logoutUser();
          }
        } catch (e) {
          console.error(e);
          setError("載入使用者資料失敗：" + e.message);
        }
      } else {
        setCurrentUser(null);
      }
      setAuthLoading(false);
    });
    return unsub;
  }, []);

  // 嘗試初始化預設資料（會自動跳過已有資料）
  useEffect(() => {
    if (currentUser?.role === "admin") {
      initializeDefaultData().catch(err => console.error("Init failed:", err));
    }
  }, [currentUser]);

  const handleLogout = async () => {
    await logoutUser();
    setView("front");
  };

  if (authLoading) {
    return <LoadingScreen text="正在驗證身分..." />;
  }

  if (!authUser || !currentUser) {
    return <Login error={error} onError={setError} />;
  }

  if (view === "admin") {
    return <Admin currentUser={currentUser} onLogout={handleLogout} setView={setView} />;
  }
  return <Front currentUser={currentUser} onLogout={handleLogout} setView={setView} />;
}

/* ─── 載入畫面 ─── */
function LoadingScreen({ text }) {
  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:`linear-gradient(135deg, ${C.navyDark}, ${C.navy}, ${C.navyLight})`, flexDirection:"column", gap:16, color:"#FFF", fontFamily:"'Noto Sans TC',sans-serif" }}>
      <div style={{ width:48, height:48, border:"4px solid rgba(255,255,255,0.2)", borderTopColor:C.gold, borderRadius:"50%", animation:"spin 1s linear infinite" }} />
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      <div style={{ fontSize:15 }}>{text || "載入中..."}</div>
    </div>
  );
}

/* ══════════════════════════════════════
   登入頁
   ══════════════════════════════════════ */
function Login({ error, onError }) {
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [loading, setLoading] = useState(false);
  const [localErr, setLocalErr] = useState("");

  const submit = async () => {
    setLocalErr("");
    onError && onError(null);
    if (!email || !pw) { setLocalErr("請輸入帳號與密碼"); return; }
    setLoading(true);
    try {
      await loginWithEmail(email.trim(), pw);
      // 成功後 App 的 watchAuthState 會自動處理
    } catch (e) {
      const msg = e.code === "auth/invalid-credential" || e.code === "auth/wrong-password" || e.code === "auth/user-not-found"
        ? "帳號或密碼錯誤"
        : e.code === "auth/too-many-requests"
        ? "嘗試太多次，請稍後再試"
        : "登入失敗：" + e.message;
      setLocalErr(msg);
      setLoading(false);
    }
  };

  const displayErr = localErr || error;

  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:`linear-gradient(135deg, ${C.navyDark} 0%, ${C.navy} 50%, ${C.navyLight} 100%)`, fontFamily:"'Noto Sans TC',sans-serif", position:"relative", padding:20 }}>
      <div style={{ position:"absolute", top:-100, right:-100, width:350, height:350, borderRadius:"50%", background:`radial-gradient(circle, ${C.gold}18 0%, transparent 70%)` }} />
      <div style={{ position:"relative", width:"100%", maxWidth:400, padding:"36px 30px", borderRadius:18, background:"rgba(255,255,255,0.06)", backdropFilter:"blur(20px)", border:"1px solid rgba(255,255,255,0.1)", boxShadow:"0 24px 48px rgba(0,0,0,0.3)" }}>
        <div style={{ textAlign:"center", marginBottom:24 }}>
          <div style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", padding:"14px 22px", borderRadius:14, background:"#000", margin:"0 auto 12px", boxShadow:`0 4px 16px ${C.gold}40` }}>
            <LKLogo size={40} color={C.gold} />
          </div>
          <h1 style={{ color:"#FFF", fontSize:22, fontWeight:700, margin:0 }}>亞翔學習平台</h1>
          <p style={{ color:"rgba(255,255,255,0.5)", fontSize:12, marginTop:6 }}>L&K Engineering Learning</p>
        </div>

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
          {displayErr && <div style={{ color:"#FCA5A5", fontSize:12, textAlign:"center", padding:7, background:"rgba(220,38,38,0.12)", borderRadius:7 }}>{displayErr}</div>}
          <button onClick={submit} disabled={loading} style={{ width:"100%", padding:12, borderRadius:9, border:"none", background:`linear-gradient(135deg, ${C.gold}, ${C.goldLight})`, color:C.navyDark, fontSize:14, fontWeight:700, cursor:"pointer", marginTop:4, opacity:loading?0.7:1, boxShadow:`0 4px 12px ${C.gold}30` }}>
            {loading ? "登入中..." : "登入"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════
   修改密碼彈窗
   ══════════════════════════════════════ */
function ChangePasswordModal({ currentUser, onClose, force }) {
  const [oldPw, setOldPw] = useState("");
  const [newPw, setNewPw] = useState("");
  const [confirmPw, setConfirmPw] = useState("");
  const [err, setErr] = useState("");
  const [success, setSuccess] = useState(false);
  const [loading, setLoading] = useState(false);

  const submit = async () => {
    setErr("");
    if (newPw.length < 6) { setErr("新密碼至少 6 個字元"); return; }
    if (newPw !== confirmPw) { setErr("兩次輸入的密碼不一致"); return; }
    if (newPw === currentUser.empNo) { setErr("新密碼不可與員工編號相同"); return; }
    setLoading(true);
    try {
      await changePassword(oldPw, newPw);
      setSuccess(true);
      setTimeout(() => onClose(), 1500);
    } catch (e) {
      setErr(e.code === "auth/invalid-credential" || e.code === "auth/wrong-password" ? "舊密碼錯誤" : "修改失敗：" + e.message);
      setLoading(false);
    }
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
              <Btn onClick={submit} variant="gold" disabled={loading}>{loading?"處理中...":"確認修改"}</Btn>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ══════════════════════════════════════
   FRONT - 前台
   ══════════════════════════════════════ */
function Front({ currentUser, onLogout, setView }) {
  const [page, setPage] = useState("home");
  const [selectedCourse, setSelectedCourse] = useState(null);
  const [search, setSearch] = useState("");
  const [filterCat, setFilterCat] = useState("all");
  const [showPwModal, setShowPwModal] = useState(currentUser.mustChangePw || false);

  // 即時資料訂閱
  const [categories, setCategories] = useState([]);
  const [allCourses, setAllCourses] = useState([]);
  const [watchHistory, setWatchHistory] = useState({});
  const [quizResults, setQuizResults] = useState({});
  const [dataLoading, setDataLoading] = useState(true);

  useEffect(() => {
    const unsubs = [
      watchCategories(setCategories),
      watchCourses((cs) => { setAllCourses(cs); setDataLoading(false); }),
      watchUserHistory(currentUser.id, setWatchHistory),
      watchAllQuizResults(setQuizResults),
    ];
    return () => unsubs.forEach(u => u());
  }, [currentUser.id]);

  const sortedCategories = useMemo(() =>
    [...categories].sort((a,b) => (a.order ?? 9999) - (b.order ?? 9999)),
    [categories]
  );

  const courses = useMemo(() => allCourses.filter(c => c.status === "published"), [allCourses]);

  const userHistory = useMemo(() => Object.values(watchHistory), [watchHistory]);
  const userQuizzes = useMemo(() => 
    Object.values(quizResults).filter(q => q.userId === currentUser.id), 
    [quizResults, currentUser.id]
  );

  const filtered = useMemo(() => {
    let l = courses;
    if (filterCat !== "all") l = l.filter(c => c.category===filterCat);
    if (search) l = l.filter(c => c.title.includes(search) || c.description.includes(search) || c.instructor.includes(search));
    return l;
  }, [courses, filterCat, search]);
  const newest = useMemo(() => [...courses].sort((a,b) => (b.publishDate||"").localeCompare(a.publishDate||"")).slice(0,4), [courses]);
  const popular = useMemo(() => [...courses].sort((a,b) => (b.views||0)-(a.views||0)).slice(0,4), [courses]);
  const inProgress = useMemo(() => 
    userHistory.filter(h => h.progress < 100).map(h => ({...h, course: courses.find(c => c.id===h.courseId)})).filter(h => h.course), 
    [userHistory, courses]
  );

  const handleRecordWatch = async (courseId, chapterIndex, progress) => {
    try {
      await recordWatchProgress(currentUser.id, courseId, chapterIndex, progress);
      const course = allCourses.find(c => c.id === courseId);
      if (course) await incrementViews(courseId, course.views);
    } catch (e) {
      console.error("Record watch failed:", e);
    }
  };

  const handleSaveQuiz = async (courseId, score, total) => {
    try {
      await saveQuizResult(currentUser.id, courseId, score, total, currentUser.name);
    } catch (e) {
      console.error("Save quiz failed:", e);
    }
  };

  if (dataLoading) return <LoadingScreen text="載入課程資料..." />;

  const Card = ({ course }) => {
    const cat = sortedCategories.find(c => c.id===course.category);
    return (
      <div onClick={() => { setSelectedCourse(course); setPage("course"); }} style={{ background:"#FFF", borderRadius:10, overflow:"hidden", cursor:"pointer", border:`1px solid ${C.border}`, transition:"all 0.2s" }}
        onMouseOver={e => { e.currentTarget.style.transform="translateY(-2px)"; e.currentTarget.style.boxShadow="0 6px 16px rgba(0,0,0,0.08)"; }}
        onMouseOut={e => { e.currentTarget.style.transform="none"; e.currentTarget.style.boxShadow="none"; }}>
        <div style={{ height:90, background:`linear-gradient(135deg, ${cat?.color||C.navy}15, ${cat?.color||C.navy}30)`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:32 }}>{course.thumbnail}</div>
        <div style={{ padding:12 }}>
          <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:5 }}>
            <span style={{ fontSize:10, padding:"2px 6px", borderRadius:4, background:`${cat?.color||C.navy}12`, color:cat?.color||C.navy, fontWeight:500 }}>{cat?.name||"未分類"}</span>
            <span style={{ fontSize:10, color:C.textLight }}>👁 {course.views||0}</span>
          </div>
          <h3 style={{ margin:0, fontSize:13, fontWeight:600, color:C.text, lineHeight:1.4 }}>{course.title}</h3>
          <p style={{ margin:"4px 0 0", fontSize:11, color:C.textLight }}>{course.instructor} · {course.duration}分</p>
        </div>
      </div>
    );
  };

  return (
    <div style={{ minHeight:"100vh", background:C.bg, fontFamily:"'Noto Sans TC',sans-serif" }}>
      {showPwModal && <ChangePasswordModal currentUser={currentUser} onClose={() => setShowPwModal(false)} force={currentUser.mustChangePw} />}

      <div style={{ background:"#FFF", borderBottom:`2px solid ${C.gold}40`, padding:"0 20px", display:"flex", alignItems:"center", height:56, gap:12, position:"sticky", top:0, zIndex:100, boxShadow:"0 1px 4px rgba(0,0,0,0.04)", flexWrap:"wrap" }}>
        <div onClick={() => { setSelectedCourse(null); setPage("home"); }} style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer" }}>
          <div style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", padding:"4px 8px", borderRadius:6, background:"#000" }}>
            <LKLogo size={20} color={C.gold} />
          </div>
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
        <div style={{ width:28, height:28, borderRadius:"50%", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, display:"flex", alignItems:"center", justifyContent:"center", color:"#FFF", fontSize:12, fontWeight:600 }}>{currentUser.name?.[0]||"?"}</div>
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
            {sortedCategories.map(cat => (
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

          <div style={{ background:"#FFF", borderRadius:10, padding:16, border:`1px solid ${C.border}`, marginBottom:16, display:"flex", alignItems:"center", gap:14, flexWrap:"wrap" }}>
            <div style={{ width:54, height:54, borderRadius:"50%", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, display:"flex", alignItems:"center", justifyContent:"center", color:"#FFF", fontSize:22, fontWeight:600 }}>{currentUser.name?.[0]||"?"}</div>
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

      {page==="course" && selectedCourse && <CoursePage {...{categories:sortedCategories,course:selectedCourse,goBack:()=>{setSelectedCourse(null);setPage("home");},watchHistory,currentUser,recordWatch:handleRecordWatch,saveQuiz:handleSaveQuiz}} />}
    </div>
  );
}

/* ─── CoursePage ─── */
function CoursePage({ categories, course, goBack, watchHistory, currentUser, recordWatch, saveQuiz }) {
  const [activeCh, setActiveCh] = useState(0);
  const [showQuiz, setShowQuiz] = useState(false);
  const progress = watchHistory[`${currentUser.id}_${course.id}`]?.progress || 0;
  const cat = categories.find(c => c.id===course.category);
  const currentChapter = course.chapters?.[activeCh];
  const videoId = getYouTubeId(currentChapter?.youtubeUrl);

  useEffect(() => {
    if (videoId) {
      const op = Math.min(100, Math.round((activeCh + 1)/(course.chapters?.length||1)*100));
      recordWatch(course.id, activeCh, op);
    }
  }, [activeCh, videoId]);

  if (showQuiz) return <Quiz course={course} goBack={() => setShowQuiz(false)} saveQuiz={saveQuiz} />;

  return (
    <div style={{ padding:"24px 20px" }}>
      <button onClick={goBack} style={{ border:"none", background:"none", color:C.navy, fontSize:12, cursor:"pointer", padding:0, fontWeight:500, marginBottom:12 }}>← 返回</button>
      <div style={{ display:"flex", gap:18, flexWrap:"wrap" }}>
        <div style={{ flex:"1 1 380px", minWidth:0 }}>
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
          {/* YouTube 警告訊息已移除 */}
          <div style={{ marginTop:14 }}>
            <span style={{ fontSize:10, padding:"3px 7px", borderRadius:4, background:`${cat?.color||C.navy}12`, color:cat?.color||C.navy, fontWeight:500 }}>{cat?.name||"未分類"}</span>
            <span style={{ fontSize:11, color:C.textLight, marginLeft:8 }}>👁 {course.views||0}</span>
            <h1 style={{ fontSize:20, fontWeight:700, color:C.text, margin:"8px 0 4px" }}>{course.title}</h1>
            <p style={{ fontSize:13, color:C.navy, margin:"4px 0 0", fontWeight:500 }}>目前章節：{currentChapter?.title}</p>
            <p style={{ fontSize:12, color:C.textMid, margin:"4px 0 0" }}>講師：{course.instructor} · {course.duration} 分鐘</p>
            {course.publishDate && (
              <p style={{ fontSize:12, color:C.textLight, margin:"2px 0 0" }}>授課日期：{course.publishDate}</p>
            )}
            <p style={{ fontSize:13, color:C.text, marginTop:10, lineHeight:1.7 }}>{course.description}</p>

            {/* 授權聲明 */}
            <div style={{ marginTop:14, padding:"10px 12px", background:`${C.danger}08`, borderRadius:7, border:`1px solid ${C.danger}20`, display:"flex", alignItems:"flex-start", gap:8 }}>
              <span style={{ fontSize:16, color:C.danger, flexShrink:0 }}>🚫</span>
              <p style={{ margin:0, fontSize:11, color:C.textMid, lineHeight:1.6 }}>
                本課程由「<strong>{course.instructor || "課程作者"}</strong>」授權使用，您如需利用本作品，請另行向權利人取得授權。
              </p>
            </div>

            {/* 課程附件下載區 */}
            {course.files?.length > 0 && (
              <div style={{ marginTop:14, padding:14, background:C.goldPale, borderRadius:9, border:`1px solid ${C.gold}40` }}>
                <p style={{ fontSize:13, fontWeight:600, color:C.navy, margin:"0 0 8px" }}>📎 課程附件</p>
                {course.files.map((f, i) => (
                  <a
                    key={i}
                    href={f.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 10px", marginBottom:5, background:"#FFF", borderRadius:6, border:`1px solid ${C.border}`, textDecoration:"none", color:C.text, transition:"all 0.2s" }}
                    onMouseOver={e => { e.currentTarget.style.borderColor = C.gold; e.currentTarget.style.background = `${C.gold}08`; }}
                    onMouseOut={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.background = "#FFF"; }}
                  >
                    <span style={{ fontSize:18 }}>📄</span>
                    <span style={{ flex:1, fontSize:12, fontWeight:500 }}>{f.name}</span>
                    <span style={{ fontSize:11, color:C.accent }}>開啟 ↗</span>
                  </a>
                ))}
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
            {course.chapters?.map((ch,i) => {
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
  const [submitting, setSubmitting] = useState(false);
  const score = submitted ? course.quiz.reduce((s,q,i) => s + (answers[i]===q.answer?1:0), 0) : 0;
  const pct = Math.round(score/course.quiz.length*100);

  const submit = async () => {
    let s=0; course.quiz.forEach((q,i) => { if(answers[i]===q.answer) s++; });
    setSubmitting(true);
    try {
      await saveQuiz(course.id, s, course.quiz.length);
      setSubmitted(true);
    } catch (e) {
      alert("儲存測驗結果失敗：" + e.message);
    }
    setSubmitting(false);
  };

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
          <Btn onClick={submit} disabled={Object.keys(answers).length<course.quiz.length || submitting} style={{ alignSelf:"center", padding:"10px 28px" }}>{submitting?"送出中...":"提交答案"}</Btn>
        </div>
      )}
    </div>
  );
}

/* ══════════════════════════════════════
   ADMIN - 後台
   ══════════════════════════════════════ */
function Admin({ currentUser, onLogout, setView }) {
  const [tab, setTab] = useState("dashboard");
  const [showPwModal, setShowPwModal] = useState(false);

  const [categories, setCategories] = useState([]);
  const [courses, setCourses] = useState([]);
  const [users, setUsers] = useState([]);
  const [quizResults, setQuizResults] = useState({});
  const [allWatchHistory, setAllWatchHistory] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let loadCount = 0;
    const onLoad = () => { loadCount++; if (loadCount >= 2) setLoading(false); };
    const unsubs = [
      watchCategories((cs) => { setCategories(cs); onLoad(); }),
      watchCourses((cs) => { setCourses(cs); onLoad(); }),
      watchAllUsers(setUsers),
      watchAllQuizResults(setQuizResults),
      watchAllWatchHistory(setAllWatchHistory),
    ];
    return () => unsubs.forEach(u => u());
  }, []);

  const sortedCategories = useMemo(() =>
    [...categories].sort((a,b) => (a.order ?? 9999) - (b.order ?? 9999)),
    [categories]
  );

  const tabs = [
    { id:"dashboard", label:"總覽", icon:"📊" },
    { id:"courses", label:"課程管理", icon:"📚" },
    { id:"categories", label:"分類管理", icon:"🏷️" },
    { id:"users", label:"使用者管理", icon:"👥" },
    { id:"analytics", label:"學習分析", icon:"📈" },
    { id:"quizzes", label:"測驗紀錄", icon:"📝" },
  ];

  if (loading) return <LoadingScreen text="載入後台資料..." />;

  return (
    <div style={{ display:"flex", minHeight:"100vh", background:C.bg, fontFamily:"'Noto Sans TC',sans-serif" }}>
      {showPwModal && <ChangePasswordModal currentUser={currentUser} onClose={() => setShowPwModal(false)} />}
      <div style={{ width:200, background:C.navy, color:"#FFF", display:"flex", flexDirection:"column", flexShrink:0 }}>
        <div style={{ padding:"16px 14px", display:"flex", alignItems:"center", gap:10, borderBottom:"1px solid rgba(255,255,255,0.08)" }}>
          <div style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", padding:"5px 9px", borderRadius:6, background:"#000" }}>
            <LKLogo size={20} color={C.gold} />
          </div>
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
        {tab==="dashboard" && <Dashboard {...{courses,users,categories:sortedCategories}} />}
        {tab==="courses" && <CourseAdmin categories={sortedCategories} courses={courses} />}
        {tab==="categories" && <CategoryAdmin categories={sortedCategories} courses={courses} />}
        {tab==="users" && <UserAdmin users={users} />}
        {tab==="analytics" && <Analytics courses={courses} users={users} allWatchHistory={allWatchHistory} />}
        {tab==="quizzes" && <QuizRecords quizResults={quizResults} users={users} courses={courses} />}
      </div>
    </div>
  );
}

function Dashboard({ courses, users, categories }) {
  const totalViews = courses.reduce((s,c) => s+(c.views||0), 0);
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
          {[...courses].sort((a,b)=>(b.views||0)-(a.views||0)).slice(0,5).map((c,i) => (
            <div key={c.id} style={{ display:"flex", alignItems:"center", gap:7, padding:"5px 0", borderBottom:i<4?`1px solid ${C.border}`:"none" }}>
              <span style={{ width:18, height:18, borderRadius:"50%", background:i<3?C.navy:`${C.navy}20`, color:i<3?"#FFF":C.navy, display:"flex", alignItems:"center", justifyContent:"center", fontSize:9, fontWeight:600 }}>{i+1}</span>
              <span style={{ flex:1, fontSize:12, color:C.text }}>{c.title}</span>
              <span style={{ fontSize:11, color:C.textLight }}>{c.views||0}</span>
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

function CourseAdmin({ categories, courses }) {
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [title, setTitle] = useState("");
  const [category, setCategory] = useState(categories[0]?.id || "");
  const [instructor, setInstructor] = useState("");
  const [description, setDescription] = useState("");
  const [chapters, setChapters] = useState([{ title:"第一章", duration:15, youtubeUrl:"" }]);
  const [quiz, setQuiz] = useState([]);  // 測驗題目
  const [files, setFiles] = useState([]);  // 課程附件（連結方式）
  const [saving, setSaving] = useState(false);

  const reset = () => {
    setTitle(""); setCategory(categories[0]?.id || ""); setInstructor(""); setDescription("");
    setChapters([{ title:"第一章", duration:15, youtubeUrl:"" }]);
    setQuiz([]);
    setFiles([]);
    setEditing(null); setShowForm(false);
  };
  const startEdit = (c) => {
    setTitle(c.title); setCategory(c.category); setInstructor(c.instructor); setDescription(c.description);
    setChapters((c.chapters||[{title:"第一章",duration:15,youtubeUrl:""}]).map(ch => ({ ...ch, youtubeUrl: ch.youtubeUrl || "" })));
    setQuiz(c.quiz || []);
    setFiles(c.files || []);
    setEditing(c.id); setShowForm(true);
  };

  const save = async (publishNow = false) => {
    if (!title.trim()) { alert("請輸入課程名稱"); return; }
    // 驗證測驗題目
    for (let i = 0; i < quiz.length; i++) {
      const q = quiz[i];
      if (!q.q?.trim()) { alert(`第 ${i+1} 題題目不能空白`); return; }
      if (q.options.some(opt => !opt?.trim())) { alert(`第 ${i+1} 題的選項不能空白`); return; }
      if (q.answer === undefined || q.answer === null) { alert(`第 ${i+1} 題未選擇正確答案`); return; }
    }
    setSaving(true);
    const totalDuration = chapters.reduce((s,c) => s + (+c.duration||0), 0);
    try {
      if (editing) {
        await updateCourse(editing, { title, category, instructor, description, chapters, quiz, files, duration: totalDuration });
      } else {
        await addCourse({
          title, category, instructor, description,
          duration: totalDuration, thumbnail:"📘", views:0,
          publishDate: new Date().toISOString().split("T")[0],
          status: publishNow ? "published" : "draft",
          files, chapters, quiz
        });
      }
      reset();
    } catch (e) {
      alert("儲存失敗：" + e.message);
    }
    setSaving(false);
  };

  const togglePublish = async (c) => {
    try {
      await updateCourse(c.id, { status: c.status==="published" ? "draft" : "published" });
    } catch (e) { alert("更新失敗：" + e.message); }
  };

  const removeCourse = async (c) => {
    if (!confirm("確定刪除此課程？")) return;
    try { await deleteCourse(c.id); } catch (e) { alert("刪除失敗：" + e.message); }
  };

  const updateChapter = (idx, key, val) => setChapters(prev => prev.map((c,i) => i===idx ? { ...c, [key]: val } : c));
  const addChapter = () => setChapters(prev => [...prev, { title:`第${prev.length+1}章`, duration:15, youtubeUrl:"" }]);
  const removeChapter = (idx) => setChapters(prev => prev.length > 1 ? prev.filter((_,i) => i!==idx) : prev);

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

          <div style={{ marginTop:14, padding:12, background:C.bg, borderRadius:8 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
              <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>📋 章節設定（每章節對應一支 YouTube 影片）</p>
              <Btn onClick={addChapter} variant="outline" style={{ padding:"4px 10px", fontSize:11 }}>+ 新增章節</Btn>
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
                <input value={ch.youtubeUrl} onChange={e => updateChapter(idx, "youtubeUrl", e.target.value)} placeholder="🎬 YouTube 影片網址" style={inp} />
                {ch.youtubeUrl && (
                  getYouTubeId(ch.youtubeUrl)
                    ? <p style={{ fontSize:10, color:C.success, margin:"4px 0 0" }}>✅ 影片 ID 已偵測：{getYouTubeId(ch.youtubeUrl)}</p>
                    : <p style={{ fontSize:10, color:C.danger, margin:"4px 0 0" }}>⚠️ 網址格式無法識別</p>
                )}
              </div>
            ))}
          </div>

          {/* ══════ 測驗題目編輯區 ══════ */}
          <div style={{ marginTop:14, padding:12, background:C.bg, borderRadius:8 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
              <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>📝 課後測驗題目（共 {quiz.length} 題）</p>
              <Btn onClick={() => setQuiz(p => [...p, { q:"", options:["","","",""], answer:0 }])} variant="outline" style={{ padding:"4px 10px", fontSize:11 }}>+ 新增題目</Btn>
            </div>
            <div style={{ padding:"8px 10px", background:`${C.gold}10`, borderRadius:6, fontSize:11, color:C.navy, marginBottom:10, lineHeight:1.6 }}>
              💡 <strong>說明</strong>：每題有 4 個選項，請勾選正確答案。前台測驗需答對 60% 才算通過。沒有題目的課程，前台不會顯示測驗按鈕。
            </div>
            {quiz.length === 0 ? (
              <div style={{ padding:20, textAlign:"center", color:C.textLight, fontSize:12, background:"#FFF", borderRadius:7, border:`1px dashed ${C.border}` }}>
                目前尚無題目，點上方「+ 新增題目」開始建立
              </div>
            ) : quiz.map((q, qi) => (
              <div key={qi} style={{ background:"#FFF", borderRadius:7, padding:12, marginBottom:8, border:`1px solid ${C.border}` }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                  <span style={{ fontSize:12, fontWeight:600, color:C.navy }}>第 {qi+1} 題</span>
                  <button onClick={() => setQuiz(p => p.filter((_,i) => i!==qi))} style={{ border:"none", background:"none", color:C.danger, fontSize:11, cursor:"pointer" }}>移除</button>
                </div>
                <Field label="題目">
                  <input
                    value={q.q}
                    onChange={e => setQuiz(p => p.map((qq,i) => i===qi ? {...qq, q:e.target.value} : qq))}
                    placeholder="例：5S 中的「整理」指的是什麼？"
                    style={inp}
                  />
                </Field>
                <div style={{ marginTop:4 }}>
                  <p style={{ fontSize:11, color:C.textMid, marginBottom:6 }}>選項（請勾選正確答案）</p>
                  {q.options.map((opt, oi) => (
                    <div key={oi} style={{ display:"flex", alignItems:"center", gap:7, marginBottom:5 }}>
                      <input
                        type="radio"
                        name={`q${qi}-answer`}
                        checked={q.answer === oi}
                        onChange={() => setQuiz(p => p.map((qq,i) => i===qi ? {...qq, answer:oi} : qq))}
                        style={{ flexShrink:0, accentColor:C.navy, width:14, height:14 }}
                      />
                      <span style={{ fontSize:11, color:q.answer===oi?C.success:C.textLight, fontWeight:q.answer===oi?600:400, width:24 }}>{["A","B","C","D"][oi]}</span>
                      <input
                        value={opt}
                        onChange={e => setQuiz(p => p.map((qq,i) => i===qi ? {...qq, options: qq.options.map((o,j) => j===oi ? e.target.value : o)} : qq))}
                        placeholder={`選項 ${["A","B","C","D"][oi]}`}
                        style={{ ...inp, padding:"6px 10px", fontSize:12 }}
                      />
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>

          {/* ══════ 課程附件區（連結方式）══════ */}
          <div style={{ marginTop:14, padding:12, background:C.bg, borderRadius:8 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
              <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>📎 課程附件（共 {files.length} 個）</p>
              <Btn onClick={() => setFiles(p => [...p, { name:"", url:"" }])} variant="outline" style={{ padding:"4px 10px", fontSize:11 }}>+ 新增附件</Btn>
            </div>
            <div style={{ padding:"8px 10px", background:`${C.gold}10`, borderRadius:6, fontSize:11, color:C.navy, marginBottom:10, lineHeight:1.7 }}>
              💡 <strong>附件連結說明：</strong>請於下方欄位貼上檔案的網址。可使用任何**可開啟的網址**（公司內部檔案伺服器、雲端硬碟分享連結、ERP 系統連結等皆可）。<br />
              ⚠️ 請確認連結權限設定為「<strong>同仁可開啟</strong>」，否則使用者點擊會無法存取。
            </div>
            {files.length === 0 ? (
              <div style={{ padding:20, textAlign:"center", color:C.textLight, fontSize:12, background:"#FFF", borderRadius:7, border:`1px dashed ${C.border}` }}>
                目前沒有附件，點上方「+ 新增附件」加入連結
              </div>
            ) : files.map((f, fi) => (
              <div key={fi} style={{ background:"#FFF", borderRadius:7, padding:12, marginBottom:8, border:`1px solid ${C.border}` }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                  <span style={{ fontSize:12, fontWeight:600, color:C.navy }}>附件 {fi+1}</span>
                  <button onClick={() => setFiles(p => p.filter((_,i) => i!==fi))} style={{ border:"none", background:"none", color:C.danger, fontSize:11, cursor:"pointer" }}>移除</button>
                </div>
                <Field label="檔案名稱（顯示給使用者看的）">
                  <input
                    value={f.name}
                    onChange={e => setFiles(p => p.map((ff,i) => i===fi ? {...ff, name:e.target.value} : ff))}
                    placeholder="例：5S 管理實務 講義.pdf"
                    style={inp}
                  />
                </Field>
                <Field label="檔案連結網址">
                  <input
                    value={f.url}
                    onChange={e => setFiles(p => p.map((ff,i) => i===fi ? {...ff, url:e.target.value} : ff))}
                    placeholder="https://example.com/檔案網址"
                    style={inp}
                  />
                </Field>
                {f.url && !/^https?:\/\//i.test(f.url) && (
                  <p style={{ fontSize:10, color:C.danger, margin:"4px 0 0" }}>⚠️ 網址應以 http:// 或 https:// 開頭</p>
                )}
              </div>
            ))}
          </div>

          <div style={{ display:"flex", gap:6, marginTop:14, flexWrap:"wrap" }}>
            {editing ? (
              <Btn onClick={() => save(false)} variant="gold" disabled={saving}>{saving?"儲存中...":"儲存變更"}</Btn>
            ) : (
              <>
                <Btn onClick={() => save(true)} variant="gold" disabled={saving}>{saving?"建立中...":"✅ 建立並直接上架"}</Btn>
                <Btn onClick={() => save(false)} variant="outline" disabled={saving}>建立草稿</Btn>
              </>
            )}
            <Btn onClick={reset} variant="outline">取消</Btn>
          </div>
        </div>
      )}

      <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:700 }}>
          <thead><tr style={{ borderBottom:`1px solid ${C.border}` }}>
            {["課程名稱","分類","講師","章節/影片","題目","附件","瀏覽","狀態","操作"].map(h => <th key={h} style={{ padding:"9px 10px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
          </tr></thead>
          <tbody>
            {courses.map(c => {
              const cat = categories.find(cc => cc.id===c.category);
              const videoCount = (c.chapters||[]).filter(ch => getYouTubeId(ch.youtubeUrl)).length;
              const quizCount = (c.quiz||[]).length;
              const fileCount = (c.files||[]).length;
              return (
                <tr key={c.id} style={{ borderBottom:`1px solid ${C.border}` }}>
                  <td style={{ padding:"8px 10px", fontSize:12, color:C.text }}>{c.title}</td>
                  <td style={{ padding:"8px 10px" }}><span style={{ fontSize:10, padding:"2px 6px", borderRadius:4, background:`${cat?.color||C.navy}12`, color:cat?.color||C.navy }}>{cat?.name||"未分類"}</span></td>
                  <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>{c.instructor}</td>
                  <td style={{ padding:"8px 10px", fontSize:11, color:C.textMid }}>{videoCount}/{(c.chapters||[]).length} 🎬</td>
                  <td style={{ padding:"8px 10px", fontSize:11, color: quizCount > 0 ? C.success : C.textLight }}>{quizCount > 0 ? `${quizCount} 📝` : "—"}</td>
                  <td style={{ padding:"8px 10px", fontSize:11, color: fileCount > 0 ? C.success : C.textLight }}>{fileCount > 0 ? `${fileCount} 📎` : "—"}</td>
                  <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>{c.views||0}</td>
                  <td style={{ padding:"8px 10px" }}>
                    <span style={{ fontSize:10, padding:"3px 7px", borderRadius:7, background:c.status==="published"?`${C.success}12`:`${C.warning}12`, color:c.status==="published"?C.success:C.warning }}>{c.status==="published"?"已上架":"草稿"}</span>
                  </td>
                  <td style={{ padding:"8px 10px" }}>
                    <div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>
                      <Btn onClick={() => startEdit(c)} variant="outline" style={{ padding:"3px 7px", fontSize:10 }}>編輯</Btn>
                      <Btn onClick={() => togglePublish(c)} variant="outline" style={{ padding:"3px 7px", fontSize:10 }}>{c.status==="published"?"下架":"上架"}</Btn>
                      <Btn onClick={() => removeCourse(c)} variant="danger" style={{ padding:"3px 7px", fontSize:10 }}>刪除</Btn>
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

function CategoryAdmin({ categories, courses }) {
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState(null);
  const [name, setName] = useState("");
  const [icon, setIcon] = useState("⚙️");
  const [color, setColor] = useState("#0066CC");
  const [viewMode, setViewMode] = useState("card");
  const [saving, setSaving] = useState(false);

  const reset = () => { setName(""); setIcon("⚙️"); setColor("#0066CC"); setEditId(null); setShowForm(false); };
  const startEdit = (cat) => { setName(cat.name); setIcon(cat.icon); setColor(cat.color); setEditId(cat.id); setShowForm(true); };
  
  const save = async () => {
    if (!name.trim()) return;
    setSaving(true);
    try {
      if (editId) {
        await updateCategory(editId, { name, icon, color });
      } else {
        const maxOrder = Math.max(0, ...categories.map(c => c.order ?? 0));
        await addCategory({ name, icon, color, order: maxOrder + 1 });
      }
      reset();
    } catch (e) { alert("儲存失敗：" + e.message); }
    setSaving(false);
  };
  
  const remove = async (id) => {
    if (courses.some(c => c.category === id)) { alert("此分類下仍有課程，請先移動或刪除課程。"); return; }
    if (!confirm("確定要刪除此分類嗎？")) return;
    try { await deleteCategory(id); } catch (e) { alert("刪除失敗：" + e.message); }
  };

  const move = async (id, direction) => {
    const sorted = [...categories].sort((a,b) => (a.order ?? 9999) - (b.order ?? 9999));
    const idx = sorted.findIndex(c => c.id === id);
    const target = direction === "up" ? idx - 1 : idx + 1;
    if (target < 0 || target >= sorted.length) return;
    const orderA = sorted[idx].order ?? idx + 1;
    const orderB = sorted[target].order ?? target + 1;
    try {
      await updateCategory(sorted[idx].id, { order: orderB });
      await updateCategory(sorted[target].id, { order: orderA });
    } catch (e) { alert("排序失敗：" + e.message); }
  };

  const updateOrder = async (id, newOrder) => {
    try { await updateCategory(id, { order: +newOrder || 0 }); } catch (e) { alert("更新失敗：" + e.message); }
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>分類管理</h2>
        <div style={{ display:"flex", gap:6, alignItems:"center", flexWrap:"wrap" }}>
          <div style={{ display:"inline-flex", border:`1px solid ${C.border}`, borderRadius:7, overflow:"hidden", background:"#FFF" }}>
            <button onClick={() => setViewMode("card")} style={{ padding:"6px 12px", border:"none", background: viewMode==="card"?C.navy:"transparent", color: viewMode==="card"?"#FFF":C.textMid, fontSize:12, cursor:"pointer", fontWeight:500 }}>🔲 卡片</button>
            <button onClick={() => setViewMode("list")} style={{ padding:"6px 12px", border:"none", background: viewMode==="list"?C.navy:"transparent", color: viewMode==="list"?"#FFF":C.textMid, fontSize:12, cursor:"pointer", fontWeight:500 }}>📋 列表</button>
          </div>
          <Btn onClick={() => { reset(); setShowForm(true); }}>+ 新增分類</Btn>
        </div>
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
            <div style={{ display:"flex", gap:5, flexWrap:"wrap", marginBottom:8 }}>
              {COLORS.map(co => (
                <button key={co} onClick={() => setColor(co)} title={co} style={{ width:26, height:26, borderRadius:"50%", border:`3px solid ${color===co?"#333":"transparent"}`, background:co, cursor:"pointer", padding:0 }} />
              ))}
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:8, marginTop:8, padding:"8px 10px", background:C.bg, borderRadius:6 }}>
              <span style={{ fontSize:11, color:C.textMid }}>自訂顏色：</span>
              <input
                type="color"
                value={color}
                onChange={e => setColor(e.target.value)}
                style={{ width:32, height:28, border:"none", padding:0, borderRadius:4, cursor:"pointer", background:"transparent" }}
              />
              <input
                value={color}
                onChange={e => /^#[0-9A-Fa-f]{6}$/.test(e.target.value) && setColor(e.target.value)}
                placeholder="#000000"
                style={{ width:90, padding:"4px 8px", fontSize:11, fontFamily:"monospace", borderRadius:5, border:`1px solid ${C.border}`, outline:"none" }}
              />
              <span style={{ display:"inline-block", width:20, height:20, borderRadius:"50%", background:color, border:`1px solid ${C.border}` }} />
            </div>
          </Field>
          <div style={{ display:"flex", gap:6, marginTop:10 }}>
            <Btn onClick={save} variant="gold" disabled={saving}>{saving?"處理中...":(editId?"儲存":"建立")}</Btn>
            <Btn onClick={reset} variant="outline">取消</Btn>
          </div>
        </div>
      )}

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

function UserAdmin({ users }) {
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState(null);
  const [empNo, setEmpNo] = useState("");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [department, setDepartment] = useState("");
  const [role, setRole] = useState("user");
  const [saving, setSaving] = useState(false);
  const [showWarning, setShowWarning] = useState(true);

  const reset = () => { setEmpNo(""); setName(""); setEmail(""); setPassword(""); setDepartment(""); setRole("user"); setEditId(null); setShowForm(false); };
  const startEdit = (u) => { setEmpNo(u.empNo||""); setName(u.name); setEmail(u.email); setPassword(""); setDepartment(u.department); setRole(u.role); setEditId(u.id); setShowForm(true); };
  
  const save = async () => {
    if (!name.trim() || !email.trim() || !empNo.trim()) { alert("員工編號、姓名、Email 為必填"); return; }
    setSaving(true);
    try {
      if (editId) {
        // 編輯只能改 Firestore 資料（empNo, name, dept, role），密碼和 email 由 Firebase Auth 管
        await updateUserData(editId, { empNo, name, department, role });
        reset();
      } else {
        const finalPw = password.trim() || empNo;
        await createUserAccount(empNo, name, email, finalPw, role, department);
        alert("✅ 使用者建立成功！\n\n⚠️ 重要：建立新使用者會自動切換到該帳號登入，請重新登入管理員帳號繼續操作。");
        reset();
        // Auth 會自動切換到新使用者，這時候 App 的 watchAuthState 會處理
      }
    } catch (e) {
      alert("失敗：" + (e.code === "auth/email-already-in-use" ? "此 Email 已被使用" : e.message));
    }
    setSaving(false);
  };

  const remove = async (u) => {
    if (!confirm(`確定要刪除 ${u.name} 的資料嗎？\n\n⚠️ 注意：這只刪除 Firestore 資料，登入帳號需要管理員到 Firebase Console 手動刪除。`)) return;
    try { await deleteUserData(u.id); } catch (e) { alert("刪除失敗：" + e.message); }
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>使用者管理</h2>
        <Btn onClick={() => { reset(); setShowForm(true); }}>+ 新增使用者</Btn>
      </div>

      {showWarning && (
        <div style={{ padding:"10px 14px", background:`${C.warning}10`, borderRadius:8, border:`1px solid ${C.warning}30`, marginBottom:14, position:"relative" }}>
          <button onClick={() => setShowWarning(false)} style={{ position:"absolute", top:6, right:8, border:"none", background:"none", cursor:"pointer", color:C.textLight, fontSize:14 }}>✕</button>
          <p style={{ margin:0, fontSize:12, color:C.warning, lineHeight:1.6 }}>
            ⚠️ <strong>關於新增使用者</strong>：因為使用 Firebase 免費方案，新增使用者時系統會自動切換到該新帳號（這是 Firebase 的限制）。建立完一個使用者後，您需要<strong>重新登入管理員帳號</strong>才能繼續操作。
            <br />
            💡 <strong>建議</strong>：批次新增多位使用者請用「Firebase Console」操作，或請 MIS 協助。
          </p>
        </div>
      )}

      {showForm && (
        <div style={{ background:"#FFF", borderRadius:9, padding:16, border:`1px solid ${C.border}`, marginBottom:14 }}>
          <h3 style={{ color:C.text, fontSize:14, fontWeight:600, margin:"0 0 10px" }}>{editId?"編輯使用者":"新增使用者"}</h3>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(180px,1fr))", gap:10 }}>
            <Field label="員工編號 *"><input value={empNo} onChange={e => setEmpNo(e.target.value)} style={inp} placeholder="例：E00001" /></Field>
            <Field label="姓名 *"><input value={name} onChange={e => setName(e.target.value)} style={inp} /></Field>
            <Field label="電子信箱 *"><input value={email} onChange={e => setEmail(e.target.value)} style={inp} placeholder="user@lkeng.com" disabled={!!editId} /></Field>
            {!editId && <Field label="密碼（留空 = 員工編號）"><input value={password} onChange={e => setPassword(e.target.value)} style={inp} placeholder="密碼" /></Field>}
            <Field label="處別"><input value={department} onChange={e => setDepartment(e.target.value)} style={inp} placeholder="例：管理處" /></Field>
            <Field label="角色">
              <select value={role} onChange={e => setRole(e.target.value)} style={inp}>
                <option value="user">一般使用者</option>
                <option value="admin">管理員</option>
              </select>
            </Field>
          </div>
          <div style={{ display:"flex", gap:6, marginTop:10 }}>
            <Btn onClick={save} variant="gold" disabled={saving}>{saving?"處理中...":(editId?"儲存":"建立")}</Btn>
            <Btn onClick={reset} variant="outline">取消</Btn>
          </div>
        </div>
      )}

      <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:650 }}>
          <thead><tr style={{ borderBottom:`1px solid ${C.border}` }}>
            {["員工編號","姓名","信箱","處別","角色","狀態","操作"].map(h => <th key={h} style={{ padding:"9px 10px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
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
                    <Btn onClick={() => remove(u)} variant="danger" style={{ padding:"3px 7px", fontSize:10 }}>刪除</Btn>
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

function Analytics({ courses, users, allWatchHistory }) {
  const [selectedCourse, setSelectedCourse] = useState(null);
  const max = Math.max(...courses.map(c => c.views||0), 1);

  // 計算每個課程的觀看者明細
  const courseViewers = useMemo(() => {
    const map = {};
    allWatchHistory.forEach(record => {
      if (!map[record.courseId]) map[record.courseId] = [];
      map[record.courseId].push(record);
    });
    return map;
  }, [allWatchHistory]);

  // 點選課程後顯示明細
  if (selectedCourse) {
    const viewers = (courseViewers[selectedCourse.id] || []).map(r => {
      const user = users.find(u => u.id === r.userId);
      return {
        ...r,
        userName: user?.name || "未知使用者",
        empNo: user?.empNo || "—",
        department: user?.department || "—",
      };
    }).sort((a,b) => {
      const ta = a.lastWatched?.toMillis ? a.lastWatched.toMillis() : 0;
      const tb = b.lastWatched?.toMillis ? b.lastWatched.toMillis() : 0;
      return tb - ta;
    });

    return (
      <div>
        <button onClick={() => setSelectedCourse(null)} style={{ border:"none", background:"none", color:C.navy, fontSize:13, cursor:"pointer", padding:0, fontWeight:500, marginBottom:12 }}>← 返回學習分析</button>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:6 }}>{selectedCourse.title}</h2>
        <p style={{ fontSize:12, color:C.textLight, marginBottom:16 }}>共 {viewers.length} 筆觀看紀錄</p>

        {viewers.length === 0 ? (
          <div style={{ background:"#FFF", borderRadius:9, padding:36, border:`1px solid ${C.border}`, textAlign:"center", color:C.textLight }}>
            <p style={{ fontSize:30, margin:0 }}>👀</p>
            <p style={{ margin:"8px 0 0", fontSize:13 }}>還沒有人觀看這門課程</p>
          </div>
        ) : (
          <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:600 }}>
              <thead>
                <tr style={{ borderBottom:`1px solid ${C.border}`, background:C.bg }}>
                  {["員工編號","姓名","處別","進度","學習時數","最後觀看"].map(h => <th key={h} style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {viewers.map((v,i) => (
                  <tr key={i} style={{ borderBottom:`1px solid ${C.border}` }}>
                    <td style={{ padding:"8px 12px", fontSize:12, color:C.text, fontFamily:"monospace" }}>{v.empNo}</td>
                    <td style={{ padding:"8px 12px", fontSize:12, color:C.text, fontWeight:500 }}>{v.userName}</td>
                    <td style={{ padding:"8px 12px", fontSize:12, color:C.textMid }}>{v.department}</td>
                    <td style={{ padding:"8px 12px" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                        <div style={{ flex:1, maxWidth:100, height:6, borderRadius:3, background:`${C.navy}10` }}>
                          <div style={{ height:"100%", borderRadius:3, background:`linear-gradient(90deg, ${C.navy}, ${C.gold})`, width:`${v.progress||0}%` }} />
                        </div>
                        <span style={{ fontSize:11, color: (v.progress||0)>=100?C.success:C.textMid, fontWeight:500, width:32, textAlign:"right" }}>{v.progress||0}%</span>
                      </div>
                    </td>
                    <td style={{ padding:"8px 12px", fontSize:12, color:C.textMid }}>{v.totalTime || 0} 分</td>
                    <td style={{ padding:"8px 12px", fontSize:11, color:C.textLight }}>
                      {v.lastWatched?.toDate ? v.lastWatched.toDate().toLocaleString("zh-TW", { year:"numeric", month:"2-digit", day:"2-digit", hour:"2-digit", minute:"2-digit" }) : "—"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    );
  }

  return (
    <div>
      <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:16 }}>學習分析</h2>
      <div style={{ background:"#FFF", borderRadius:9, padding:16, border:`1px solid ${C.border}` }}>
        <h3 style={{ color:C.text, fontSize:14, fontWeight:600, margin:"0 0 14px" }}>各課程瀏覽次數</h3>
        <p style={{ fontSize:11, color:C.textLight, margin:"0 0 14px" }}>💡 點擊右側次數可查看明細</p>
        {[...courses].sort((a,b)=>(b.views||0)-(a.views||0)).map(c => (
          <div key={c.id} style={{ display:"flex", alignItems:"center", gap:12, marginBottom:10 }}>
            <span style={{ width:160, fontSize:13, color:C.text, fontWeight:500, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", flexShrink:0 }}>{c.title}</span>
            <div style={{ flex:1, height:24, borderRadius:6, background:`${C.navy}08`, overflow:"hidden", position:"relative" }}>
              <div style={{ height:"100%", borderRadius:6, background:`linear-gradient(90deg, ${C.navy}, ${C.gold})`, width:`${(c.views||0)/max*100}%`, transition:"width 0.5s" }} />
            </div>
            <button
              onClick={() => setSelectedCourse(c)}
              style={{ width:64, padding:"6px 10px", textAlign:"center", fontSize:18, fontWeight:700, color:C.navy, background:"#FFF", border:`1.5px solid ${C.gold}`, borderRadius:7, cursor:"pointer", transition:"all 0.2s" }}
              onMouseOver={e => { e.currentTarget.style.background = C.goldPale; }}
              onMouseOut={e => { e.currentTarget.style.background = "#FFF"; }}
              title="點擊查看明細"
            >
              {c.views || 0}
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

function QuizRecords({ quizResults, users, courses }) {
  const [filterUser, setFilterUser] = useState("");
  const [filterCourse, setFilterCourse] = useState("");
  const [filterResult, setFilterResult] = useState("all");  // all, pass, fail
  const [filterDateFrom, setFilterDateFrom] = useState("");
  const [filterDateTo, setFilterDateTo] = useState("");
  const [selected, setSelected] = useState({});  // {recordId: true}

  const all = useMemo(() => Object.values(quizResults).map(r => {
    const user = users.find(u=>u.id===r.userId);
    return {
      ...r,
      key: r.id || `${r.userId}_${r.courseId}`,
      userName: user?.name || r.userName || "未知",
      empNo: user?.empNo || "—",
      department: user?.department || "—",
      courseName: courses.find(c=>c.id===r.courseId)?.title || "未知",
      pct: Math.round((r.score||0)/(r.total||1)*100),
      dateObj: r.date?.toDate ? r.date.toDate() : null,
    };
  }).sort((a,b) => (b.dateObj?.getTime()||0) - (a.dateObj?.getTime()||0)), [quizResults, users, courses]);

  const filtered = useMemo(() => {
    return all.filter(r => {
      if (filterUser && !(r.userName.includes(filterUser) || r.empNo.includes(filterUser))) return false;
      if (filterCourse && r.courseId !== filterCourse) return false;
      if (filterResult === "pass" && r.pct < 60) return false;
      if (filterResult === "fail" && r.pct >= 60) return false;
      if (filterDateFrom && r.dateObj) {
        const from = new Date(filterDateFrom);
        if (r.dateObj < from) return false;
      }
      if (filterDateTo && r.dateObj) {
        const to = new Date(filterDateTo);
        to.setHours(23,59,59);
        if (r.dateObj > to) return false;
      }
      return true;
    });
  }, [all, filterUser, filterCourse, filterResult, filterDateFrom, filterDateTo]);

  const allChecked = filtered.length > 0 && filtered.every(r => selected[r.key]);
  const someChecked = filtered.some(r => selected[r.key]);
  const selectedCount = filtered.filter(r => selected[r.key]).length;

  const toggleAll = () => {
    if (allChecked) {
      const next = {...selected};
      filtered.forEach(r => delete next[r.key]);
      setSelected(next);
    } else {
      const next = {...selected};
      filtered.forEach(r => { next[r.key] = true; });
      setSelected(next);
    }
  };

  const toggleOne = (key) => {
    setSelected(p => ({...p, [key]: !p[key]}));
  };

  const resetFilter = () => {
    setFilterUser(""); setFilterCourse(""); setFilterResult("all");
    setFilterDateFrom(""); setFilterDateTo(""); setSelected({});
  };

  const exportExcel = () => {
    const rowsToExport = selectedCount > 0 ? filtered.filter(r => selected[r.key]) : filtered;
    if (rowsToExport.length === 0) { alert("沒有可匯出的資料"); return; }
    const data = rowsToExport.map(r => ({
      "員工編號": r.empNo,
      "姓名": r.userName,
      "處別": r.department,
      "課程名稱": r.courseName,
      "得分": r.score,
      "滿分": r.total,
      "百分比": `${r.pct}%`,
      "結果": r.pct >= 60 ? "通過" : "未通過",
      "測驗日期": r.dateObj ? r.dateObj.toLocaleString("zh-TW") : "",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{wch:10},{wch:10},{wch:12},{wch:25},{wch:8},{wch:8},{wch:8},{wch:8},{wch:20}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "測驗紀錄");
    const filename = `亞翔學習平台_測驗紀錄_${new Date().toISOString().split("T")[0]}.xlsx`;
    XLSX.writeFile(wb, filename);
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>測驗紀錄</h2>
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {(filterUser || filterCourse || filterResult !== "all" || filterDateFrom || filterDateTo) && (
            <Btn onClick={resetFilter} variant="outline" style={{ fontSize:11 }}>✕ 清除篩選</Btn>
          )}
          <Btn onClick={exportExcel} variant="gold" style={{ fontSize:12 }}>
            📥 匯出 Excel {selectedCount > 0 ? `（${selectedCount} 筆）` : `（全部 ${filtered.length} 筆）`}
          </Btn>
        </div>
      </div>

      {/* 篩選列 */}
      <div style={{ background:"#FFF", borderRadius:9, padding:12, border:`1px solid ${C.border}`, marginBottom:12 }}>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(160px,1fr))", gap:8 }}>
          <div>
            <label style={{ fontSize:10, color:C.textLight, display:"block", marginBottom:3 }}>🔍 員工 (姓名/編號)</label>
            <input value={filterUser} onChange={e => setFilterUser(e.target.value)} placeholder="輸入姓名或員工編號" style={{ ...inp, padding:"6px 10px", fontSize:12 }} />
          </div>
          <div>
            <label style={{ fontSize:10, color:C.textLight, display:"block", marginBottom:3 }}>📚 課程</label>
            <select value={filterCourse} onChange={e => setFilterCourse(e.target.value)} style={{ ...inp, padding:"6px 10px", fontSize:12 }}>
              <option value="">全部課程</option>
              {courses.map(c => <option key={c.id} value={c.id}>{c.title}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize:10, color:C.textLight, display:"block", marginBottom:3 }}>✅ 結果</label>
            <select value={filterResult} onChange={e => setFilterResult(e.target.value)} style={{ ...inp, padding:"6px 10px", fontSize:12 }}>
              <option value="all">全部</option>
              <option value="pass">通過</option>
              <option value="fail">未通過</option>
            </select>
          </div>
          <div>
            <label style={{ fontSize:10, color:C.textLight, display:"block", marginBottom:3 }}>📅 起始日期</label>
            <input type="date" value={filterDateFrom} onChange={e => setFilterDateFrom(e.target.value)} style={{ ...inp, padding:"6px 10px", fontSize:12 }} />
          </div>
          <div>
            <label style={{ fontSize:10, color:C.textLight, display:"block", marginBottom:3 }}>📅 結束日期</label>
            <input type="date" value={filterDateTo} onChange={e => setFilterDateTo(e.target.value)} style={{ ...inp, padding:"6px 10px", fontSize:12 }} />
          </div>
        </div>
      </div>

      <p style={{ fontSize:11, color:C.textLight, marginBottom:8 }}>共 {filtered.length} 筆紀錄（全部 {all.length} 筆）{selectedCount > 0 && <span style={{ color:C.navy, fontWeight:600 }}>· 已勾選 {selectedCount} 筆</span>}</p>

      {filtered.length === 0 ? (
        <div style={{ textAlign:"center", padding:36, color:C.textLight, background:"#FFF", borderRadius:9, border:`1px solid ${C.border}` }}>
          <p style={{ fontSize:30, margin:0 }}>📝</p>
          <p style={{ margin:"8px 0 0" }}>{all.length === 0 ? "尚無測驗紀錄" : "沒有符合條件的紀錄"}</p>
        </div>
      ) : (
        <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:700 }}>
            <thead>
              <tr style={{ borderBottom:`1px solid ${C.border}`, background:C.bg }}>
                <th style={{ padding:"10px 8px", textAlign:"center", width:40 }}>
                  <input
                    type="checkbox"
                    checked={allChecked}
                    ref={el => { if (el) el.indeterminate = !allChecked && someChecked; }}
                    onChange={toggleAll}
                    style={{ width:14, height:14, accentColor:C.navy, cursor:"pointer" }}
                  />
                </th>
                {["員工編號","姓名","處別","課程","分數","結果","日期"].map(h => <th key={h} style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {filtered.map((r,i) => (
                <tr key={r.key} style={{ borderBottom:`1px solid ${C.border}`, background: selected[r.key] ? `${C.gold}08` : "transparent" }}>
                  <td style={{ padding:"8px 8px", textAlign:"center" }}>
                    <input type="checkbox" checked={!!selected[r.key]} onChange={() => toggleOne(r.key)} style={{ width:14, height:14, accentColor:C.navy, cursor:"pointer" }} />
                  </td>
                  <td style={{ padding:"8px 12px", fontSize:12, color:C.text, fontFamily:"monospace" }}>{r.empNo}</td>
                  <td style={{ padding:"8px 12px", fontSize:12, color:C.text, fontWeight:500 }}>{r.userName}</td>
                  <td style={{ padding:"8px 12px", fontSize:12, color:C.textMid }}>{r.department}</td>
                  <td style={{ padding:"8px 12px", fontSize:12, color:C.textMid }}>{r.courseName}</td>
                  <td style={{ padding:"8px 12px", fontSize:12, color:C.textMid }}>{r.score}/{r.total}（{r.pct}%）</td>
                  <td style={{ padding:"8px 12px" }}>
                    <span style={{ fontSize:10, padding:"3px 7px", borderRadius:7, background: r.pct>=60?`${C.success}12`:`${C.danger}12`, color: r.pct>=60?C.success:C.danger, fontWeight:500 }}>{r.pct>=60?"通過":"未通過"}</span>
                  </td>
                  <td style={{ padding:"8px 12px", fontSize:11, color:C.textLight }}>{r.dateObj ? r.dateObj.toLocaleDateString("zh-TW") : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
