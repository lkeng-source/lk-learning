import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  loginWithEmail, logoutUser, watchAuthState, changePassword,
  getCurrentUserData, getAllUsers, watchAllUsers, createUserAccount, updateUserData, deleteUserData,
  watchUserData, toggleFavorite,
  scheduleUserChange, cancelUserChange, applyUserChange, reactivateUser, setUserStatus, checkAndApplyPendingChange,
  watchCourses, addCourse, updateCourse, deleteCourse, incrementViews,
  watchCategories, addCategory, updateCategory, deleteCategory,
  watchUserHistory, recordWatchProgress, watchAllWatchHistory,
  watchAllQuizResults, saveQuizResult, deleteQuizResult, deleteQuizResultsBatch,
  watchCourseReviews, watchAllReviews, saveReview, toggleReviewHelpful, deleteReview,
  sendResetPasswordEmail,
  watchCourseQuestions, watchAllQuestions, watchMyQuestions, addQuestion, answerQuestion, markQuestionRead, toggleQuestionShared, deleteQuestion,
  initializeDefaultData
} from "./firebase-data";

/* ─── 配色 ─── */
const C = {
  // 主色：較柔和的藍綠色，帶點青藍，更現代
  navy: "#2C5F7C",        // 主藍（從深藏青改為較淺的鋼藍）
  navyDark: "#1F4860",    // 漸層深色用
  navyLight: "#4A8AAC",   // 漸層淺色用
  // 金色保持（L&K 品牌色）
  gold: "#D4A528",
  goldLight: "#F0C850",
  goldPale: "#FFF8E7",
  // 點綴色 - 較活潑
  accent: "#4A90E2",      // 清亮藍（用於連結、強調）
  accentSoft: "#6FB8FF",  // 軟柔藍
  // 背景與卡片
  bg: "#F8F9FB",          // 偏暖的淺灰，比原本柔和
  bgSoft: "#EEF2F7",      // 區塊背景
  card: "#FFFFFF",
  // 文字
  text: "#2D3748",        // 主文字，比純黑柔和
  textMid: "#5A6878",     // 中等
  textLight: "#9BA8B8",   // 較淺
  border: "#E4E9F0",      // 邊框，比原本柔和
  // 狀態色
  success: "#22A06B",     // 較鮮明的綠
  warning: "#F0934A",     // 較柔和的橘
  danger: "#E25555",      // 較柔和的紅
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

// 影片連結解析（支援 YouTube、OneDrive、SharePoint、Vimeo 等）
// 回傳 { type, embedUrl, originalUrl } 或 null
const parseVideoUrl = (url) => {
  if (!url || typeof url !== "string") return null;
  const trimmed = url.trim();
  if (!trimmed) return null;

  // 如果是完整 iframe 嵌入碼，先抽出 src
  let workUrl = trimmed;
  if (workUrl.includes("<iframe")) {
    const srcMatch = workUrl.match(/src=["']([^"']+)["']/i);
    if (srcMatch) workUrl = srcMatch[1];
  }

  // 1. YouTube
  const ytMatch = workUrl.match(/(?:youtu\.be\/|youtube\.com\/(?:watch\?v=|embed\/|v\/|shorts\/))([^?&\s]+)/);
  if (ytMatch) {
    return {
      type: "youtube",
      videoId: ytMatch[1],
      embedUrl: `https://www.youtube.com/embed/${ytMatch[1]}?rel=0&modestbranding=1`,
      originalUrl: workUrl,
    };
  }

  // 2. OneDrive embed (onedrive.live.com/embed?...)
  if (/onedrive\.live\.com\/embed/i.test(workUrl)) {
    return {
      type: "onedrive",
      embedUrl: workUrl,
      originalUrl: workUrl,
    };
  }

  // 3. OneDrive share link (1drv.ms 或 onedrive.live.com 但沒有 embed)
  if (/1drv\.ms|onedrive\.live\.com/i.test(workUrl)) {
    // 1drv.ms 的分享連結加上 ?embed 也能嵌入（但不是所有都行）
    // 安全做法：以「外部連結」模式呈現
    return {
      type: "external",
      embedUrl: workUrl,
      originalUrl: workUrl,
    };
  }

  // 4. SharePoint (公司 SharePoint 影片)
  if (/sharepoint\.com.*\/embed/i.test(workUrl) || /sharepoint\.com.*\/video/i.test(workUrl)) {
    return {
      type: "sharepoint",
      embedUrl: workUrl,
      originalUrl: workUrl,
    };
  }

  // 5. Vimeo
  const vimMatch = workUrl.match(/vimeo\.com\/(?:video\/)?(\d+)/);
  if (vimMatch) {
    return {
      type: "vimeo",
      videoId: vimMatch[1],
      embedUrl: `https://player.vimeo.com/video/${vimMatch[1]}`,
      originalUrl: workUrl,
    };
  }

  // 6. 其他：當作外部連結
  if (/^https?:\/\//i.test(workUrl)) {
    return {
      type: "external",
      embedUrl: workUrl,
      originalUrl: workUrl,
    };
  }

  return null;
};

// 保留向後相容
const getYouTubeId = (url) => {
  const parsed = parseVideoUrl(url);
  return parsed?.type === "youtube" ? parsed.videoId : null;
};

// 姓氏馬賽克：保留姓氏，名字用 〇 取代
// 王小明 → 王〇〇 ；陳大 → 陳〇 ；歐陽小明 → 歐〇〇〇（保留首字）
// 英文名 John Smith → J〇〇〇
const maskName = (name) => {
  if (!name || typeof name !== "string") return "匿名";
  const trimmed = name.trim();
  if (trimmed.length <= 1) return trimmed + "〇";
  const first = trimmed[0];
  const rest = "〇".repeat(Math.max(1, trimmed.length - 1));
  return first + rest;
};

// 星星顯示元件（唯讀或可點選）
function StarRating({ value = 0, size = 18, onChange, readonly = true }) {
  const [hover, setHover] = useState(0);
  const display = hover || value;
  return (
    <span style={{ display: "inline-flex", gap: 2 }}>
      {[1, 2, 3, 4, 5].map((n) => (
        <span
          key={n}
          onClick={readonly ? undefined : () => onChange?.(n)}
          onMouseEnter={readonly ? undefined : () => setHover(n)}
          onMouseLeave={readonly ? undefined : () => setHover(0)}
          style={{
            fontSize: size,
            cursor: readonly ? "default" : "pointer",
            color: n <= display ? "#F5A623" : "#D8DEE6",
            lineHeight: 1,
            transition: "color 0.15s",
            userSelect: "none",
          }}
        >
          ★
        </span>
      ))}
    </span>
  );
}


/* ─── L&K Logo 元件（仿亞翔工程原版 logo）─── */
/* ─── 人像 icon（講師用，乾淨線條風格）─── */
function PersonIcon({ size = 16, color = "#5A6878", filled = false }) {
  if (filled) {
    return (
      <svg viewBox="0 0 24 24" width={size} height={size} style={{ display:"block", flexShrink:0 }}>
        <circle cx="12" cy="8" r="4" fill={color} />
        <path d="M4 20c0-4 3.6-7 8-7s8 3 8 7" fill={color} />
      </svg>
    );
  }
  return (
    <svg viewBox="0 0 24 24" width={size} height={size} style={{ display:"block", flexShrink:0 }}>
      <circle cx="12" cy="8" r="3.5" fill="none" stroke={color} strokeWidth="1.8" />
      <path d="M5 20c0-3.9 3.1-6.5 7-6.5s7 2.6 7 6.5" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" />
    </svg>
  );
}

function LKLogo({ size = 36, color = "#D4A528" }) {
  // viewBox 比例設計：L、(&)、K、® 四個元素
  // 整體比例約 寬:高 = 2.6:1
  const w = size * 2.6;
  const h = size;
  return (
    <svg
      viewBox="0 0 260 100"
      width={w}
      height={h}
      xmlns="http://www.w3.org/2000/svg"
      style={{ display: "block", flexShrink: 0 }}
    >
      <g fill={color}>
        {/* L 字母 */}
        <text
          x="5"
          y="78"
          fontSize="92"
          fontFamily="'Times New Roman', 'Noto Serif TC', serif"
          fontWeight="bold"
          letterSpacing="-2"
        >
          L
        </text>

        {/* & 符號外的圓圈 */}
        <circle
          cx="120"
          cy="55"
          r="33"
          fill="none"
          stroke={color}
          strokeWidth="3.5"
        />
        {/* & 符號 */}
        <text
          x="120"
          y="74"
          fontSize="52"
          fontFamily="'Times New Roman', 'Noto Serif TC', serif"
          fontWeight="bold"
          fontStyle="italic"
          textAnchor="middle"
        >
          &amp;
        </text>

        {/* K 字母 */}
        <text
          x="165"
          y="78"
          fontSize="92"
          fontFamily="'Times New Roman', 'Noto Serif TC', serif"
          fontWeight="bold"
          letterSpacing="-1"
        >
          K
        </text>

        {/* ® 註冊商標（右上角，圈起來） */}
        <circle
          cx="240"
          cy="22"
          r="9"
          fill="none"
          stroke={color}
          strokeWidth="1.5"
        />
        <text
          x="240"
          y="27"
          fontSize="12"
          fontFamily="'Times New Roman', serif"
          fontWeight="bold"
          textAnchor="middle"
        >
          R
        </text>
      </g>
    </svg>
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

/* ─── 可摺疊面板（瀏覽紀錄 / 測驗紀錄用） ─── */
function CollapsiblePanel({ title, icon, count, emptyText, children, defaultOpen = true }) {
  const [open, setOpen] = useState(defaultOpen);
  const hasContent = count > 0;
  return (
    <div style={{ background:C.bgSoft, borderRadius:9, border:`1px solid ${C.border}`, overflow:"hidden" }}>
      <div
        onClick={() => setOpen(o => !o)}
        style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"10px 14px", cursor:"pointer", background:"#FFF", borderBottom: open ? `1px solid ${C.border}` : "none", userSelect:"none" }}
      >
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          <span style={{ fontSize:16 }}>{icon}</span>
          <h3 style={{ fontSize:14, fontWeight:600, color:C.text, margin:0 }}>{title}</h3>
          <span style={{ fontSize:11, color:C.textLight, padding:"2px 8px", borderRadius:10, background:C.bg }}>{count} 筆</span>
        </div>
        <span style={{ fontSize:13, color:C.textLight, transition:"transform 0.2s", transform: open ? "rotate(180deg)" : "rotate(0deg)" }}>▼</span>
      </div>
      {open && (
        <div style={{ padding:"10px 12px", maxHeight:400, overflowY:"auto" }}>
          {hasContent ? children : (
            <p style={{ color:C.textLight, fontSize:12, textAlign:"center", padding:"20px 0", margin:0 }}>{emptyText}</p>
          )}
        </div>
      )}
    </div>
  );
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
            // 檢查是否有到期的異動 / 帳號是否已被鎖
            const check = await checkAndApplyPendingChange(userData);
            if (check.blocked) {
              setError(check.message);
              await logoutUser();
              setCurrentUser(null);
              setAuthLoading(false);
              return;
            }
            // 若剛套用了調部門異動，重新拿最新資料
            const freshData = check.applied ? await getCurrentUserData(user.uid) : userData;
            setCurrentUser(freshData);
            const isAdmin = freshData.role === "admin" || freshData.role === "superadmin";
            setView(isAdmin ? "admin" : "front");
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
    if (currentUser?.role === "admin" || currentUser?.role === "superadmin") {
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
  const [showForgot, setShowForgot] = useState(false);  // 忘記密碼視窗
  const [resetEmail, setResetEmail] = useState("");
  const [resetMsg, setResetMsg] = useState("");
  const [resetLoading, setResetLoading] = useState(false);

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

  const sendReset = async () => {
    setResetMsg("");
    if (!resetEmail.trim()) { setResetMsg("請輸入您的電子信箱"); return; }
    setResetLoading(true);
    try {
      await sendResetPasswordEmail(resetEmail.trim());
      setResetMsg("✅ 重設密碼信件已寄出！請至信箱收信（也請檢查垃圾郵件匣）。");
    } catch (e) {
      const msg = e.code === "auth/user-not-found"
        ? "查無此信箱，請確認輸入正確"
        : e.code === "auth/invalid-email"
        ? "信箱格式不正確"
        : "寄送失敗：" + e.message;
      setResetMsg("⚠️ " + msg);
    }
    setResetLoading(false);
  };

  const displayErr = localErr || error;

  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:`linear-gradient(135deg, ${C.navyDark} 0%, ${C.navy} 50%, ${C.navyLight} 100%)`, fontFamily:"'Noto Sans TC',sans-serif", position:"relative", padding:20 }}>
      <div style={{ position:"absolute", top:-100, right:-100, width:350, height:350, borderRadius:"50%", background:`radial-gradient(circle, ${C.gold}18 0%, transparent 70%)` }} />
      <div style={{ position:"relative", width:"100%", maxWidth:400, padding:"36px 30px", borderRadius:18, background:"rgba(255,255,255,0.06)", backdropFilter:"blur(20px)", border:"1px solid rgba(255,255,255,0.1)", boxShadow:"0 24px 48px rgba(0,0,0,0.3)" }}>
        <div style={{ textAlign:"center", marginBottom:24 }}>
          <div style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", margin:"0 auto 14px" }}>
            <LKLogo size={50} color={C.gold} />
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
          <button onClick={() => { setShowForgot(true); setResetEmail(email); setResetMsg(""); }} style={{ background:"none", border:"none", color:"rgba(255,255,255,0.6)", fontSize:12, cursor:"pointer", marginTop:2, textDecoration:"underline" }}>
            忘記密碼？
          </button>
        </div>
      </div>

      {/* 忘記密碼彈窗 */}
      {showForgot && (
        <div onClick={() => setShowForgot(false)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:20 }}>
          <div onClick={e => e.stopPropagation()} style={{ background:"#FFF", borderRadius:14, padding:"24px 22px", width:"100%", maxWidth:380, boxShadow:"0 24px 48px rgba(0,0,0,0.3)" }}>
            <h3 style={{ fontSize:17, fontWeight:700, color:C.text, margin:"0 0 6px" }}>🔑 重設密碼</h3>
            <p style={{ fontSize:12, color:C.textMid, margin:"0 0 16px", lineHeight:1.6 }}>
              輸入您的電子信箱，系統會寄送重設密碼的連結到您的信箱。
            </p>
            <label style={{ display:"block", color:C.textMid, fontSize:12, marginBottom:4 }}>電子信箱</label>
            <input value={resetEmail} onChange={e => setResetEmail(e.target.value)} onKeyDown={e => e.key==="Enter" && sendReset()} placeholder="your@lkeng.com" style={inp} />
            {resetMsg && (
              <div style={{ marginTop:10, padding:"8px 10px", borderRadius:7, fontSize:12, lineHeight:1.5, background: resetMsg.startsWith("✅") ? `${C.success}12` : `${C.danger}12`, color: resetMsg.startsWith("✅") ? C.success : C.danger }}>
                {resetMsg}
              </div>
            )}
            <div style={{ display:"flex", gap:8, marginTop:16, justifyContent:"flex-end" }}>
              <Btn onClick={() => setShowForgot(false)} variant="outline">關閉</Btn>
              <Btn onClick={sendReset} disabled={resetLoading} variant="gold">{resetLoading ? "寄送中..." : "寄送重設信"}</Btn>
            </div>
          </div>
        </div>
      )}
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
/* ─── 空狀態提示元件 ─── */
function EmptyState({ icon, title, hint, actionLabel, onAction }) {
  return (
    <div style={{ textAlign:"center", padding:"56px 20px", color:C.textLight, background:"#FFF", borderRadius:12, border:`1px solid ${C.border}` }}>
      <p style={{ fontSize:46, margin:0 }}>{icon}</p>
      <p style={{ fontSize:15, color:C.textMid, fontWeight:600, margin:"14px 0 4px" }}>{title}</p>
      {hint && <p style={{ fontSize:12, margin:0 }}>{hint}</p>}
      {actionLabel && onAction && (
        <button onClick={onAction} style={{ marginTop:16, padding:"8px 20px", borderRadius:8, border:"none", background:C.navy, color:"#FFF", fontSize:13, fontWeight:600, cursor:"pointer" }}>{actionLabel}</button>
      )}
    </div>
  );
}

function Front({ currentUser, onLogout, setView }) {
  const [page, setPage] = useState("courses");
  const [selectedCourse, setSelectedCourse] = useState(null);
  const [search, setSearch] = useState("");
  const [filterCat, setFilterCat] = useState("all");
  const [showPwModal, setShowPwModal] = useState(currentUser.mustChangePw || false);
  const [quizDetailRecord, setQuizDetailRecord] = useState(null);  // 查看自己過往的測驗詳解
  const [showUserMenu, setShowUserMenu] = useState(false);  // 頭像下拉選單
  const [learnTab, setLearnTab] = useState("progress");  // 我的學習分頁：progress/done/fav/quiz

  // 即時資料訂閱
  const [categories, setCategories] = useState([]);
  const [allCourses, setAllCourses] = useState([]);
  const [watchHistory, setWatchHistory] = useState({});
  const [quizResults, setQuizResults] = useState({});
  const [myQuestions, setMyQuestions] = useState([]);  // 我提出的問題（含講師回覆）
  const [userData, setUserData] = useState(currentUser);  // 即時使用者資料（含收藏）
  const [allUsers, setAllUsers] = useState([]);          // 全體使用者（公開儀表板用）
  const [allWatchHistory, setAllWatchHistory] = useState([]);  // 全體學習紀錄（公開儀表板用）
  const [dashMonth, setDashMonth] = useState("all");     // 儀表板月份篩選
  const [dataLoading, setDataLoading] = useState(true);

  useEffect(() => {
    const unsubs = [
      watchCategories(setCategories),
      watchCourses((cs) => { setAllCourses(cs); setDataLoading(false); }),
      watchUserHistory(currentUser.id, setWatchHistory),
      watchAllQuizResults(setQuizResults),
      watchMyQuestions(currentUser.id, setMyQuestions),
      watchUserData(currentUser.id, setUserData),
      watchAllUsers(setAllUsers),
      watchAllWatchHistory(setAllWatchHistory),
    ];
    return () => unsubs.forEach(u => u());
  }, [currentUser.id]);

  const favorites = userData?.favorites || [];
  const handleToggleFavorite = async (courseId) => {
    try { await toggleFavorite(currentUser.id, courseId, favorites); }
    catch (e) { console.error("toggle favorite failed:", e); }
  };

  // ─── 公開儀表板數據（依月份篩選）───
  // 產生今年 1 月到當月的選項
  const monthOptions = useMemo(() => {
    const opts = [{ value:"all", label:"全部期間" }];
    const now = new Date();
    const year = now.getFullYear();
    for (let m = now.getMonth(); m >= 0; m--) {
      const value = `${year}-${String(m+1).padStart(2,"0")}`;
      opts.push({ value, label: `${year} 年 ${m+1} 月` });
    }
    return opts;
  }, []);

  const dashStats = useMemo(() => {
    // 課程數、人數：顯示目前總數（若選特定月，課程數=該月(含)前已上架的課程）
    const publishedCourses = allCourses.filter(c => c.status === "published");
    let courseCount = publishedCourses.length;
    let totalMinutes = 0;

    if (dashMonth === "all") {
      totalMinutes = allWatchHistory.reduce((s, h) => s + (h.totalTime || 0), 0);
    } else {
      // 該月的學習時數：用 lastWatched 落在該月的紀錄加總
      allWatchHistory.forEach(h => {
        const t = h.lastWatched?.toDate ? h.lastWatched.toDate() : null;
        if (t) {
          const ym = `${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;
          if (ym === dashMonth) totalMinutes += (h.totalTime || 0);
        }
      });
      // 該月(含)之前已上架的課程數
      courseCount = publishedCourses.filter(c => {
        if (!c.publishDate) return true;
        const ym = c.publishDate.slice(0, 7);
        return ym <= dashMonth;
      }).length;
    }
    return {
      userCount: allUsers.length,
      courseCount,
      totalHours: Math.round(totalMinutes / 60 * 10) / 10,  // 換算小時，保留 1 位
      totalMinutes,
    };
  }, [allUsers, allCourses, allWatchHistory, dashMonth]);

  // 未讀回覆數（信箱紅點）：已回覆但同仁還沒讀
  const unreadCount = myQuestions.filter(q => q.status === "answered" && !q.readByUser).length;

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
    if (search) {
      const kw = search.trim().toLowerCase();
      l = l.filter(c =>
        (c.title || "").toLowerCase().includes(kw) ||
        (c.description || "").toLowerCase().includes(kw) ||
        (c.instructor || "").toLowerCase().includes(kw)
      );
    }
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

  const handleSaveQuiz = async (courseId, score, total, answers, questionsSnapshot) => {
    try {
      await saveQuizResult(currentUser.id, courseId, score, total, currentUser.name, answers, questionsSnapshot);
    } catch (e) {
      console.error("Save quiz failed:", e);
    }
  };

  if (dataLoading) return <LoadingScreen text="載入課程資料..." />;

  const Card = ({ course }) => {
    const cat = sortedCategories.find(c => c.id===course.category);
    const coverColor = course.coverColor || cat?.color || C.navy;
    const isFav = favorites.includes(course.id);
    const isArticle = course.contentType === "article";
    return (
      <div onClick={() => { setSelectedCourse(course); setPage("course"); }} style={{ background:"#FFF", borderRadius:12, overflow:"hidden", cursor:"pointer", border:`1px solid ${C.border}`, transition:"all 0.2s", display:"flex", flexDirection:"column" }}
        onMouseOver={e => { e.currentTarget.style.transform="translateY(-4px)"; e.currentTarget.style.boxShadow="0 12px 24px rgba(0,0,0,0.12)"; }}
        onMouseOut={e => { e.currentTarget.style.transform="none"; e.currentTarget.style.boxShadow="none"; }}>
        {/* 封面 16:9 */}
        <div style={{ aspectRatio:"16/9", background: course.coverUrl ? "#000" : `linear-gradient(135deg, ${coverColor}, ${coverColor}CC)`, display:"flex", alignItems:"center", justifyContent:"center", position:"relative", overflow:"hidden" }}>
          {course.coverUrl ? (
            <img src={course.coverUrl} alt={course.title} style={{ width:"100%", height:"100%", objectFit:"cover" }} onError={(e) => { e.target.style.display="none"; e.target.parentElement.style.background=`linear-gradient(135deg, ${coverColor}, ${coverColor}CC)`; }} />
          ) : (
            <div style={{ textAlign:"center", color:"#FFF", padding:14 }}>
              <div style={{ fontSize:34, marginBottom:6, opacity:0.9 }}>{isArticle ? "📄" : "📘"}</div>
              <div style={{ fontSize:14, fontWeight:700, lineHeight:1.3, textShadow:"0 1px 3px rgba(0,0,0,0.2)" }}>{course.title}</div>
            </div>
          )}
          {/* 分類標籤浮在封面左上 */}
          <span style={{ position:"absolute", top:8, left:8, fontSize:10, padding:"3px 8px", borderRadius:12, background:"rgba(255,255,255,0.92)", color:coverColor, fontWeight:600, backdropFilter:"blur(4px)" }}>{cat?.icon} {cat?.name||"未分類"}</span>
          {/* 收藏愛心（右上）*/}
          <button
            onClick={(e) => { e.stopPropagation(); handleToggleFavorite(course.id); }}
            title={isFav ? "取消收藏" : "加入收藏"}
            style={{ position:"absolute", top:6, right:6, width:32, height:32, borderRadius:"50%", border:"none", background:"rgba(255,255,255,0.92)", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", fontSize:16, backdropFilter:"blur(4px)", transition:"transform 0.15s" }}
            onMouseOver={e => e.currentTarget.style.transform="scale(1.15)"}
            onMouseOut={e => e.currentTarget.style.transform="scale(1)"}
          >
            {isFav ? "❤️" : "🤍"}
          </button>
        </div>
        {/* 內容 */}
        <div style={{ padding:16, flex:1, display:"flex", flexDirection:"column" }}>
          <h3 style={{ margin:0, fontSize:17, fontWeight:700, color:C.text, lineHeight:1.4, display:"-webkit-box", WebkitLineClamp:2, WebkitBoxOrient:"vertical", overflow:"hidden", minHeight:48 }}>{course.title}</h3>
          <div style={{ display:"flex", alignItems:"center", gap:6, marginTop:10 }}>
            <PersonIcon size={15} color={C.textMid} />
            <span style={{ fontSize:14, color:C.textMid, fontWeight:600 }}>{course.instructor}</span>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:14, marginTop:10, fontSize:13, color:C.textMid }}>
            <span>{isArticle ? "📄 圖文" : `🕐 ${course.duration} 分鐘`}</span>
            <span>👁 {course.views||0}</span>
            {course.quiz?.length > 0 && <span>📝 {course.quiz.length} 題</span>}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div style={{ minHeight:"100vh", background:C.bg, fontFamily:"'Noto Sans TC',sans-serif" }}>
      {showPwModal && <ChangePasswordModal currentUser={currentUser} onClose={() => setShowPwModal(false)} force={currentUser.mustChangePw} />}

      <div style={{ background:"#FFF", borderBottom:`2px solid ${C.gold}40`, padding:"0 20px", display:"flex", alignItems:"center", height:56, gap:12, position:"sticky", top:0, zIndex:100, boxShadow:"0 1px 4px rgba(0,0,0,0.04)", flexWrap:"wrap" }}>
        <div onClick={() => { setSelectedCourse(null); setPage("courses"); }} style={{ display:"flex", alignItems:"center", gap:10, cursor:"pointer", flexShrink:0 }}>
          <LKLogo size={22} color={C.gold} />
          <span style={{ fontSize:15, fontWeight:700, color:C.navy }}>亞翔學習</span>
        </div>
        <div style={{ flex:1, maxWidth:380, minWidth:120 }}>
          <input value={search} onChange={e => { setSearch(e.target.value); if (page !== "courses") setPage("courses"); }} placeholder="🔍 搜尋課程、講師..." style={{ width:"100%", padding:"8px 16px", borderRadius:20, border:`1px solid ${C.border}`, fontSize:13, outline:"none", boxSizing:"border-box", background:C.bg }} />
        </div>
        <div style={{ display:"flex", gap:2 }}>
          {[{l:"探索課程",p:"courses"},{l:"我的學習",p:"profile"}].map(n => (
            <button key={n.p} onClick={() => setPage(n.p)} style={{ padding:"6px 12px", borderRadius:8, border:"none", background:page===n.p?`${C.navy}10`:"transparent", color:page===n.p?C.navy:C.textLight, fontWeight:page===n.p?600:400, fontSize:13, cursor:"pointer" }}>{n.l}</button>
          ))}
        </div>
        {/* 右側：後台 + 信箱 + 頭像 + 登出，靠最右 */}
        <div style={{ display:"flex", alignItems:"center", gap:10, marginLeft:"auto", flexShrink:0 }}>
          {(currentUser.role==="admin" || currentUser.role==="superadmin") && <Btn onClick={() => setView("admin")} variant="outline" style={{ padding:"4px 10px", fontSize:11 }}>後台</Btn>}
          {/* 信箱 icon（站內信通知）*/}
          <button
            onClick={() => setPage("inbox")}
            title="我的信箱"
            style={{ position:"relative", width:36, height:36, borderRadius:"50%", border:"none", background: page==="inbox" ? `${C.navy}12` : "transparent", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", fontSize:18, transition:"all 0.2s" }}
            onMouseOver={e => { e.currentTarget.style.background = `${C.navy}10`; }}
            onMouseOut={e => { e.currentTarget.style.background = page==="inbox" ? `${C.navy}12` : "transparent"; }}
          >
            ✉️
            {/* 未讀紅點 */}
            {unreadCount > 0 && (
              <span style={{ position:"absolute", top:2, right:2, minWidth:16, height:16, padding:"0 4px", borderRadius:8, background:C.danger, color:"#FFF", fontSize:9, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", border:"1.5px solid #FFF" }}>{unreadCount}</span>
            )}
          </button>
          {/* 頭像 + 下拉 */}
          <div style={{ position:"relative" }}>
            <button
              onClick={() => setShowUserMenu(v => !v)}
              style={{ display:"flex", alignItems:"center", gap:6, border:"none", background:"transparent", cursor:"pointer", padding:"3px 6px", borderRadius:8 }}
              onMouseOver={e => { e.currentTarget.style.background = `${C.navy}08`; }}
              onMouseOut={e => { e.currentTarget.style.background = "transparent"; }}
            >
              <div style={{ width:30, height:30, borderRadius:"50%", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, display:"flex", alignItems:"center", justifyContent:"center", color:"#FFF", fontSize:13, fontWeight:600 }}>{currentUser.name?.[0]||"?"}</div>
              <span style={{ fontSize:12, color:C.textMid, fontWeight:500 }}>▾</span>
            </button>
            {showUserMenu && (
              <>
                <div onClick={() => setShowUserMenu(false)} style={{ position:"fixed", inset:0, zIndex:200 }} />
                <div style={{ position:"absolute", top:"calc(100% + 6px)", right:0, background:"#FFF", borderRadius:10, boxShadow:"0 8px 24px rgba(0,0,0,0.15)", border:`1px solid ${C.border}`, minWidth:180, zIndex:201, overflow:"hidden" }}>
                  <div style={{ padding:"12px 14px", borderBottom:`1px solid ${C.border}`, background:C.bgSoft }}>
                    <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>{currentUser.name}</p>
                    <p style={{ margin:"2px 0 0", fontSize:11, color:C.textLight }}>{currentUser.empNo} · {currentUser.department}</p>
                  </div>
                  <button onClick={() => { setPage("settings"); setShowUserMenu(false); }} style={{ width:"100%", padding:"10px 14px", border:"none", background:"transparent", textAlign:"left", fontSize:13, color:C.text, cursor:"pointer", display:"flex", alignItems:"center", gap:8 }} onMouseOver={e=>e.currentTarget.style.background=C.bgSoft} onMouseOut={e=>e.currentTarget.style.background="transparent"}>👤 個人檔案</button>
                  <button onClick={() => { setPage("profile"); setLearnTab("fav"); setShowUserMenu(false); }} style={{ width:"100%", padding:"10px 14px", border:"none", background:"transparent", textAlign:"left", fontSize:13, color:C.text, cursor:"pointer", display:"flex", alignItems:"center", gap:8 }} onMouseOver={e=>e.currentTarget.style.background=C.bgSoft} onMouseOut={e=>e.currentTarget.style.background="transparent"}>❤️ 我的收藏{favorites.length > 0 && <span style={{ marginLeft:"auto", fontSize:11, color:C.textLight }}>{favorites.length}</span>}</button>
                  <button onClick={() => { setPage("settings"); setShowUserMenu(false); }} style={{ width:"100%", padding:"10px 14px", border:"none", background:"transparent", textAlign:"left", fontSize:13, color:C.text, cursor:"pointer", display:"flex", alignItems:"center", gap:8 }} onMouseOver={e=>e.currentTarget.style.background=C.bgSoft} onMouseOut={e=>e.currentTarget.style.background="transparent"}>⚙️ 帳號設定</button>
                  <button onClick={() => { setPage("inbox"); setShowUserMenu(false); }} style={{ width:"100%", padding:"10px 14px", border:"none", background:"transparent", textAlign:"left", fontSize:13, color:C.text, cursor:"pointer", display:"flex", alignItems:"center", gap:8 }} onMouseOver={e=>e.currentTarget.style.background=C.bgSoft} onMouseOut={e=>e.currentTarget.style.background="transparent"}>✉️ 我的信箱{unreadCount > 0 && <span style={{ marginLeft:"auto", minWidth:18, height:18, padding:"0 5px", borderRadius:9, background:C.danger, color:"#FFF", fontSize:10, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center" }}>{unreadCount}</span>}</button>
                  <button onClick={() => { onLogout(); }} style={{ width:"100%", padding:"10px 14px", border:"none", borderTop:`1px solid ${C.border}`, background:"transparent", textAlign:"left", fontSize:13, color:C.danger, cursor:"pointer", display:"flex", alignItems:"center", gap:8 }} onMouseOver={e=>e.currentTarget.style.background=`${C.danger}08`} onMouseOut={e=>e.currentTarget.style.background="transparent"}>🚪 登出</button>
                </div>
              </>
            )}
          </div>
        </div>
      </div>

      {page==="courses" && (
        <div style={{ padding:"24px 20px", maxWidth:1200, margin:"0 auto" }}>
          {/* ══════ 公開儀表板 ══════ */}
          <div style={{ background:`linear-gradient(120deg, ${C.navy}, ${C.navyLight} 70%, ${C.gold} 150%)`, borderRadius:14, padding:"20px 24px", marginBottom:24, color:"#FFF", position:"relative", overflow:"hidden" }}>
            <div style={{ position:"absolute", right:-30, top:-40, width:160, height:160, borderRadius:"50%", background:"rgba(255,255,255,0.07)" }} />
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10, marginBottom:16, position:"relative" }}>
              <div>
                <h2 style={{ fontSize:17, fontWeight:700, margin:0 }}>📊 學習平台總覽</h2>
                <p style={{ fontSize:11, opacity:0.8, margin:"3px 0 0" }}>全公司學習數據一覽</p>
              </div>
              <select value={dashMonth} onChange={e => setDashMonth(e.target.value)} style={{ padding:"7px 12px", borderRadius:8, border:"none", background:"rgba(255,255,255,0.2)", color:"#FFF", fontSize:12, cursor:"pointer", outline:"none" }}>
                {monthOptions.map(o => <option key={o.value} value={o.value} style={{ color:"#333" }}>{o.label}</option>)}
              </select>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(140px,1fr))", gap:14, position:"relative" }}>
              {[
                { l:"系統總人數", v:dashStats.userCount, u:"人", i:"👥" },
                { l:dashMonth==="all"?"課程總數":"當期課程數", v:dashStats.courseCount, u:"門", i:"📚" },
                { l:dashMonth==="all"?"總學習時數":"當月學習時數", v:dashStats.totalHours, u:"小時", i:"⏱️" },
              ].map(s => (
                <div key={s.l} style={{ background:"rgba(255,255,255,0.13)", borderRadius:10, padding:"14px 16px", backdropFilter:"blur(4px)" }}>
                  <div style={{ fontSize:20, marginBottom:4 }}>{s.i}</div>
                  <div style={{ fontSize:11, opacity:0.85 }}>{s.l}</div>
                  <div style={{ fontSize:26, fontWeight:700, marginTop:2 }}>{s.v}<span style={{ fontSize:13, fontWeight:400, marginLeft:3, opacity:0.8 }}>{s.u}</span></div>
                </div>
              ))}
            </div>
          </div>

          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, marginBottom:14 }}>
            <h2 style={{ fontSize:20, fontWeight:700, color:C.text, margin:0 }}>探索課程</h2>
            {search && <span style={{ fontSize:12, color:C.textLight }}>搜尋「<strong style={{ color:C.navy }}>{search}</strong>」· {filtered.length} 筆結果</span>}
          </div>
          {/* 分類膠囊 */}
          <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:20 }}>
            <button onClick={() => setFilterCat("all")} style={{ padding:"6px 16px", borderRadius:18, border:`1.5px solid ${filterCat==="all"?C.navy:C.border}`, background:filterCat==="all"?C.navy:"#FFF", color:filterCat==="all"?"#FFF":C.textMid, fontSize:12, cursor:"pointer", fontWeight:600, transition:"all 0.2s" }}>全部課程</button>
            {sortedCategories.map(cat => (
              <button key={cat.id} onClick={() => setFilterCat(cat.id)} style={{ padding:"6px 16px", borderRadius:18, border:`1.5px solid ${filterCat===cat.id?cat.color:C.border}`, background:filterCat===cat.id?cat.color:"#FFF", color:filterCat===cat.id?"#FFF":C.textMid, fontSize:12, cursor:"pointer", fontWeight:600, transition:"all 0.2s" }}>{cat.icon} {cat.name}</button>
            ))}
          </div>
          {filtered.length===0 ? (
            <div style={{ textAlign:"center", padding:60, color:C.textLight }}>
              <p style={{ fontSize:42, margin:0 }}>🔍</p>
              <p style={{ fontSize:14, margin:"12px 0 0" }}>沒有找到符合條件的課程</p>
              {search && <p style={{ fontSize:12, margin:"4px 0 0" }}>試試其他關鍵字，或<button onClick={() => setSearch("")} style={{ border:"none", background:"none", color:C.navy, cursor:"pointer", fontSize:12, textDecoration:"underline" }}>清除搜尋</button></p>}
            </div>
          ) : (
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(240px,1fr))", gap:18 }}>{filtered.map(c => <Card key={c.id} course={c} />)}</div>
          )}
        </div>
      )}

      {page==="profile" && (
        <div>
          {/* 歡迎橫幅（金色系，呼應 LOGO）*/}
          <div style={{ background:`linear-gradient(135deg, ${C.navy} 0%, ${C.navyLight} 60%, ${C.gold} 160%)`, padding:"32px 24px", color:"#FFF", position:"relative", overflow:"hidden" }}>
            <div style={{ position:"absolute", right:-40, top:-40, width:200, height:200, borderRadius:"50%", background:"rgba(255,255,255,0.08)" }} />
            <div style={{ position:"absolute", right:60, bottom:-50, width:120, height:120, borderRadius:"50%", background:`${C.gold}25` }} />
            <div style={{ maxWidth:1100, margin:"0 auto", position:"relative" }}>
              <h1 style={{ fontSize:24, fontWeight:700, margin:0 }}>歡迎回來，{currentUser.name} 👋</h1>
              <p style={{ fontSize:13, opacity:0.85, marginTop:6 }}>持續學習，提升專業技能</p>
              <div style={{ display:"flex", gap:12, marginTop:20, flexWrap:"wrap" }}>
                {[
                  {l:"已完成",v:userHistory.filter(h=>h.progress>=100).length},
                  {l:"進行中",v:inProgress.length},
                  {l:"學習時數",v:`${userHistory.reduce((s,h)=>s+(h.totalTime||0),0)} 分`},
                  {l:"我的收藏",v:favorites.length},
                ].map(s => (
                  <div key={s.l} style={{ padding:"10px 18px", borderRadius:10, background:"rgba(255,255,255,0.15)", backdropFilter:"blur(4px)", minWidth:70 }}>
                    <span style={{ fontSize:11, opacity:0.8 }}>{s.l}</span>
                    <span style={{ display:"block", fontSize:22, fontWeight:700, marginTop:2 }}>{s.v}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          <div style={{ padding:"24px 20px", maxWidth:1100, margin:"0 auto" }}>
            {/* 分頁籤 */}
            <div style={{ display:"flex", gap:4, marginBottom:20, borderBottom:`1px solid ${C.border}`, flexWrap:"wrap" }}>
              {[
                { id:"progress", label:"進行中", count:inProgress.length },
                { id:"done", label:"已完成", count:userHistory.filter(h=>h.progress>=100).length },
                { id:"fav", label:"我的收藏", count:favorites.length },
                { id:"quiz", label:"測驗紀錄", count:userQuizzes.length },
              ].map(t => (
                <button key={t.id} onClick={() => setLearnTab(t.id)} style={{ padding:"10px 18px", border:"none", background:"transparent", borderBottom:`3px solid ${learnTab===t.id?C.gold:"transparent"}`, color:learnTab===t.id?C.navy:C.textLight, fontSize:14, fontWeight:learnTab===t.id?700:500, cursor:"pointer", marginBottom:-1, display:"flex", alignItems:"center", gap:6 }}>
                  {t.label}
                  {t.count > 0 && <span style={{ fontSize:11, minWidth:18, height:18, padding:"0 5px", borderRadius:9, background:learnTab===t.id?C.gold:C.bgSoft, color:learnTab===t.id?"#FFF":C.textLight, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:600 }}>{t.count}</span>}
                </button>
              ))}
            </div>

            {/* === 進行中 === */}
            {learnTab==="progress" && (
              inProgress.length === 0 ? (
                <EmptyState icon="📖" title="目前沒有進行中的課程" hint="去探索課程開始學習吧！" actionLabel="探索課程 →" onAction={() => setPage("courses")} />
              ) : (
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(280px,1fr))", gap:14 }}>
                  {inProgress.map(h => (
                    <div key={h.courseId} onClick={() => { setSelectedCourse(h.course); setPage("course"); }} style={{ background:"#FFF", borderRadius:10, padding:14, cursor:"pointer", border:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:12 }}
                      onMouseOver={e => e.currentTarget.style.boxShadow="0 4px 12px rgba(0,0,0,0.08)"}
                      onMouseOut={e => e.currentTarget.style.boxShadow="none"}>
                      <span style={{ fontSize:28 }}>{h.course.contentType==="article"?"📄":"📘"}</span>
                      <div style={{ flex:1, minWidth:0 }}>
                        <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{h.course.title}</p>
                        <div style={{ marginTop:6, height:5, borderRadius:3, background:C.border }}><div style={{ height:"100%", borderRadius:3, background:`linear-gradient(90deg, ${C.navy}, ${C.gold})`, width:`${h.progress}%` }} /></div>
                        <p style={{ margin:"4px 0 0", fontSize:11, color:C.textLight }}>{h.progress}% 完成</p>
                      </div>
                    </div>
                  ))}
                </div>
              )
            )}

            {/* === 已完成 === */}
            {learnTab==="done" && (() => {
              const doneList = userHistory.filter(h => h.progress >= 100);
              return doneList.length === 0 ? (
                <EmptyState icon="🎓" title="還沒有完成的課程" hint="完成課程後會出現在這裡" />
              ) : (
                <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                  {doneList.map(h => {
                    const c = courses.find(cc => cc.id===h.courseId);
                    if (!c) return null;
                    const dateStr = h.lastWatched?.toDate ? h.lastWatched.toDate().toLocaleDateString("zh-TW") : "";
                    return (
                      <div key={h.courseId} onClick={() => { setSelectedCourse(c); setPage("course"); }} style={{ background:"#FFF", borderRadius:9, padding:12, border:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:12, cursor:"pointer" }}>
                        <span style={{ fontSize:24 }}>{c.contentType==="article"?"📄":"📘"}</span>
                        <div style={{ flex:1, minWidth:0 }}>
                          <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>{c.title}</p>
                          <p style={{ margin:"2px 0 0", fontSize:11, color:C.textLight }}>完成於 {dateStr || "—"}</p>
                        </div>
                        <span style={{ fontSize:12, color:C.success, fontWeight:600, flexShrink:0 }}>✅ 已完成</span>
                      </div>
                    );
                  })}
                </div>
              );
            })()}

            {/* === 我的收藏 === */}
            {learnTab==="fav" && (
              favorites.length === 0 ? (
                <EmptyState icon="🤍" title="還沒有收藏任何課程" hint="在課程卡片右上角點愛心即可收藏" actionLabel="探索課程 →" onAction={() => setPage("courses")} />
              ) : (
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(240px,1fr))", gap:18 }}>
                  {courses.filter(c => favorites.includes(c.id)).map(c => <Card key={c.id} course={c} />)}
                </div>
              )
            )}

            {/* === 測驗紀錄 === */}
            {learnTab==="quiz" && (
              userQuizzes.length === 0 ? (
                <EmptyState icon="📝" title="還沒有測驗紀錄" hint="完成課程測驗後會出現在這裡" />
              ) : (
                <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                  {[...userQuizzes]
                    .sort((a,b) => {
                      const ta = a.date?.toMillis ? a.date.toMillis() : 0;
                      const tb = b.date?.toMillis ? b.date.toMillis() : 0;
                      return tb - ta;
                    })
                    .map((q, i) => {
                      const c = courses.find(cc => cc.id===q.courseId);
                      const pct = Math.round(q.score/q.total*100);
                      const dateStr = q.date?.toDate ? q.date.toDate().toLocaleDateString("zh-TW") : "";
                      const recordForModal = {
                        ...q,
                        userName: currentUser.name,
                        empNo: currentUser.empNo,
                        courseName: c?.title || "未知",
                        pct,
                        dateObj: q.date?.toDate ? q.date.toDate() : null,
                      };
                      return (
                        <div key={i} style={{ background:"#FFF", borderRadius:9, padding:12, border:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:12 }}>
                          <span style={{ fontSize:22 }}>{pct>=60?"🎉":"📖"}</span>
                          <div style={{ flex:1, minWidth:0 }}>
                            <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{c?.title||"未知"}</p>
                            <p style={{ margin:"2px 0 0", fontSize:11, color:C.textLight }}>{q.score}/{q.total}（{pct}%）{dateStr && ` · ${dateStr}`}</p>
                          </div>
                          <button onClick={() => setQuizDetailRecord(recordForModal)} style={{ padding:"4px 10px", fontSize:11, border:`1px solid ${C.border}`, background:"#FFF", borderRadius:6, cursor:"pointer", color:C.navy, flexShrink:0 }}>🔍 詳解</button>
                          <span style={{ fontSize:12, padding:"3px 9px", borderRadius:6, background:pct>=60?`${C.success}12`:`${C.danger}12`, color:pct>=60?C.success:C.danger, fontWeight:600, flexShrink:0 }}>{pct>=60?"通過":"未通過"}</span>
                        </div>
                      );
                    })}
                </div>
              )
            )}
          </div>
        </div>
      )}

      {/* 學員自己的測驗詳解彈窗 */}
      {quizDetailRecord && <QuizDetailModal record={quizDetailRecord} courses={courses} onClose={() => setQuizDetailRecord(null)} />}

      {page==="course" && selectedCourse && <CoursePage {...{categories:sortedCategories,course:selectedCourse,goBack:()=>{setSelectedCourse(null);setPage("courses");},watchHistory,currentUser,recordWatch:handleRecordWatch,saveQuiz:handleSaveQuiz}} />}

      {page==="inbox" && (
        <InboxPage myQuestions={myQuestions} courses={courses} currentUser={currentUser} onOpenCourse={(c) => { setSelectedCourse(c); setPage("course"); }} />
      )}

      {/* 帳號設定 */}
      {page==="settings" && (
        <div style={{ padding:"24px 20px", maxWidth:680, margin:"0 auto" }}>
          <h2 style={{ fontSize:20, fontWeight:700, color:C.text, marginBottom:18 }}>⚙️ 帳號設定</h2>

          {/* 個人資料卡 */}
          <div style={{ background:"#FFF", borderRadius:12, padding:20, border:`1px solid ${C.border}`, marginBottom:16 }}>
            <h3 style={{ fontSize:14, fontWeight:700, color:C.text, margin:"0 0 14px" }}>個人資料</h3>
            <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:16 }}>
              <div style={{ width:56, height:56, borderRadius:"50%", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, display:"flex", alignItems:"center", justifyContent:"center", color:"#FFF", fontSize:24, fontWeight:600 }}>{currentUser.name?.[0]||"?"}</div>
              <div>
                <p style={{ margin:0, fontSize:16, fontWeight:600, color:C.text }}>{currentUser.name}</p>
                <p style={{ margin:"3px 0 0", fontSize:12, color:C.textLight }}>{currentUser.role==="superadmin" ? "系統管理員" : currentUser.role==="admin" ? "管理員" : "一般使用者"}</p>
              </div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(200px,1fr))", gap:12 }}>
              {[
                {l:"員工編號", v:currentUser.empNo},
                {l:"處別", v:currentUser.department || "—"},
                {l:"電子信箱", v:currentUser.email || "—"},
              ].map(f => (
                <div key={f.l} style={{ padding:"10px 12px", background:C.bg, borderRadius:8 }}>
                  <p style={{ margin:0, fontSize:11, color:C.textLight }}>{f.l}</p>
                  <p style={{ margin:"3px 0 0", fontSize:13, color:C.text, fontWeight:500 }}>{f.v}</p>
                </div>
              ))}
            </div>
            <p style={{ fontSize:11, color:C.textLight, margin:"12px 0 0", lineHeight:1.6 }}>
              💡 個人資料由管理員統一維護，如需修改請聯絡管理處。
            </p>
          </div>

          {/* 安全設定卡 */}
          <div style={{ background:"#FFF", borderRadius:12, padding:20, border:`1px solid ${C.border}` }}>
            <h3 style={{ fontSize:14, fontWeight:700, color:C.text, margin:"0 0 6px" }}>安全設定</h3>
            <p style={{ fontSize:12, color:C.textLight, margin:"0 0 14px" }}>定期更換密碼能提升帳號安全</p>
            <Btn onClick={() => setShowPwModal(true)} variant="outline">🔒 修改密碼</Btn>
          </div>
        </div>
      )}
    </div>
  );
}

/* ─── 前台：我的信箱（站內信：提問與講師回覆）─── */
function InboxPage({ myQuestions, courses, currentUser, onOpenCourse }) {
  const [expanded, setExpanded] = useState(null);  // 展開的問題 id

  // 依時間排序（新到舊），未讀回覆優先
  const sorted = [...myQuestions].sort((a,b) => {
    const aUnread = a.status === "answered" && !a.readByUser;
    const bUnread = b.status === "answered" && !b.readByUser;
    if (aUnread !== bUnread) return aUnread ? -1 : 1;
    const ta = a.createdAt?.toMillis ? a.createdAt.toMillis() : 0;
    const tb = b.createdAt?.toMillis ? b.createdAt.toMillis() : 0;
    return tb - ta;
  });

  const openQuestion = async (q) => {
    const willExpand = expanded !== q.id;
    setExpanded(willExpand ? q.id : null);
    // 展開已回覆且未讀的 → 標記為已讀（清紅點）
    if (willExpand && q.status === "answered" && !q.readByUser) {
      try { await markQuestionRead(q.id); } catch (e) { console.error(e); }
    }
  };

  return (
    <div style={{ padding:"24px 20px", maxWidth:800, margin:"0 auto" }}>
      <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:4 }}>✉️ 我的信箱</h2>
      <p style={{ color:C.textLight, fontSize:12, marginBottom:18 }}>您向講師提出的問題與回覆都會出現在這裡</p>

      {sorted.length === 0 ? (
        <div style={{ background:"#FFF", borderRadius:12, padding:48, border:`1px solid ${C.border}`, textAlign:"center" }}>
          <div style={{ fontSize:48, marginBottom:12 }}>📭</div>
          <p style={{ fontSize:14, color:C.textMid, fontWeight:600, margin:0 }}>目前沒有訊息</p>
          <p style={{ fontSize:12, color:C.textLight, margin:"8px 0 0", lineHeight:1.6 }}>
            在課程頁點「🙋 我想問問題」向講師提問，<br />回覆後會出現在這裡。
          </p>
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          {sorted.map(q => {
            const course = courses.find(c => c.id === q.courseId);
            const isUnread = q.status === "answered" && !q.readByUser;
            const isOpen = expanded === q.id;
            const answered = q.status === "answered";
            return (
              <div key={q.id} style={{ background:"#FFF", borderRadius:12, border:`1px solid ${isUnread ? C.gold+"70" : C.border}`, overflow:"hidden" }}>
                {/* 標題列（可點開）*/}
                <div onClick={() => openQuestion(q)} style={{ padding:"14px 16px", cursor:"pointer", display:"flex", alignItems:"center", gap:10 }}>
                  <span style={{ fontSize:22, flexShrink:0 }}>{answered ? "📬" : "📨"}</span>
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                      <span style={{ fontSize:14, fontWeight:600, color:C.text }}>{q.subject}</span>
                      {isUnread && <span style={{ fontSize:10, padding:"1px 7px", borderRadius:8, background:C.danger, color:"#FFF", fontWeight:600 }}>新回覆</span>}
                    </div>
                    <p style={{ margin:"3px 0 0", fontSize:11, color:C.textLight }}>
                      {course?.title || q.courseName} · {answered ? "✓ 講師已回覆" : "⏳ 等待回覆中"}
                      {q.createdAt?.toDate && ` · ${q.createdAt.toDate().toLocaleDateString("zh-TW")}`}
                    </p>
                  </div>
                  <span style={{ fontSize:12, color:C.textLight, flexShrink:0 }}>{isOpen ? "▲" : "▼"}</span>
                </div>

                {/* 展開內容 */}
                {isOpen && (
                  <div style={{ padding:"0 16px 16px", borderTop:`1px solid ${C.border}` }}>
                    {/* 我的提問 */}
                    <div style={{ marginTop:14, marginBottom:12 }}>
                      <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:5 }}>
                        <span style={{ fontSize:13 }}>🙋</span>
                        <span style={{ fontSize:12, fontWeight:600, color:C.navy }}>我的提問</span>
                      </div>
                      <p style={{ margin:0, fontSize:13, color:C.textMid, lineHeight:1.6, whiteSpace:"pre-wrap" }}>{q.content}</p>
                    </div>
                    {/* 講師回覆 */}
                    {answered ? (
                      <div style={{ padding:"12px 14px", background:`${C.gold}08`, borderRadius:8, border:`1px solid ${C.gold}30` }}>
                        <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:5 }}>
                          <PersonIcon size={14} color={C.navy} filled />
                          <span style={{ fontSize:12, fontWeight:600, color:C.navy }}>講師回覆</span>
                          {q.answeredAt?.toDate && <span style={{ fontSize:10, color:C.textLight }}>· {q.answeredAt.toDate().toLocaleDateString("zh-TW")}</span>}
                        </div>
                        <p style={{ margin:0, fontSize:13, color:C.text, lineHeight:1.6, whiteSpace:"pre-wrap" }}>{q.answer}</p>
                      </div>
                    ) : (
                      <div style={{ padding:"10px 14px", background:C.bgSoft, borderRadius:8, fontSize:12, color:C.textLight, textAlign:"center" }}>
                        ⏳ 講師尚未回覆，請耐心等候
                      </div>
                    )}
                    {course && (
                      <button onClick={() => onOpenCourse(course)} style={{ marginTop:10, padding:"6px 14px", borderRadius:7, border:`1px solid ${C.border}`, background:"#FFF", color:C.navy, fontSize:12, fontWeight:500, cursor:"pointer" }}>
                        前往課程 →
                      </button>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ─── 文章內容渲染（解析 ## 標題、圖片網址、段落）─── */
function ArticleRenderer({ content }) {
  if (!content || !content.trim()) {
    return <p style={{ color:C.textLight, fontSize:13, textAlign:"center", padding:30 }}>（本課程尚無內容）</p>;
  }
  // 以空行分段
  const blocks = content.split(/\n\s*\n/).filter(b => b.trim());
  const isImageUrl = (s) => /^https?:\/\/\S+\.(jpe?g|png|gif|webp|svg)(\?\S*)?$/i.test(s.trim());

  return (
    <div style={{ fontSize:15, lineHeight:1.9, color:C.text }}>
      {blocks.map((block, i) => {
        const trimmed = block.trim();
        // 標題（## 開頭）
        if (trimmed.startsWith("## ")) {
          return <h2 key={i} style={{ fontSize:19, fontWeight:700, color:C.navy, margin:"24px 0 12px" }}>{trimmed.slice(3)}</h2>;
        }
        // 圖片網址（單獨一行）
        if (isImageUrl(trimmed)) {
          return <img key={i} src={trimmed} alt="" style={{ width:"100%", borderRadius:8, margin:"16px 0", display:"block" }} onError={(e)=>{e.target.style.display="none";}} />;
        }
        // 一般段落（段內換行轉成 <br>）
        const lines = trimmed.split("\n");
        return (
          <p key={i} style={{ margin:"0 0 16px" }}>
            {lines.map((line, j) => (
              <span key={j}>{line}{j < lines.length-1 && <br />}</span>
            ))}
          </p>
        );
      })}
    </div>
  );
}

/* ─── CoursePage ─── */
function CoursePage({ categories, course, goBack, watchHistory, currentUser, recordWatch, saveQuiz }) {
  const [activeCh, setActiveCh] = useState(0);
  const [showQuiz, setShowQuiz] = useState(false);
  const [isIdle, setIsIdle] = useState(false);  // 閒置狀態（自動暫停）
  const [reviews, setReviews] = useState([]);    // 本課程的所有評價
  const [myRating, setMyRating] = useState(0);    // 我要給的星數
  const [myReviewText, setMyReviewText] = useState("");  // 我的評價內容
  const [savingReview, setSavingReview] = useState(false);
  const [questions, setQuestions] = useState([]);   // 本課程所有提問
  const [showAskModal, setShowAskModal] = useState(false);  // 我想問問題視窗
  const [askSubject, setAskSubject] = useState("");
  const [askContent, setAskContent] = useState("");
  const [askingSubmit, setAskingSubmit] = useState(false);
  const historyRecord = watchHistory[`${currentUser.id}_${course.id}`];
  const progress = historyRecord?.progress || 0;
  const totalTime = historyRecord?.totalTime || 0;
  const cat = categories.find(c => c.id===course.category);
  const currentChapter = course.chapters?.[activeCh];
  const videoInfo = parseVideoUrl(currentChapter?.youtubeUrl);  // 欄位名仍叫 youtubeUrl 是為了向後相容
  const hasEmbed = videoInfo && (videoInfo.type === "youtube" || videoInfo.type === "onedrive" || videoInfo.type === "sharepoint" || videoInfo.type === "vimeo");
  const isArticle = course.contentType === "article";

  // ─── 訂閱本課程的評價 ───
  useEffect(() => {
    const unsub = watchCourseReviews(course.id, (list) => {
      setReviews(list);
      // 找出自己的評價，帶入編輯欄位
      const mine = list.find(r => r.userId === currentUser.id);
      if (mine) {
        setMyRating(mine.rating || 0);
        setMyReviewText(mine.content || "");
      }
    });
    return () => unsub && unsub();
  }, [course.id, currentUser.id]);

  // 評價統計
  const avgRating = reviews.length > 0 ? (reviews.reduce((s, r) => s + (r.rating || 0), 0) / reviews.length) : 0;
  const myReview = reviews.find(r => r.userId === currentUser.id);

  // 送出評價
  const submitReview = async () => {
    if (myRating < 1) { alert("請至少給 1 顆星"); return; }
    setSavingReview(true);
    try {
      await saveReview(currentUser.id, course.id, myRating, myReviewText.trim(), currentUser.name, myReview?.helpfulUserIds || []);
    } catch (e) {
      alert("送出評價失敗：" + e.message);
    }
    setSavingReview(false);
  };

  // 按「有幫助」
  const clickHelpful = async (review) => {
    try {
      await toggleReviewHelpful(review.id, currentUser.id, review.helpfulUserIds || []);
    } catch (e) {
      console.error("toggle helpful failed:", e);
    }
  };

  // ─── 訂閱本課程的提問 ───
  useEffect(() => {
    const unsub = watchCourseQuestions(course.id, setQuestions);
    return () => unsub && unsub();
  }, [course.id]);

  // 已分享到課後交流的問答（公開顯示，不透露提問者姓名）
  const sharedQuestions = questions
    .filter(q => q.shared && q.status === "answered")
    .sort((a,b) => {
      const ta = a.answeredAt?.toMillis ? a.answeredAt.toMillis() : 0;
      const tb = b.answeredAt?.toMillis ? b.answeredAt.toMillis() : 0;
      return tb - ta;
    });

  // 送出提問
  const submitQuestion = async () => {
    if (!askSubject.trim()) { alert("請輸入問題主旨"); return; }
    if (!askContent.trim()) { alert("請輸入問題內容"); return; }
    setAskingSubmit(true);
    try {
      await addQuestion({
        userId: currentUser.id,
        userName: currentUser.name,
        userEmail: currentUser.email || "",
        courseId: course.id,
        courseName: course.title,
        helperEmail: course.helperEmail || "",  // 課程小幫手信箱
        subject: askSubject.trim(),
        content: askContent.trim(),
      });
      setShowAskModal(false);
      setAskSubject("");
      setAskContent("");
      alert("✅ 您的問題已送出！講師回覆後會在「我的信箱」通知您。");
    } catch (e) {
      alert("送出失敗：" + e.message);
    }
    setAskingSubmit(false);
  };

  // ─── 閒置偵測：一段時間沒操作 → 自動暫停影片並跳提示 ───
  const IDLE_LIMIT_MS = 3 * 60 * 1000;  // 3 分鐘無操作即暫停（可調整）
  const idleTimerRef = useRef(null);

  useEffect(() => {
    // 只有在有影片、且正在觀看（非閒置、非測驗）時才啟動偵測
    if (!hasEmbed || showQuiz) return;

    const resetTimer = () => {
      if (idleTimerRef.current) clearTimeout(idleTimerRef.current);
      idleTimerRef.current = setTimeout(() => setIsIdle(true), IDLE_LIMIT_MS);
    };

    // 監聽各種使用者活動
    const events = ["mousemove", "mousedown", "keydown", "scroll", "touchstart", "click"];
    const onActivity = () => { if (!isIdle) resetTimer(); };
    events.forEach(ev => window.addEventListener(ev, onActivity));
    resetTimer();  // 啟動計時

    return () => {
      events.forEach(ev => window.removeEventListener(ev, onActivity));
      if (idleTimerRef.current) clearTimeout(idleTimerRef.current);
    };
  }, [hasEmbed, showQuiz, isIdle, activeCh]);

  // 使用者點「繼續觀看」→ 解除閒置
  const resumeWatching = () => setIsIdle(false);

  // 進度規則：依照「實際觀看時間 / 課程總時長」計算
  // 觀看時間達 80% 即視為完成（100%）
  const calculateProgress = (watchedMinutes, courseDuration) => {
    if (!courseDuration || courseDuration <= 0) return 0;
    const ratio = watchedMinutes / courseDuration;
    // 觀看時間達 80% 即視為 100%
    if (ratio >= 0.8) return 100;
    // 線性換算：0~80% 觀看時間 → 0~100% 進度
    return Math.min(100, Math.round((ratio / 0.8) * 100));
  };

  useEffect(() => {
    if (videoInfo) {
      // 每次切換章節時，記錄一次（totalTime 在後端會 +1）
      // 進度依據累計觀看時間動態計算
      const newProgress = calculateProgress(totalTime + 1, course.duration);
      recordWatch(course.id, activeCh, newProgress);
    }
  }, [activeCh, videoInfo?.embedUrl]);

  // 文章型課程：捲動到接近底部 → 視為閱讀完成（記 100%）
  useEffect(() => {
    if (!isArticle) return;
    if (progress >= 100) return;  // 已完成就不用再監聽
    const onScroll = () => {
      const scrollBottom = window.innerHeight + window.scrollY;
      const docHeight = document.documentElement.scrollHeight;
      // 捲到距離底部 120px 內，視為讀完
      if (scrollBottom >= docHeight - 120) {
        recordWatch(course.id, 0, 100);
      }
    };
    window.addEventListener("scroll", onScroll);
    // 內容很短不需捲動時，也檢查一次
    onScroll();
    return () => window.removeEventListener("scroll", onScroll);
  }, [isArticle, progress, course.id]);

  if (showQuiz) return <Quiz course={course} goBack={() => setShowQuiz(false)} saveQuiz={saveQuiz} />;

  // 影片平台標籤
  const platformLabels = {
    youtube: { name: "YouTube", color: "#FF0000", icon: "▶" },
    onedrive: { name: "OneDrive", color: "#0078D4", icon: "☁" },
    sharepoint: { name: "SharePoint", color: "#0078D4", icon: "☁" },
    vimeo: { name: "Vimeo", color: "#1AB7EA", icon: "▶" },
    external: { name: "外部連結", color: "#666666", icon: "🔗" },
  };
  const platform = videoInfo ? platformLabels[videoInfo.type] : null;

  return (
    <div style={{ padding:"24px 20px" }}>
      <button onClick={goBack} style={{ border:"none", background:"none", color:C.navy, fontSize:12, cursor:"pointer", padding:0, fontWeight:500, marginBottom:12 }}>← 返回</button>
      <div style={{ display:"flex", gap:18, flexWrap:"wrap" }}>
        <div style={{ flex:"1 1 380px", minWidth:0 }}>
          {!isArticle && (
          <div style={{ background:"#000", borderRadius:10, aspectRatio:"16/9", overflow:"hidden", position:"relative" }}>
            {hasEmbed && isIdle ? (
              /* 閒置暫停遮罩 */
              <div style={{ width:"100%", height:"100%", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", color:"#FFF", background:"linear-gradient(135deg, #1a1a1a, #2a2a2a)", padding:24, textAlign:"center" }}>
                <span style={{ fontSize:48, marginBottom:12 }}>⏸️</span>
                <p style={{ fontSize:16, fontWeight:600, margin:"0 0 6px" }}>影片已暫停</p>
                <p style={{ fontSize:12, opacity:0.6, margin:"0 0 20px", lineHeight:1.6 }}>
                  系統偵測到您已有一段時間沒有操作，<br />為確保學習效果，影片已自動暫停。
                </p>
                <button
                  onClick={resumeWatching}
                  style={{ padding:"12px 32px", background:C.gold, color:"#FFF", border:"none", borderRadius:10, fontSize:15, fontWeight:700, cursor:"pointer", boxShadow:`0 4px 16px ${C.gold}50` }}
                >
                  ▶ 繼續觀看
                </button>
              </div>
            ) : hasEmbed ? (
              <iframe
                key={videoInfo.embedUrl}
                src={videoInfo.embedUrl}
                title={currentChapter.title}
                style={{ width:"100%", height:"100%", border:"none" }}
                allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; fullscreen"
                allowFullScreen
              />
            ) : videoInfo?.type === "external" ? (
              <div style={{ width:"100%", height:"100%", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", color:"#FFF", background:"linear-gradient(135deg, #1a1a1a, #2a2a2a)", padding:24 }}>
                <span style={{ fontSize:52, marginBottom:14 }}>🎬</span>
                <p style={{ fontSize:14, margin:"0 0 4px", textAlign:"center" }}>{currentChapter.title}</p>
                <p style={{ fontSize:11, opacity:0.5, margin:"0 0 16px", textAlign:"center" }}>影片來源：{platform?.name || "外部連結"}</p>
                <a
                  href={videoInfo.embedUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{ padding:"10px 24px", background:C.gold, color:"#FFF", borderRadius:8, textDecoration:"none", fontSize:13, fontWeight:600, boxShadow:`0 4px 12px ${C.gold}40` }}
                >
                  ▶ 在新分頁播放影片
                </a>
              </div>
            ) : (
              <div style={{ width:"100%", height:"100%", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", color:"#FFF", background:"#1a1a1a" }}>
                <span style={{ fontSize:42 }}>📹</span>
                <p style={{ fontSize:13, opacity:0.7, margin:"8px 0 0" }}>本章節尚未提供影片</p>
                <p style={{ fontSize:11, opacity:0.4, margin:"4px 0 0" }}>請聯絡管理員上傳影片</p>
              </div>
            )}
          </div>
          )}

          {/* 文章型課程：顯示文章內容 */}
          {isArticle && (
            <div style={{ background:"#FFF", borderRadius:10, padding:"28px 30px", border:`1px solid ${C.border}` }}>
              {course.coverUrl && (
                <img src={course.coverUrl} alt={course.title} style={{ width:"100%", borderRadius:8, marginBottom:20, display:"block" }} onError={(e)=>{e.target.style.display="none";}} />
              )}
              <ArticleRenderer content={course.articleContent || ""} />
            </div>
          )}
          {/* YouTube 警告訊息已移除 */}
          <div style={{ marginTop:14 }}>
            <span style={{ fontSize:10, padding:"3px 7px", borderRadius:4, background:`${cat?.color||C.navy}12`, color:cat?.color||C.navy, fontWeight:500 }}>{cat?.name||"未分類"}</span>
            <span style={{ fontSize:11, color:C.textLight, marginLeft:8 }}>👁 {course.views||0}</span>
            <h1 style={{ fontSize:20, fontWeight:700, color:C.text, margin:"8px 0 4px" }}>{course.title}</h1>
            {!isArticle && <p style={{ fontSize:13, color:C.navy, margin:"4px 0 0", fontWeight:500 }}>目前章節：{currentChapter?.title}</p>}
            <p style={{ fontSize:12, color:C.textMid, margin:"4px 0 0" }}>講師：{course.instructor} · {isArticle ? `閱讀約 ${course.duration} 分鐘` : `${course.duration} 分鐘`}</p>
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
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginTop:5 }}>
              <span style={{ fontSize:10, color:C.textLight }}>{isArticle ? (progress>=100 ? "已閱讀完成" : "閱讀中") : `觀看 ${totalTime}/${course.duration||0} 分`}</span>
              <span style={{ fontSize:11, color: progress>=100?C.success:C.textMid, fontWeight:600 }}>{progress}%</span>
            </div>
          </div>
          {!isArticle && (
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
          )}
          {course.quiz?.length > 0 && (
            <button onClick={() => setShowQuiz(true)} style={{ width:"100%", padding:11, borderRadius:9, border:"none", background:`linear-gradient(135deg, ${C.gold}, ${C.goldLight})`, color:C.navyDark, fontSize:13, fontWeight:600, cursor:"pointer", boxShadow:`0 2px 8px ${C.gold}30`, marginBottom:10 }}>
              📝 開始課後測驗（{course.quiz.length} 題）
            </button>
          )}
          {/* 課程評分摘要卡 */}
          <div style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}` }}>
            <p style={{ margin:"0 0 8px", fontSize:12, fontWeight:600 }}>課程評分</p>
            {reviews.length > 0 ? (
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <span style={{ fontSize:26, fontWeight:700, color:C.text }}>{avgRating.toFixed(1)}</span>
                <div>
                  <StarRating value={Math.round(avgRating)} size={14} readonly />
                  <p style={{ margin:"2px 0 0", fontSize:10, color:C.textLight }}>{reviews.length} 則評價</p>
                </div>
              </div>
            ) : (
              <p style={{ margin:0, fontSize:11, color:C.textLight }}>尚無評價，成為第一個評價的人！</p>
            )}
            <button onClick={() => { const el = document.getElementById("review-section"); el?.scrollIntoView({ behavior:"smooth" }); }} style={{ width:"100%", marginTop:10, padding:"7px", borderRadius:7, border:`1px solid ${C.gold}`, background:"#FFF", color:C.navy, fontSize:12, fontWeight:600, cursor:"pointer" }}>
              ⭐ {myReview ? "查看 / 修改我的評價" : "我要評價"}
            </button>
          </div>
          {/* Q&A 提問 */}
          <button
            onClick={() => setShowAskModal(true)}
            style={{ width:"100%", marginTop:10, padding:"12px", borderRadius:9, border:"none", background:`linear-gradient(135deg, #FF8C42, #FFA62B)`, color:"#FFF", fontSize:14, fontWeight:700, cursor:"pointer", boxShadow:"0 2px 8px rgba(255,140,66,0.3)", display:"flex", alignItems:"center", justifyContent:"center", gap:6 }}
            onMouseOver={e => { e.currentTarget.style.transform="translateY(-1px)"; e.currentTarget.style.boxShadow="0 4px 12px rgba(255,140,66,0.4)"; }}
            onMouseOut={e => { e.currentTarget.style.transform="none"; e.currentTarget.style.boxShadow="0 2px 8px rgba(255,140,66,0.3)"; }}
          >
            ✋ Q&A 提問
          </button>
          {sharedQuestions.length > 0 && (
            <button onClick={() => { const el = document.getElementById("forum-section"); el?.scrollIntoView({ behavior:"smooth" }); }} style={{ width:"100%", marginTop:8, padding:"7px", borderRadius:7, border:`1px solid ${C.border}`, background:"#FFF", color:C.textMid, fontSize:12, fontWeight:600, cursor:"pointer" }}>
              💬 課後交流（{sharedQuestions.length}）
            </button>
          )}
        </div>
      </div>

      {/* ══════════ 課後交流（公開問答）══════════ */}
      {sharedQuestions.length > 0 && (
        <div id="forum-section" style={{ marginTop:28, maxWidth:780 }}>
          <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:4 }}>💬 課後交流</h2>
          <p style={{ fontSize:12, color:C.textLight, marginBottom:16 }}>講師精選的問答分享（為保護隱私，提問者統一以「學生」顯示）</p>
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            {sharedQuestions.map(q => (
              <div key={q.id} style={{ background:"#FFF", borderRadius:12, border:`1px solid ${C.border}`, overflow:"hidden" }}>
                {/* 學生提問 */}
                <div style={{ padding:"14px 16px", borderBottom:`1px solid ${C.border}` }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
                    <span style={{ width:28, height:28, borderRadius:"50%", background:`${C.navy}12`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:14, flexShrink:0 }}>🙋</span>
                    <span style={{ fontSize:13, fontWeight:600, color:C.navy }}>學生提問</span>
                  </div>
                  <p style={{ margin:"0 0 4px", fontSize:14, fontWeight:600, color:C.text }}>{q.subject}</p>
                  <p style={{ margin:0, fontSize:13, color:C.textMid, lineHeight:1.6 }}>{q.content}</p>
                </div>
                {/* 講師回覆 */}
                <div style={{ padding:"14px 16px", background:`${C.gold}08` }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
                    <span style={{ width:28, height:28, borderRadius:"50%", background:`${C.gold}25`, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}><PersonIcon size={16} color={C.navy} filled /></span>
                    <span style={{ fontSize:13, fontWeight:600, color:C.navy }}>講師回覆</span>
                  </div>
                  <p style={{ margin:0, fontSize:13, color:C.text, lineHeight:1.6 }}>{q.answer}</p>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ══════════ 課程評價區塊 ══════════ */}
      <div id="review-section" style={{ marginTop:28, maxWidth:780 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:4 }}>⭐ 課程評價</h2>
        <p style={{ fontSize:12, color:C.textLight, marginBottom:16 }}>分享你的學習心得，幫助其他同仁（顯示時僅保留姓氏，例如「王〇〇」）</p>

        {/* 評分總覽 */}
        <div style={{ display:"flex", gap:24, alignItems:"center", padding:"18px 20px", background:"#FFF", borderRadius:12, border:`1px solid ${C.border}`, marginBottom:18, flexWrap:"wrap" }}>
          <div style={{ textAlign:"center" }}>
            <div style={{ fontSize:40, fontWeight:700, color:C.text, lineHeight:1 }}>{reviews.length > 0 ? avgRating.toFixed(1) : "—"}</div>
            <div style={{ marginTop:6 }}><StarRating value={Math.round(avgRating)} size={16} readonly /></div>
            <p style={{ margin:"4px 0 0", fontSize:11, color:C.textLight }}>{reviews.length} 則評價</p>
          </div>
          {/* 星等分布 */}
          <div style={{ flex:1, minWidth:200 }}>
            {[5,4,3,2,1].map(star => {
              const count = reviews.filter(r => r.rating === star).length;
              const pct = reviews.length > 0 ? (count / reviews.length * 100) : 0;
              return (
                <div key={star} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:3 }}>
                  <span style={{ fontSize:11, color:C.textMid, width:30 }}>{star} ★</span>
                  <div style={{ flex:1, height:7, background:C.bgSoft, borderRadius:4, overflow:"hidden" }}>
                    <div style={{ width:`${pct}%`, height:"100%", background:"#F5A623", borderRadius:4 }} />
                  </div>
                  <span style={{ fontSize:10, color:C.textLight, width:24, textAlign:"right" }}>{count}</span>
                </div>
              );
            })}
          </div>
        </div>

        {/* 我的評價輸入區 */}
        <div style={{ padding:"16px 20px", background:C.goldPale, borderRadius:12, border:`1px solid ${C.gold}40`, marginBottom:18 }}>
          <p style={{ margin:"0 0 10px", fontSize:14, fontWeight:600, color:C.navy }}>{myReview ? "✏️ 修改我的評價" : "✍️ 撰寫評價"}</p>
          <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
            <span style={{ fontSize:13, color:C.textMid }}>我的評分：</span>
            <StarRating value={myRating} size={26} readonly={false} onChange={setMyRating} />
            {myRating > 0 && <span style={{ fontSize:12, color:C.textLight }}>{myRating} 顆星</span>}
          </div>
          <textarea
            value={myReviewText}
            onChange={e => setMyReviewText(e.target.value)}
            rows={3}
            placeholder="這門課對你有什麼幫助？對講師或內容的建議？（選填）"
            style={{ width:"100%", padding:"10px 12px", borderRadius:8, border:`1px solid ${C.border}`, fontSize:13, resize:"vertical", boxSizing:"border-box", fontFamily:"inherit", outline:"none" }}
          />
          <div style={{ display:"flex", justifyContent:"flex-end", marginTop:10 }}>
            <Btn onClick={submitReview} disabled={savingReview || myRating < 1} variant="gold">
              {savingReview ? "送出中..." : myReview ? "更新評價" : "送出評價"}
            </Btn>
          </div>
        </div>

        {/* 評價列表 */}
        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
          {reviews.length === 0 ? (
            <div style={{ textAlign:"center", padding:36, color:C.textLight, background:"#FFF", borderRadius:12, border:`1px solid ${C.border}` }}>
              <p style={{ fontSize:32, margin:0 }}>💬</p>
              <p style={{ fontSize:13, margin:"8px 0 0" }}>還沒有人評價，當第一個分享心得的人吧！</p>
            </div>
          ) : (
            [...reviews]
              .sort((a,b) => {
                // 自己的評價排最前，其餘依有幫助數排序
                if (a.userId === currentUser.id) return -1;
                if (b.userId === currentUser.id) return 1;
                return (b.helpfulUserIds?.length || 0) - (a.helpfulUserIds?.length || 0);
              })
              .map(r => {
                const isMine = r.userId === currentUser.id;
                const helpfulCount = r.helpfulUserIds?.length || 0;
                const iFoundHelpful = r.helpfulUserIds?.includes(currentUser.id);
                return (
                  <div key={r.id} style={{ padding:"14px 16px", background:"#FFF", borderRadius:12, border:`1px solid ${isMine ? C.gold+"60" : C.border}` }}>
                    <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:8 }}>
                      <div style={{ width:34, height:34, borderRadius:"50%", background:`linear-gradient(135deg, ${C.navy}, ${C.navyLight})`, display:"flex", alignItems:"center", justifyContent:"center", color:"#FFF", fontSize:14, fontWeight:600, flexShrink:0 }}>{(r.userName||"匿")[0]}</div>
                      <div style={{ flex:1, minWidth:0 }}>
                        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                          <span style={{ fontSize:13, fontWeight:600, color:C.text }}>{maskName(r.userName)}</span>
                          {isMine && <span style={{ fontSize:10, padding:"1px 7px", borderRadius:8, background:`${C.gold}20`, color:C.navy, fontWeight:600 }}>我的評價</span>}
                        </div>
                        <StarRating value={r.rating} size={13} readonly />
                      </div>
                      {r.date?.toDate && <span style={{ fontSize:10, color:C.textLight }}>{r.date.toDate().toLocaleDateString("zh-TW")}</span>}
                    </div>
                    {r.content && <p style={{ margin:"0 0 10px", fontSize:13, color:C.text, lineHeight:1.6 }}>{r.content}</p>}
                    {/* 有幫助按鈕 */}
                    <button
                      onClick={() => clickHelpful(r)}
                      style={{ display:"inline-flex", alignItems:"center", gap:6, padding:"5px 12px", borderRadius:16, border:`1px solid ${iFoundHelpful ? C.success : C.border}`, background: iFoundHelpful ? `${C.success}10` : "#FFF", color: iFoundHelpful ? C.success : C.textMid, fontSize:12, cursor:"pointer", fontWeight:500, transition:"all 0.2s" }}
                    >
                      {iFoundHelpful ? "👍 已說讚" : "👍 有幫助"}
                      {helpfulCount > 0 && <span style={{ fontSize:11, opacity:0.8 }}>· {helpfulCount} 人覺得有幫助</span>}
                    </button>
                  </div>
                );
              })
          )}
        </div>
      </div>

      {/* ══════════ 「我想問問題」彈窗 ══════════ */}
      {showAskModal && (
        <div onClick={() => setShowAskModal(false)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:20 }}>
          <div onClick={e => e.stopPropagation()} style={{ background:"#FFF", borderRadius:14, width:"100%", maxWidth:480, maxHeight:"90vh", overflow:"auto", boxShadow:"0 24px 48px rgba(0,0,0,0.3)" }}>
            <div style={{ padding:"18px 20px", borderBottom:`1px solid ${C.border}`, background:`linear-gradient(135deg, #FF8C42, #FFA62B)` }}>
              <h3 style={{ fontSize:17, fontWeight:700, color:"#FFF", margin:0, display:"flex", alignItems:"center", gap:8 }}>✋ Q&A 提問</h3>
              <p style={{ fontSize:12, color:"rgba(255,255,255,0.9)", margin:"4px 0 0" }}>《{course.title}》</p>
            </div>
            <div style={{ padding:"20px" }}>
              <div style={{ padding:"8px 12px", background:`${C.accent}10`, borderRadius:7, fontSize:11, color:C.navy, marginBottom:16, lineHeight:1.6 }}>
                💡 您的問題會送給課程講師/小幫手，回覆後會在「✉️ 我的信箱」通知您。
              </div>
              <label style={{ display:"block", color:C.textMid, fontSize:12, marginBottom:4, fontWeight:500 }}>問題主旨</label>
              <input value={askSubject} onChange={e => setAskSubject(e.target.value)} placeholder="例如：第二章的計算公式不太懂" style={{ ...inp, marginBottom:14 }} maxLength={50} />
              <label style={{ display:"block", color:C.textMid, fontSize:12, marginBottom:4, fontWeight:500 }}>問題內容</label>
              <textarea value={askContent} onChange={e => setAskContent(e.target.value)} rows={5} placeholder="請詳細描述您的問題..." style={{ width:"100%", padding:"10px 12px", borderRadius:8, border:`1px solid ${C.border}`, fontSize:13, resize:"vertical", boxSizing:"border-box", fontFamily:"inherit", outline:"none" }} maxLength={500} />
              <div style={{ display:"flex", gap:8, marginTop:18, justifyContent:"flex-end" }}>
                <Btn onClick={() => setShowAskModal(false)} variant="outline">取消</Btn>
                <Btn onClick={submitQuestion} disabled={askingSubmit} style={{ background:"linear-gradient(135deg, #FF8C42, #FFA62B)", border:"none", color:"#FFF" }}>{askingSubmit ? "送出中..." : "送出問題"}</Btn>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function Quiz({ course, goBack, saveQuiz }) {
  const [answers, setAnswers] = useState({});
  const [submitted, setSubmitted] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [showReview, setShowReview] = useState(false);  // 控制是否顯示詳解

  const score = submitted ? course.quiz.reduce((s,q,i) => s + (answers[i]===q.answer?1:0), 0) : 0;
  const pct = Math.round(score/course.quiz.length*100);

  const submit = async () => {
    let s=0; course.quiz.forEach((q,i) => { if(answers[i]===q.answer) s++; });
    setSubmitting(true);
    try {
      // 也存學員的作答和題目快照
      await saveQuiz(course.id, s, course.quiz.length, answers, course.quiz);
      setSubmitted(true);
    } catch (e) {
      alert("儲存測驗結果失敗：" + e.message);
    }
    setSubmitting(false);
  };

  const reset = () => {
    setAnswers({});
    setSubmitted(false);
    setShowReview(false);
  };

  // 進入測驗頁時，掛上全域鍵盤攔截器（離開頁面時自動移除）
  useEffect(() => {
    const blockKey = (e) => {
      // 在「未提交」狀態下才需要強保護
      if (submitted) return;
      // Ctrl/Cmd + C/A/P/S/U/X/V
      if (e.ctrlKey || e.metaKey) {
        const k = e.key.toLowerCase();
        if (["c", "a", "p", "s", "u", "x"].includes(k)) {
          e.preventDefault();
        }
        // Ctrl+Shift+I/J/C 開發者工具
        if (e.shiftKey && ["i", "j", "c"].includes(k)) {
          e.preventDefault();
        }
      }
      // F12
      if (e.key === "F12") e.preventDefault();
      // PrintScreen 無法被網頁攔截（瀏覽器層級），但仍試
      if (e.key === "PrintScreen") e.preventDefault();
    };
    const blockCopy = (e) => { if (!submitted) e.preventDefault(); };
    const blockContext = (e) => { if (!submitted) e.preventDefault(); };

    document.addEventListener("keydown", blockKey);
    document.addEventListener("copy", blockCopy);
    document.addEventListener("cut", blockCopy);
    document.addEventListener("contextmenu", blockContext);

    return () => {
      document.removeEventListener("keydown", blockKey);
      document.removeEventListener("copy", blockCopy);
      document.removeEventListener("cut", blockCopy);
      document.removeEventListener("contextmenu", blockContext);
    };
  }, [submitted]);

  return (
    <div
      style={{
        padding:"24px 20px",
        maxWidth:720,
        margin:"0 auto",
        userSelect: submitted ? "auto" : "none",
        WebkitUserSelect: submitted ? "auto" : "none",
        MozUserSelect: submitted ? "auto" : "none",
        msUserSelect: submitted ? "auto" : "none",
      }}
      onDragStart={(e) => { if (!submitted) e.preventDefault(); }}
    >
      <button onClick={goBack} style={{ border:"none", background:"none", color:C.navy, fontSize:12, cursor:"pointer", padding:0, fontWeight:500, marginBottom:12 }}>← 返回課程</button>
      <h2 style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:4 }}>📝 {course.title} — 課後測驗</h2>
      <p style={{ color:C.textLight, fontSize:12, marginBottom:18 }}>共 {course.quiz.length} 題，及格 60%</p>
      {!submitted && (
        <div style={{ padding:"8px 12px", background:`${C.gold}10`, borderRadius:6, fontSize:11, color:C.navy, marginBottom:14, lineHeight:1.6 }}>
          🔒 <strong>測驗中：</strong>此頁面已開啟保護模式，禁止複製、列印、右鍵
        </div>
      )}

      {submitted ? (
        <>
          {/* 結果摘要 */}
          <div style={{ textAlign:"center", padding:28, background:"#FFF", borderRadius:12, border:`1px solid ${C.border}`, marginBottom:14 }}>
            <span style={{ fontSize:48 }}>{pct>=60?"🎉":"📖"}</span>
            <h3 style={{ fontSize:22, fontWeight:700, color:C.text, marginTop:12 }}>{pct>=60?"恭喜通過！":"再接再厲！"}</h3>
            <p style={{ fontSize:15, color:C.textMid, marginTop:4 }}>得分：<strong style={{ color:pct>=60?C.success:C.danger }}>{score}/{course.quiz.length}</strong>（{pct}%）</p>
            <div style={{ display:"flex", gap:8, justifyContent:"center", marginTop:18, flexWrap:"wrap" }}>
              <Btn onClick={() => setShowReview(!showReview)} variant={showReview?"outline":"gold"}>
                {showReview ? "🔼 收合詳解" : "🔍 查看答題詳解"}
              </Btn>
              <Btn onClick={reset} variant="outline">重新測驗</Btn>
              <Btn onClick={goBack}>返回課程</Btn>
            </div>
          </div>

          {/* 答題詳解 */}
          {showReview && (
            <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
              <div style={{ padding:"10px 14px", background:`${C.accent}10`, borderRadius:8, fontSize:11, color:C.navy, lineHeight:1.7 }}>
                💡 <strong>詳解說明</strong>：✅ 綠底為正確答案，❌ 紅底為您選錯的答案，⚪ 灰底為其他選項。
              </div>
              {course.quiz.map((q, qi) => {
                const userAnswer = answers[qi];
                const isCorrect = userAnswer === q.answer;
                return (
                  <div key={qi} style={{ background:"#FFF", borderRadius:9, padding:16, border:`1px solid ${isCorrect?C.success:C.danger}40` }}>
                    <div style={{ display:"flex", alignItems:"flex-start", gap:8, marginBottom:10 }}>
                      <span style={{ fontSize:14, padding:"3px 9px", borderRadius:6, background:isCorrect?`${C.success}15`:`${C.danger}15`, color:isCorrect?C.success:C.danger, fontWeight:600, flexShrink:0 }}>
                        {isCorrect ? "✓ 答對" : "✗ 答錯"}
                      </span>
                      <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text, flex:1 }}>{qi+1}. {q.q}</p>
                    </div>
                    <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
                      {q.options.map((opt, oi) => {
                        const isUserPick = userAnswer === oi;
                        const isCorrectAnswer = q.answer === oi;
                        let style = { background:"transparent", border:`1px solid ${C.border}`, color:C.text };
                        let prefix = ["A", "B", "C", "D"][oi];
                        let suffix = null;

                        if (isCorrectAnswer) {
                          style = { background:`${C.success}10`, border:`1.5px solid ${C.success}`, color:C.text };
                          suffix = <span style={{ fontSize:11, color:C.success, fontWeight:600, marginLeft:"auto", flexShrink:0 }}>✓ 正確答案</span>;
                        } else if (isUserPick) {
                          style = { background:`${C.danger}10`, border:`1.5px solid ${C.danger}`, color:C.text };
                          suffix = <span style={{ fontSize:11, color:C.danger, fontWeight:600, marginLeft:"auto", flexShrink:0 }}>✗ 您的答案</span>;
                        }

                        return (
                          <div key={oi} style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:6, ...style }}>
                            <span style={{ width:22, height:22, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:600, color: isCorrectAnswer?C.success: isUserPick?C.danger:C.textLight, border:`1.5px solid ${isCorrectAnswer?C.success: isUserPick?C.danger:C.border}`, flexShrink:0 }}>{prefix}</span>
                            <span style={{ fontSize:12 }}>{opt}</span>
                            {suffix}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </>
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
  const [questions, setQuestions] = useState([]);  // 所有提問
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
      watchAllQuestions(setQuestions),
    ];
    return () => unsubs.forEach(u => u());
  }, []);

  // 待回覆問題數（紅點通知）
  const pendingQuestionsCount = questions.filter(q => q.status === "pending").length;

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
    { id:"questions", label:"問答管理", icon:"🙋", badge: pendingQuestionsCount },
  ];

  if (loading) return <LoadingScreen text="載入後台資料..." />;

  return (
    <div style={{ display:"flex", minHeight:"100vh", background:C.bg, fontFamily:"'Noto Sans TC',sans-serif" }}>
      {showPwModal && <ChangePasswordModal currentUser={currentUser} onClose={() => setShowPwModal(false)} />}
      <div style={{ width:200, background:C.navy, color:"#FFF", display:"flex", flexDirection:"column", flexShrink:0 }}>
        <div style={{ padding:"16px 14px", display:"flex", alignItems:"center", gap:10, borderBottom:"1px solid rgba(255,255,255,0.08)" }}>
          <LKLogo size={22} color={C.gold} />
          <div><p style={{ margin:0, fontSize:13, fontWeight:600 }}>管理後台</p><p style={{ margin:0, fontSize:9, opacity:0.4 }}>L&K Admin</p></div>
        </div>
        <div style={{ flex:1, padding:"8px 6px" }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{ width:"100%", padding:"8px 10px", borderRadius:5, border:"none", background:tab===t.id?"rgba(255,255,255,0.12)":"transparent", color:tab===t.id?"#FFF":"rgba(255,255,255,0.5)", fontSize:12, cursor:"pointer", textAlign:"left", display:"flex", alignItems:"center", gap:7, marginBottom:2, fontWeight:tab===t.id?500:400 }}>
              <span style={{ fontSize:13 }}>{t.icon}</span>
              <span style={{ flex:1 }}>{t.label}</span>
              {t.badge > 0 && <span style={{ fontSize:10, minWidth:18, height:18, padding:"0 5px", borderRadius:9, background:C.danger, color:"#FFF", display:"flex", alignItems:"center", justifyContent:"center", fontWeight:600 }}>{t.badge}</span>}
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
        {tab==="questions" && <QuestionAdmin questions={questions} courses={courses} />}
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
  const [helperEmail, setHelperEmail] = useState("");  // 課程小幫手信箱
  const [description, setDescription] = useState("");
  const [coverUrl, setCoverUrl] = useState("");       // 封面圖網址
  const [coverColor, setCoverColor] = useState("#2C5F7C");  // 沒圖時的封面底色
  const [contentType, setContentType] = useState("video");  // video（影片）/ article（文章）
  const [articleContent, setArticleContent] = useState("");  // 文章型課程的內文
  const [chapters, setChapters] = useState([{ title:"第一章", duration:15, youtubeUrl:"" }]);
  const [quiz, setQuiz] = useState([]);  // 測驗題目
  const [files, setFiles] = useState([]);  // 課程附件（連結方式）
  const [saving, setSaving] = useState(false);
  const coverFileRef = useRef(null);  // 封面圖上傳

  // 圖片上傳處理（demo 階段：用瀏覽器本地預覽 base64）
  // 未來換成公司伺服器：改成上傳到伺服器後取得網址
  const handleCoverUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) { alert("請選擇圖片檔案"); return; }
    if (file.size > 800 * 1024) {
      alert("⚠️ Demo 階段限制：圖片需小於 800 KB。\n\n正式版改用公司伺服器後就沒有此限制。\n建議先用「圖片網址」方式，或壓縮圖片後再上傳。");
      return;
    }
    const reader = new FileReader();
    reader.onload = (ev) => {
      setCoverUrl(ev.target.result);  // base64 字串
    };
    reader.readAsDataURL(file);
  };

  const reset = () => {
    setTitle(""); setCategory(categories[0]?.id || ""); setInstructor(""); setHelperEmail(""); setDescription("");
    setCoverUrl(""); setCoverColor("#2C5F7C");
    setContentType("video"); setArticleContent("");
    setChapters([{ title:"第一章", duration:15, youtubeUrl:"" }]);
    setQuiz([]);
    setFiles([]);
    setEditing(null); setShowForm(false);
  };
  const startEdit = (c) => {
    setTitle(c.title); setCategory(c.category); setInstructor(c.instructor); setHelperEmail(c.helperEmail || ""); setDescription(c.description);
    setCoverUrl(c.coverUrl || ""); setCoverColor(c.coverColor || "#2C5F7C");
    setContentType(c.contentType || "video"); setArticleContent(c.articleContent || "");
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
    // 影片型：時長=各章節加總；文章型：估算閱讀時間（每 400 字約 1 分鐘）
    const totalDuration = contentType === "article"
      ? Math.max(1, Math.round(articleContent.length / 400))
      : chapters.reduce((s,c) => s + (+c.duration||0), 0);
    const courseData = { title, category, instructor, helperEmail, description, coverUrl, coverColor, contentType, articleContent, chapters, quiz, files, duration: totalDuration };
    try {
      if (editing) {
        await updateCourse(editing, courseData);
      } else {
        await addCourse({
          ...courseData,
          thumbnail:"📘", views:0,
          publishDate: new Date().toISOString().split("T")[0],
          status: publishNow ? "published" : "draft",
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
            <Field label="課程小幫手信箱"><input value={helperEmail} onChange={e => setHelperEmail(e.target.value)} style={inp} placeholder="收問題通知的信箱" /></Field>
            <Field label="分類">
              <select value={category} onChange={e => setCategory(e.target.value)} style={inp}>
                {categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </Field>
          </div>
          <Field label="課程說明"><textarea value={description} onChange={e => setDescription(e.target.value)} rows={2} style={{ ...inp, resize:"vertical" }} placeholder="課程內容說明..." /></Field>

          {/* ══════ 課程類型選擇 ══════ */}
          <div style={{ marginTop:14, padding:12, background:C.bg, borderRadius:8 }}>
            <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:C.text }}>📂 課程類型</p>
            <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
              <button onClick={() => setContentType("video")} style={{ flex:1, minWidth:140, padding:"12px", borderRadius:8, border:`2px solid ${contentType==="video"?C.navy:C.border}`, background:contentType==="video"?`${C.navy}08`:"#FFF", cursor:"pointer", textAlign:"left" }}>
                <div style={{ fontSize:20, marginBottom:4 }}>🎬</div>
                <div style={{ fontSize:13, fontWeight:600, color:contentType==="video"?C.navy:C.text }}>影片課程</div>
                <div style={{ fontSize:10, color:C.textLight, marginTop:2 }}>章節 + 影片播放</div>
              </button>
              <button onClick={() => setContentType("article")} style={{ flex:1, minWidth:140, padding:"12px", borderRadius:8, border:`2px solid ${contentType==="article"?C.navy:C.border}`, background:contentType==="article"?`${C.navy}08`:"#FFF", cursor:"pointer", textAlign:"left" }}>
                <div style={{ fontSize:20, marginBottom:4 }}>📄</div>
                <div style={{ fontSize:13, fontWeight:600, color:contentType==="article"?C.navy:C.text }}>文章課程</div>
                <div style={{ fontSize:10, color:C.textLight, marginTop:2 }}>圖文閱讀內容</div>
              </button>
            </div>
          </div>

          {/* ══════ 課程封面設定 ══════ */}
          <div style={{ marginTop:14, padding:12, background:C.bg, borderRadius:8 }}>
            <p style={{ margin:"0 0 10px", fontSize:13, fontWeight:600, color:C.text }}>🖼️ 課程封面</p>
            <div style={{ display:"flex", gap:14, flexWrap:"wrap" }}>
              {/* 即時預覽 */}
              <div style={{ width:200, flexShrink:0 }}>
                <p style={{ fontSize:11, color:C.textLight, margin:"0 0 5px" }}>預覽</p>
                <div style={{ width:200, height:112, borderRadius:8, overflow:"hidden", border:`1px solid ${C.border}`, background: coverUrl ? "#000" : `linear-gradient(135deg, ${coverColor}, ${coverColor}DD)`, display:"flex", alignItems:"center", justifyContent:"center", position:"relative" }}>
                  {coverUrl ? (
                    <img src={coverUrl} alt="封面預覽" style={{ width:"100%", height:"100%", objectFit:"cover" }} onError={(e) => { e.target.style.display="none"; }} />
                  ) : (
                    <div style={{ textAlign:"center", color:"#FFF", padding:10 }}>
                      <div style={{ fontSize:28, marginBottom:4 }}>📘</div>
                      <div style={{ fontSize:12, fontWeight:600, opacity:0.95, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{title || "課程標題"}</div>
                    </div>
                  )}
                </div>
              </div>
              {/* 設定欄位 */}
              <div style={{ flex:1, minWidth:240 }}>
                {/* 上傳圖片按鈕 */}
                <div style={{ marginBottom:12 }}>
                  <label style={{ display:"block", color:C.textMid, fontSize:12, marginBottom:4, fontWeight:500 }}>方式一：上傳圖片</label>
                  <input
                    ref={coverFileRef}
                    type="file"
                    accept="image/*"
                    onChange={handleCoverUpload}
                    style={{ display:"none" }}
                  />
                  <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
                    <Btn onClick={() => coverFileRef.current?.click()} variant="outline" style={{ fontSize:12 }}>
                      📤 選擇圖片上傳
                    </Btn>
                    {coverUrl && coverUrl.startsWith("data:") && (
                      <span style={{ fontSize:11, color:C.success }}>✅ 已選擇圖片</span>
                    )}
                  </div>
                  <p style={{ fontSize:10, color:C.warning, margin:"5px 0 0", lineHeight:1.6 }}>
                    ⚠️ Demo 階段：上傳圖片需小於 800 KB（暫存於資料庫）。<br />
                    未來改用公司伺服器後，可上傳大圖且無容量限制。
                  </p>
                </div>

                <div style={{ height:1, background:C.border, margin:"12px 0" }} />

                <Field label="方式二：貼圖片網址">
                  <input value={coverUrl.startsWith("data:") ? "" : coverUrl} onChange={e => setCoverUrl(e.target.value)} placeholder="https://圖片網址.jpg" style={inp} disabled={coverUrl.startsWith("data:")} />
                </Field>
                <p style={{ fontSize:10, color:C.textLight, margin:"-4px 0 10px", lineHeight:1.6 }}>
                  💡 建議 16:9 比例（如 1280×720）。圖片需設為公開可存取。
                </p>

                {coverUrl && (
                  <Btn onClick={() => setCoverUrl("")} variant="danger" style={{ fontSize:11, padding:"4px 10px", marginBottom:10 }}>
                    ✕ 清除封面圖
                  </Btn>
                )}

                <Field label="底色（沒有封面圖時顯示）">
                  <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                    <input type="color" value={coverColor} onChange={e => setCoverColor(e.target.value)} style={{ width:40, height:32, border:"none", padding:0, borderRadius:5, cursor:"pointer", background:"transparent" }} />
                    <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
                      {["#2C5F7C","#D4A528","#22A06B","#7C3AED","#E25555","#0891B2","#B45309","#475569"].map(co => (
                        <button key={co} onClick={() => setCoverColor(co)} style={{ width:24, height:24, borderRadius:"50%", border:`2px solid ${coverColor===co?"#333":"transparent"}`, background:co, cursor:"pointer", padding:0 }} />
                      ))}
                    </div>
                  </div>
                </Field>
              </div>
            </div>
          </div>

          {/* 文章型課程：內文編輯區 */}
          {contentType === "article" && (
            <div style={{ marginTop:14, padding:12, background:C.bg, borderRadius:8 }}>
              <p style={{ margin:"0 0 8px", fontSize:13, fontWeight:600, color:C.text }}>📄 文章內容</p>
              <div style={{ padding:"8px 10px", background:`${C.gold}10`, borderRadius:6, fontSize:11, color:C.navy, marginBottom:10, lineHeight:1.7 }}>
                💡 <strong>排版小技巧：</strong><br />
                • 空一行 = 分段落<br />
                • 用「## 標題」開頭 = 顯示為小標題<br />
                • 貼圖片網址（以 http 開頭、.jpg/.png 結尾，單獨一行）= 自動顯示圖片
              </div>
              <textarea
                value={articleContent}
                onChange={e => setArticleContent(e.target.value)}
                rows={14}
                placeholder={"在這裡輸入文章內容...\n\n## 第一段標題\n\n這是內文，可以寫很多字。空一行就會分段。\n\nhttps://example.com/圖片.jpg\n\n## 第二段標題\n\n繼續寫內容..."}
                style={{ width:"100%", padding:"12px", borderRadius:8, border:`1px solid ${C.border}`, fontSize:13, lineHeight:1.7, resize:"vertical", boxSizing:"border-box", fontFamily:"inherit", outline:"none" }}
              />
              <p style={{ fontSize:10, color:C.textLight, margin:"6px 0 0" }}>
                目前約 {articleContent.length} 字，預估閱讀時間 {Math.max(1, Math.round(articleContent.length / 400))} 分鐘
              </p>
            </div>
          )}

          {/* 影片型課程：章節設定區 */}
          {contentType === "video" && (
          <div style={{ marginTop:14, padding:12, background:C.bg, borderRadius:8 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
              <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text }}>📋 章節設定（每章節對應一支影片）</p>
              <Btn onClick={addChapter} variant="outline" style={{ padding:"4px 10px", fontSize:11 }}>+ 新增章節</Btn>
            </div>
            <div style={{ padding:"8px 10px", background:`${C.gold}10`, borderRadius:6, fontSize:11, color:C.navy, marginBottom:10, lineHeight:1.7 }}>
              💡 <strong>支援的影片來源：</strong><br />
              <strong>▶ YouTube</strong>：貼上影片網址（公開或不公開且允許嵌入）<br />
              <strong>☁ OneDrive / SharePoint</strong>：在影片上點「Embed 嵌入」，複製 iframe 嵌入碼或 src 網址<br />
              <strong>▶ Vimeo</strong>：貼上影片網址<br />
              <strong>🔗 其他連結</strong>：會以「新分頁播放」按鈕呈現
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
                <input value={ch.youtubeUrl} onChange={e => updateChapter(idx, "youtubeUrl", e.target.value)} placeholder="🎬 影片網址（YouTube / OneDrive 嵌入碼 / SharePoint 等）" style={inp} />
                {ch.youtubeUrl && (() => {
                  const info = parseVideoUrl(ch.youtubeUrl);
                  if (!info) return <p style={{ fontSize:10, color:C.danger, margin:"4px 0 0" }}>⚠️ 網址格式無法識別</p>;
                  const labels = {
                    youtube: { name:"YouTube", icon:"▶", canEmbed:true },
                    onedrive: { name:"OneDrive 嵌入", icon:"☁", canEmbed:true },
                    sharepoint: { name:"SharePoint", icon:"☁", canEmbed:true },
                    vimeo: { name:"Vimeo", icon:"▶", canEmbed:true },
                    external: { name:"外部連結（新分頁開啟）", icon:"🔗", canEmbed:false },
                  };
                  const lbl = labels[info.type];
                  return (
                    <p style={{ fontSize:10, color:lbl.canEmbed?C.success:C.warning, margin:"4px 0 0" }}>
                      {lbl.canEmbed ? "✅" : "⚠️"} 已偵測：{lbl.icon} {lbl.name}
                      {!lbl.canEmbed && <span style={{ color:C.textLight, marginLeft:6 }}>（無法直接嵌入，將顯示「新分頁播放」按鈕）</span>}
                    </p>
                  );
                })()}
              </div>
            ))}
          </div>
          )}

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
  const [division, setDivision] = useState("");   // 部別
  const [group, setGroup] = useState("");          // 組別
  const [role, setRole] = useState("user");
  const [managerScope, setManagerScope] = useState("");  // 主管管轄範圍
  const [saving, setSaving] = useState(false);
  const [showWarning, setShowWarning] = useState(true);
  // 篩選 / 搜尋 / 分頁 / 檢視模式
  const [statusTab, setStatusTab] = useState("active");   // active/suspended/inactive
  const [searchKw, setSearchKw] = useState("");
  const [filterDept, setFilterDept] = useState("");
  const [viewMode, setViewMode] = useState("list");        // list / org（組織架構）
  const [moveModal, setMoveModal] = useState(null);        // 異動視窗的對象 user

  const reset = () => { setEmpNo(""); setName(""); setEmail(""); setPassword(""); setDepartment(""); setDivision(""); setGroup(""); setRole("user"); setManagerScope(""); setEditId(null); setShowForm(false); };
  const startEdit = (u) => { setEmpNo(u.empNo||""); setName(u.name); setEmail(u.email); setPassword(""); setDepartment(u.department||""); setDivision(u.division||""); setGroup(u.group||""); setRole(u.role||"user"); setManagerScope(u.managerScope||""); setEditId(u.id); setShowForm(true); };

  const save = async () => {
    if (!name.trim() || !email.trim() || !empNo.trim()) { alert("員工編號、姓名、Email 為必填"); return; }
    setSaving(true);
    try {
      if (editId) {
        await updateUserData(editId, { empNo, name, department, division, group, role, managerScope });
        reset();
      } else {
        const finalPw = password.trim() || empNo;
        await createUserAccount(empNo, name, email, finalPw, role, department, division, group);
        if (managerScope) {
          // 新建後若設定了主管範圍，需要再寫一次（createUserAccount 沒帶 managerScope）
          // 但這時已切換帳號，故提示用編輯方式設定
        }
        alert("✅ 使用者建立成功！\n\n⚠️ 重要：建立新使用者會自動切換到該帳號登入，請重新登入管理員帳號繼續操作。");
        reset();
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

  // 復職 / 解鎖
  const handleReactivate = async (u) => {
    const action = u.status === "suspended" ? "復職" : "重新啟用";
    if (!confirm(`確定要將 ${u.name} ${action}嗎？\n\n帳號將恢復為「使用中」，該員工可以重新登入。`)) return;
    try { await reactivateUser(u.id); } catch (e) { alert("操作失敗：" + e.message); }
  };

  // 角色標籤
  const roleLabel = (r) => r==="superadmin" ? "系統管理員" : r==="admin" ? "管理員" : "使用者";
  const roleColor = (r) => r==="superadmin" ? C.danger : r==="admin" ? C.navy : C.accent;

  // 狀態分頁定義
  const statusTabs = [
    { id:"active", label:"使用中", icon:"✅" },
    { id:"suspended", label:"已暫停", icon:"⏸️" },
    { id:"inactive", label:"已停用", icon:"🔒" },
  ];

  // 處別清單（用於篩選下拉）
  const deptOptions = [...new Set(users.map(u => u.department).filter(Boolean))];

  // 篩選後的使用者
  const filteredUsers = users.filter(u => {
    const st = u.status || "active";
    if (st !== statusTab) return false;
    if (filterDept && u.department !== filterDept) return false;
    if (searchKw) {
      const kw = searchKw.trim().toLowerCase();
      const hay = `${u.name||""} ${u.empNo||""} ${u.department||""} ${u.division||""} ${u.group||""}`.toLowerCase();
      if (!hay.includes(kw)) return false;
    }
    return true;
  });

  // 各狀態人數（分頁籤顯示）
  const statusCount = (st) => users.filter(u => (u.status||"active") === st).length;

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
            <Field label="部別"><input value={division} onChange={e => setDivision(e.target.value)} style={inp} placeholder="例：人資部（選填）" /></Field>
            <Field label="組別"><input value={group} onChange={e => setGroup(e.target.value)} style={inp} placeholder="少數單位才有（選填）" /></Field>
            <Field label="角色">
              <select value={role} onChange={e => setRole(e.target.value)} style={inp}>
                <option value="user">使用者</option>
                <option value="admin">管理員</option>
                <option value="superadmin">系統管理員</option>
              </select>
            </Field>
            {(role==="admin") && (
              <Field label="主管管轄範圍（報表用）">
                <select value={managerScope} onChange={e => setManagerScope(e.target.value)} style={inp}>
                  <option value="">不是主管</option>
                  <option value="division">部級主管（看本部）</option>
                  <option value="department">處級主管（看本處）</option>
                </select>
              </Field>
            )}
          </div>
          <div style={{ display:"flex", gap:6, marginTop:10 }}>
            <Btn onClick={save} variant="gold" disabled={saving}>{saving?"處理中...":(editId?"儲存":"建立")}</Btn>
            <Btn onClick={reset} variant="outline">取消</Btn>
          </div>
        </div>
      )}

      {/* 狀態分頁籤 */}
      <div style={{ display:"flex", gap:4, marginBottom:14, borderBottom:`1px solid ${C.border}`, flexWrap:"wrap" }}>
        {statusTabs.map(t => (
          <button key={t.id} onClick={() => setStatusTab(t.id)} style={{ padding:"9px 16px", border:"none", background:"transparent", borderBottom:`3px solid ${statusTab===t.id?C.gold:"transparent"}`, color:statusTab===t.id?C.navy:C.textLight, fontSize:13, fontWeight:statusTab===t.id?700:500, cursor:"pointer", marginBottom:-1, display:"flex", alignItems:"center", gap:6 }}>
            {t.icon} {t.label}
            <span style={{ fontSize:11, minWidth:18, height:18, padding:"0 5px", borderRadius:9, background:statusTab===t.id?C.gold:C.bgSoft, color:statusTab===t.id?"#FFF":C.textLight, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:600 }}>{statusCount(t.id)}</span>
          </button>
        ))}
      </div>

      {/* 篩選列：搜尋 + 處別下拉 */}
      <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap" }}>
        <input value={searchKw} onChange={e => setSearchKw(e.target.value)} placeholder="🔍 搜尋姓名、員工編號、處別、部別..." style={{ ...inp, flex:1, minWidth:200 }} />
        <select value={filterDept} onChange={e => setFilterDept(e.target.value)} style={{ ...inp, maxWidth:180 }}>
          <option value="">全部處別</option>
          {deptOptions.map(d => <option key={d} value={d}>{d}</option>)}
        </select>
        {(searchKw || filterDept) && <Btn onClick={() => { setSearchKw(""); setFilterDept(""); }} variant="outline" style={{ fontSize:11 }}>✕ 清除</Btn>}
      </div>

      <div style={{ background:"#FFF", borderRadius:9, border:`1px solid ${C.border}`, overflow:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", minWidth:720 }}>
          <thead><tr style={{ borderBottom:`1px solid ${C.border}` }}>
            {["員工編號","姓名","處別 / 部別","角色","密碼","操作"].map(h => <th key={h} style={{ padding:"9px 10px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
          </tr></thead>
          <tbody>
            {filteredUsers.length === 0 ? (
              <tr><td colSpan={6} style={{ padding:36, textAlign:"center", color:C.textLight, fontSize:13 }}>沒有符合條件的使用者</td></tr>
            ) : filteredUsers.map(u => (
              <tr key={u.id} style={{ borderBottom:`1px solid ${C.border}` }}>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.text, fontFamily:"monospace" }}>{u.empNo||"—"}</td>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.text, fontWeight:500 }}>{u.name}</td>
                <td style={{ padding:"8px 10px", fontSize:12, color:C.textMid }}>
                  {u.department || "—"}
                  {u.division && <span style={{ color:C.textLight }}> / {u.division}</span>}
                  {u.group && <span style={{ color:C.textLight }}> / {u.group}</span>}
                  {/* 待生效異動標記 */}
                  {u.pendingChange?.effectiveDate && (
                    <div style={{ fontSize:10, color:C.warning, marginTop:3 }}>
                      ⏳ {u.pendingChange.effectiveDate} 起{u.pendingChange.type==="transfer"?"調動":u.pendingChange.type==="suspend"?"留停":"離職"}
                    </div>
                  )}
                  {/* 已套用的狀態備註 */}
                  {u.status !== "active" && u.statusNote && (
                    <div style={{ fontSize:10, color:C.textLight, marginTop:3 }}>{u.statusNote}</div>
                  )}
                </td>
                <td style={{ padding:"8px 10px" }}>
                  <span style={{ fontSize:10, padding:"3px 8px", borderRadius:7, background:`${roleColor(u.role)}15`, color:roleColor(u.role), fontWeight:600 }}>{roleLabel(u.role)}</span>
                  {u.managerScope && <span style={{ fontSize:9, padding:"2px 6px", borderRadius:6, background:`${C.gold}18`, color:C.navy, marginLeft:4 }}>{u.managerScope==="department"?"處主管":"部主管"}</span>}
                </td>
                <td style={{ padding:"8px 10px" }}>
                  {u.mustChangePw ? <span style={{ fontSize:10, padding:"3px 7px", borderRadius:7, background:`${C.warning}15`, color:C.warning }}>未改</span> : <span style={{ fontSize:10, color:C.success }}>✓</span>}
                </td>
                <td style={{ padding:"8px 10px" }}>
                  <div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>
                    {u.status !== "suspended" && u.status !== "inactive" && (
                      <>
                        <Btn onClick={() => startEdit(u)} variant="outline" style={{ padding:"3px 7px", fontSize:10 }}>編輯</Btn>
                        <Btn onClick={() => setMoveModal(u)} variant="outline" style={{ padding:"3px 7px", fontSize:10, borderColor:C.warning, color:C.warning }}>異動</Btn>
                      </>
                    )}
                    {(u.status === "suspended" || u.status === "inactive") && (
                      <Btn onClick={() => handleReactivate(u)} variant="outline" style={{ padding:"3px 7px", fontSize:10, borderColor:C.success, color:C.success }}>🔓 復職</Btn>
                    )}
                    <Btn onClick={() => remove(u)} variant="danger" style={{ padding:"3px 7px", fontSize:10 }}>刪除</Btn>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* ══════ 異動視窗 ══════ */}
      {moveModal && (
        <MoveModal user={moveModal} onClose={() => setMoveModal(null)} />
      )}
    </div>
  );
}

/* ─── 人員異動視窗 ─── */
function MoveModal({ user, onClose }) {
  const [type, setType] = useState("transfer");        // transfer / suspend / resign
  const [effectiveDate, setEffectiveDate] = useState(new Date().toISOString().split("T")[0]);
  const [note, setNote] = useState("");
  const [newDepartment, setNewDepartment] = useState(user.department || "");
  const [newDivision, setNewDivision] = useState(user.division || "");
  const [newGroup, setNewGroup] = useState(user.group || "");
  const [saving, setSaving] = useState(false);

  const typeInfo = {
    transfer: { label:"調部門", icon:"🔄", color:C.accent, desc:"調整此人的處別/部別/組別" },
    suspend:  { label:"留職停薪", icon:"⏸️", color:C.warning, desc:"生效日起帳號暫停，可日後復職" },
    resign:   { label:"離職", icon:"🔒", color:C.danger, desc:"生效日起帳號停用" },
  };

  const submit = async (applyNow) => {
    if (!effectiveDate) { alert("請選擇生效日期"); return; }
    setSaving(true);
    const change = { type, effectiveDate, note, newDepartment, newDivision, newGroup };
    try {
      if (applyNow) {
        await applyUserChange(user.id, change);
      } else {
        await scheduleUserChange(user.id, change);
      }
      onClose();
    } catch (e) {
      alert("操作失敗：" + e.message);
    }
    setSaving(false);
  };

  const today = new Date().toISOString().split("T")[0];
  const isFuture = effectiveDate > today;

  return (
    <div onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:20 }}>
      <div onClick={e => e.stopPropagation()} style={{ background:"#FFF", borderRadius:14, width:"100%", maxWidth:440, maxHeight:"90vh", overflow:"auto", boxShadow:"0 24px 48px rgba(0,0,0,0.3)" }}>
        <div style={{ padding:"18px 20px", borderBottom:`1px solid ${C.border}` }}>
          <h3 style={{ fontSize:16, fontWeight:700, color:C.text, margin:0 }}>人員異動</h3>
          <p style={{ fontSize:12, color:C.textLight, margin:"4px 0 0" }}>{user.name}（{user.empNo}）· 目前：{user.department}{user.division?` / ${user.division}`:""}</p>
        </div>
        <div style={{ padding:"20px" }}>
          {/* 異動類型 */}
          <label style={{ display:"block", fontSize:12, color:C.textMid, marginBottom:6, fontWeight:500 }}>異動類型</label>
          <div style={{ display:"flex", gap:8, marginBottom:16, flexWrap:"wrap" }}>
            {Object.entries(typeInfo).map(([k, info]) => (
              <button key={k} onClick={() => setType(k)} style={{ flex:1, minWidth:100, padding:"10px", borderRadius:8, border:`2px solid ${type===k?info.color:C.border}`, background:type===k?`${info.color}10`:"#FFF", cursor:"pointer", textAlign:"center" }}>
                <div style={{ fontSize:18 }}>{info.icon}</div>
                <div style={{ fontSize:12, fontWeight:600, color:type===k?info.color:C.text, marginTop:3 }}>{info.label}</div>
              </button>
            ))}
          </div>
          <p style={{ fontSize:11, color:C.textLight, margin:"-8px 0 16px" }}>{typeInfo[type].desc}</p>

          {/* 調部門才顯示新部門欄位 */}
          {type === "transfer" && (
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:14 }}>
              <Field label="新處別"><input value={newDepartment} onChange={e => setNewDepartment(e.target.value)} style={inp} placeholder="例：工程中心" /></Field>
              <Field label="新部別"><input value={newDivision} onChange={e => setNewDivision(e.target.value)} style={inp} placeholder="選填" /></Field>
              <Field label="新組別"><input value={newGroup} onChange={e => setNewGroup(e.target.value)} style={inp} placeholder="選填" /></Field>
            </div>
          )}

          {/* 生效日期 */}
          <Field label="生效日期">
            <input type="date" value={effectiveDate} onChange={e => setEffectiveDate(e.target.value)} style={inp} />
          </Field>
          <p style={{ fontSize:11, color: isFuture?C.textLight:C.warning, margin:"4px 0 14px", lineHeight:1.5 }}>
            {isFuture
              ? `💡 此異動將於 ${effectiveDate} 由系統自動生效（該員工當天之後登入時觸發）`
              : "⚠️ 生效日為今天或過去，建議直接「立即執行」"}
          </p>

          {/* 備註 */}
          <Field label="備註（選填）">
            <input value={note} onChange={e => setNote(e.target.value)} style={inp} placeholder="調動原因、留停期間等" />
          </Field>

          <div style={{ padding:"10px 12px", background:`${C.accent}08`, borderRadius:7, fontSize:11, color:C.navy, margin:"14px 0", lineHeight:1.6 }}>
            💡 <strong>自動鎖帳號說明：</strong>設定未來日期後，系統會在該員工「生效日當天之後首次登入」時自動套用（留停/離職會擋下登入）。<br />
            <span style={{ color:C.textLight }}>※ 完整的「半夜準時自動鎖定」需公司伺服器排程，未來移轉後即可支援。</span>
          </div>

          <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
            <Btn onClick={onClose} variant="outline">取消</Btn>
            {isFuture
              ? <Btn onClick={() => submit(false)} disabled={saving} variant="gold">{saving?"處理中...":"排定異動"}</Btn>
              : <Btn onClick={() => submit(true)} disabled={saving} style={{ background:typeInfo[type].color, border:"none", color:"#FFF" }}>{saving?"處理中...":"立即執行"}</Btn>
            }
          </div>
        </div>
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
  const [detailRecord, setDetailRecord] = useState(null);  // 查看詳解的紀錄

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

  // 刪除單筆
  const deleteOne = async (record) => {
    const confirmText = `確定要刪除這筆測驗紀錄嗎？\n\n` +
      `員工：${record.userName}（${record.empNo}）\n` +
      `課程：${record.courseName}\n` +
      `分數：${record.score}/${record.total}（${record.pct}%）\n` +
      `日期：${record.dateObj ? record.dateObj.toLocaleString("zh-TW") : "—"}\n\n` +
      `⚠️ 此操作無法復原`;
    if (!confirm(confirmText)) return;
    try {
      await deleteQuizResult(record.key);
      // 也從勾選中移除
      setSelected(p => { const next = {...p}; delete next[record.key]; return next; });
    } catch (e) {
      alert("刪除失敗：" + e.message);
    }
  };

  // 批次刪除
  const deleteSelected = async () => {
    const toDelete = filtered.filter(r => selected[r.key]);
    if (toDelete.length === 0) {
      alert("請先勾選要刪除的紀錄");
      return;
    }
    const confirmText = `確定要刪除以下 ${toDelete.length} 筆測驗紀錄嗎？\n\n` +
      toDelete.slice(0, 5).map(r => `· ${r.userName} - ${r.courseName}（${r.pct}%）`).join("\n") +
      (toDelete.length > 5 ? `\n... 還有 ${toDelete.length - 5} 筆` : "") +
      `\n\n⚠️ 此操作無法復原`;
    if (!confirm(confirmText)) return;
    try {
      const ids = toDelete.map(r => r.key);
      await deleteQuizResultsBatch(ids);
      setSelected({});
    } catch (e) {
      alert("批次刪除失敗：" + e.message);
    }
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>測驗紀錄</h2>
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {(filterUser || filterCourse || filterResult !== "all" || filterDateFrom || filterDateTo) && (
            <Btn onClick={resetFilter} variant="outline" style={{ fontSize:11 }}>✕ 清除篩選</Btn>
          )}
          {selectedCount > 0 && (
            <Btn onClick={deleteSelected} variant="danger" style={{ fontSize:12 }}>
              🗑️ 刪除已勾選（{selectedCount} 筆）
            </Btn>
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
                {["員工編號","姓名","處別","課程","分數","結果","日期","操作"].map(h => <th key={h} style={{ padding:"10px 12px", textAlign:"left", color:C.textLight, fontSize:11, fontWeight:500 }}>{h}</th>)}
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
                  <td style={{ padding:"8px 12px" }}>
                    <div style={{ display:"flex", gap:4 }}>
                      <Btn onClick={() => setDetailRecord(r)} variant="outline" style={{ padding:"3px 8px", fontSize:10 }}>
                        🔍 看答題
                      </Btn>
                      <Btn onClick={() => deleteOne(r)} variant="danger" style={{ padding:"3px 8px", fontSize:10 }}>
                        🗑️ 刪除
                      </Btn>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* 答題詳解彈窗 */}
      {detailRecord && <QuizDetailModal record={detailRecord} courses={courses} onClose={() => setDetailRecord(null)} />}
    </div>
  );
}

/* ─── 後台查看學員答題詳解的彈窗 ─── */
function QuizDetailModal({ record, courses, onClose }) {
  const course = courses.find(c => c.id === record.courseId);
  // 題目來源優先用 snapshot（當下測驗的題目），否則用當前的課程題目
  const questions = (record.questionsSnapshot && record.questionsSnapshot.length > 0) ? record.questionsSnapshot : (course?.quiz || []);
  const userAnswers = record.answers || {};
  const hasSnapshot = record.questionsSnapshot && record.questionsSnapshot.length > 0;

  return (
    <div onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000, padding:20 }}>
      <div onClick={e => e.stopPropagation()} style={{ background:"#FFF", borderRadius:12, padding:0, width:"100%", maxWidth:720, maxHeight:"90vh", overflow:"hidden", boxShadow:"0 24px 48px rgba(0,0,0,0.3)", display:"flex", flexDirection:"column" }}>
        {/* 標題列 */}
        <div style={{ padding:"16px 20px", borderBottom:`1px solid ${C.border}`, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div>
            <h3 style={{ fontSize:16, fontWeight:700, color:C.text, margin:0 }}>🔍 答題詳解</h3>
            <p style={{ fontSize:11, color:C.textLight, margin:"3px 0 0" }}>
              {record.userName}（{record.empNo}）· {record.courseName} · {record.dateObj ? record.dateObj.toLocaleString("zh-TW") : "—"}
            </p>
          </div>
          <button onClick={onClose} style={{ border:"none", background:"none", fontSize:20, cursor:"pointer", color:C.textLight, padding:4 }}>✕</button>
        </div>

        {/* 成績摘要 */}
        <div style={{ padding:"14px 20px", background:C.bgSoft, borderBottom:`1px solid ${C.border}`, display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:10 }}>
          <div style={{ display:"flex", gap:18, alignItems:"center", flexWrap:"wrap" }}>
            <div>
              <p style={{ margin:0, fontSize:10, color:C.textLight }}>得分</p>
              <p style={{ margin:0, fontSize:20, fontWeight:700, color:C.text }}>{record.score}/{record.total}</p>
            </div>
            <div>
              <p style={{ margin:0, fontSize:10, color:C.textLight }}>百分比</p>
              <p style={{ margin:0, fontSize:20, fontWeight:700, color:record.pct>=60?C.success:C.danger }}>{record.pct}%</p>
            </div>
            <div>
              <p style={{ margin:0, fontSize:10, color:C.textLight }}>結果</p>
              <p style={{ margin:0, fontSize:14, fontWeight:600, color:record.pct>=60?C.success:C.danger }}>{record.pct>=60?"✓ 通過":"✗ 未通過"}</p>
            </div>
          </div>
          {!hasSnapshot && (
            <p style={{ margin:0, fontSize:10, color:C.warning, fontStyle:"italic" }}>⚠️ 此紀錄為早期版本，無作答快照，顯示當前題目</p>
          )}
        </div>

        {/* 答題詳解內容（可滾動）*/}
        <div style={{ flex:1, overflowY:"auto", padding:"16px 20px" }}>
          {questions.length === 0 ? (
            <p style={{ textAlign:"center", color:C.textLight, padding:30 }}>無題目資料</p>
          ) : questions.map((q, qi) => {
            const userPick = userAnswers[qi];
            const userPickDefined = userPick !== undefined && userPick !== null;
            const isCorrect = userPick === q.answer;
            return (
              <div key={qi} style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${isCorrect?C.success:C.danger}40`, marginBottom:12 }}>
                <div style={{ display:"flex", alignItems:"flex-start", gap:8, marginBottom:8 }}>
                  <span style={{ fontSize:11, padding:"3px 8px", borderRadius:5, background:isCorrect?`${C.success}15`:`${C.danger}15`, color:isCorrect?C.success:C.danger, fontWeight:600, flexShrink:0 }}>
                    {isCorrect ? "✓ 答對" : "✗ 答錯"}
                  </span>
                  <p style={{ margin:0, fontSize:13, fontWeight:600, color:C.text, flex:1 }}>{qi+1}. {q.q}</p>
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                  {q.options.map((opt, oi) => {
                    const isUserPick = userPick === oi;
                    const isCorrectAnswer = q.answer === oi;
                    let style = { background:"transparent", border:`1px solid ${C.border}`, color:C.text };
                    let suffix = null;
                    const prefix = ["A", "B", "C", "D"][oi];

                    if (isCorrectAnswer) {
                      style = { background:`${C.success}10`, border:`1.5px solid ${C.success}`, color:C.text };
                      suffix = <span style={{ fontSize:10, color:C.success, fontWeight:600, marginLeft:"auto", flexShrink:0 }}>✓ 正確答案</span>;
                    } else if (isUserPick) {
                      style = { background:`${C.danger}10`, border:`1.5px solid ${C.danger}`, color:C.text };
                      suffix = <span style={{ fontSize:10, color:C.danger, fontWeight:600, marginLeft:"auto", flexShrink:0 }}>✗ 學員選擇</span>;
                    }

                    return (
                      <div key={oi} style={{ display:"flex", alignItems:"center", gap:8, padding:"6px 10px", borderRadius:5, ...style }}>
                        <span style={{ width:20, height:20, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, fontWeight:600, color: isCorrectAnswer?C.success: isUserPick?C.danger:C.textLight, border:`1.5px solid ${isCorrectAnswer?C.success: isUserPick?C.danger:C.border}`, flexShrink:0 }}>{prefix}</span>
                        <span style={{ fontSize:12 }}>{opt}</span>
                        {suffix}
                      </div>
                    );
                  })}
                  {!userPickDefined && (
                    <p style={{ margin:"4px 0 0", fontSize:10, color:C.warning, fontStyle:"italic" }}>⚠️ 此題未作答</p>
                  )}
                </div>
              </div>
            );
          })}
        </div>

        {/* 底部 */}
        <div style={{ padding:"12px 20px", borderTop:`1px solid ${C.border}`, display:"flex", justifyContent:"flex-end", gap:8 }}>
          <Btn onClick={onClose} variant="outline">關閉</Btn>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════
   後台問答管理（課程小幫手回覆 + 分享）
   ══════════════════════════════════════ */
function QuestionAdmin({ questions, courses }) {
  const [filterCourse, setFilterCourse] = useState("");
  const [filterStatus, setFilterStatus] = useState("all");  // all / pending / answered
  const [answerText, setAnswerText] = useState({});  // {questionId: "回覆內容"}
  const [savingId, setSavingId] = useState(null);

  const filtered = questions
    .filter(q => !filterCourse || q.courseId === filterCourse)
    .filter(q => filterStatus === "all" || q.status === filterStatus)
    .sort((a,b) => {
      // 待回覆優先，其次依時間新到舊
      if (a.status !== b.status) return a.status === "pending" ? -1 : 1;
      const ta = a.createdAt?.toMillis ? a.createdAt.toMillis() : 0;
      const tb = b.createdAt?.toMillis ? b.createdAt.toMillis() : 0;
      return tb - ta;
    });

  const submitAnswer = async (q) => {
    const text = (answerText[q.id] || "").trim();
    if (!text) { alert("請輸入回覆內容"); return; }
    setSavingId(q.id);
    try {
      await answerQuestion(q.id, text);
      setAnswerText(prev => { const n = {...prev}; delete n[q.id]; return n; });
    } catch (e) {
      alert("回覆失敗：" + e.message);
    }
    setSavingId(null);
  };

  const handleShare = async (q, shared) => {
    try {
      await toggleQuestionShared(q.id, shared);
    } catch (e) {
      alert("操作失敗：" + e.message);
    }
  };

  const handleDelete = async (q) => {
    if (!confirm(`確定刪除這則提問？\n\n主旨：${q.subject}\n\n⚠️ 此操作無法復原`)) return;
    try {
      await deleteQuestion(q.id);
    } catch (e) {
      alert("刪除失敗：" + e.message);
    }
  };

  const pendingCount = questions.filter(q => q.status === "pending").length;

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:8 }}>
        <h2 style={{ fontSize:18, fontWeight:700, color:C.text, margin:0 }}>🙋 問答管理</h2>
        {pendingCount > 0 && <span style={{ fontSize:12, padding:"4px 12px", borderRadius:14, background:`${C.danger}12`, color:C.danger, fontWeight:600 }}>{pendingCount} 則待回覆</span>}
      </div>

      <div style={{ background:"#FFF", borderRadius:9, padding:14, border:`1px solid ${C.border}`, marginBottom:14, display:"flex", gap:10, flexWrap:"wrap" }}>
        <div style={{ flex:1, minWidth:160 }}>
          <label style={{ display:"block", fontSize:11, color:C.textLight, marginBottom:4 }}>篩選課程</label>
          <select value={filterCourse} onChange={e => setFilterCourse(e.target.value)} style={inp}>
            <option value="">全部課程</option>
            {courses.map(c => <option key={c.id} value={c.id}>{c.title}</option>)}
          </select>
        </div>
        <div style={{ flex:1, minWidth:160 }}>
          <label style={{ display:"block", fontSize:11, color:C.textLight, marginBottom:4 }}>狀態</label>
          <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} style={inp}>
            <option value="all">全部</option>
            <option value="pending">待回覆</option>
            <option value="answered">已回覆</option>
          </select>
        </div>
      </div>

      {filtered.length === 0 ? (
        <div style={{ textAlign:"center", padding:48, color:C.textLight, background:"#FFF", borderRadius:12, border:`1px solid ${C.border}` }}>
          <p style={{ fontSize:40, margin:0 }}>📭</p>
          <p style={{ fontSize:14, margin:"10px 0 0" }}>目前沒有符合條件的提問</p>
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
          {filtered.map(q => {
            const course = courses.find(c => c.id === q.courseId);
            const isPending = q.status === "pending";
            return (
              <div key={q.id} style={{ background:"#FFF", borderRadius:12, border:`1px solid ${isPending ? C.warning+"60" : C.border}`, overflow:"hidden" }}>
                {/* 標頭 */}
                <div style={{ padding:"12px 16px", background: isPending ? `${C.warning}08` : C.bgSoft, borderBottom:`1px solid ${C.border}`, display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                    <span style={{ fontSize:11, padding:"3px 9px", borderRadius:10, background: isPending ? `${C.warning}20` : `${C.success}15`, color: isPending ? C.warning : C.success, fontWeight:600 }}>
                      {isPending ? "● 待回覆" : "✓ 已回覆"}
                    </span>
                    <span style={{ fontSize:12, color:C.textMid }}>{course?.title || q.courseName || "未知課程"}</span>
                    {q.helperEmail && <span style={{ fontSize:11, color:C.textLight }}>· 小幫手：{q.helperEmail}</span>}
                  </div>
                  <span style={{ fontSize:11, color:C.textLight }}>{q.createdAt?.toDate ? q.createdAt.toDate().toLocaleString("zh-TW") : ""}</span>
                </div>

                <div style={{ padding:"14px 16px" }}>
                  {/* 提問者 + 內容 */}
                  <div style={{ marginBottom:12 }}>
                    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
                      <span style={{ width:28, height:28, borderRadius:"50%", background:`${C.navy}12`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:13, flexShrink:0 }}>🙋</span>
                      <div>
                        <span style={{ fontSize:13, fontWeight:600, color:C.text }}>{q.userName}</span>
                        {q.userEmail && <span style={{ fontSize:11, color:C.textLight, marginLeft:6 }}>{q.userEmail}</span>}
                      </div>
                    </div>
                    <p style={{ margin:"0 0 4px", fontSize:14, fontWeight:600, color:C.text }}>{q.subject}</p>
                    <p style={{ margin:0, fontSize:13, color:C.textMid, lineHeight:1.6, whiteSpace:"pre-wrap" }}>{q.content}</p>
                  </div>

                  {/* 回覆區 */}
                  {q.status === "answered" ? (
                    <div style={{ padding:"12px 14px", background:`${C.gold}08`, borderRadius:8, border:`1px solid ${C.gold}30` }}>
                      <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:5 }}>
                        <PersonIcon size={14} color={C.navy} filled />
                        <span style={{ fontSize:12, fontWeight:600, color:C.navy }}>講師回覆</span>
                        {q.answeredAt?.toDate && <span style={{ fontSize:10, color:C.textLight }}>· {q.answeredAt.toDate().toLocaleDateString("zh-TW")}</span>}
                      </div>
                      <p style={{ margin:0, fontSize:13, color:C.text, lineHeight:1.6, whiteSpace:"pre-wrap" }}>{q.answer}</p>
                    </div>
                  ) : (
                    <div>
                      <textarea
                        value={answerText[q.id] || ""}
                        onChange={e => setAnswerText(prev => ({ ...prev, [q.id]: e.target.value }))}
                        rows={3}
                        placeholder="輸入回覆內容..."
                        style={{ width:"100%", padding:"10px 12px", borderRadius:8, border:`1px solid ${C.border}`, fontSize:13, resize:"vertical", boxSizing:"border-box", fontFamily:"inherit", outline:"none" }}
                      />
                      <div style={{ display:"flex", justifyContent:"flex-end", marginTop:8 }}>
                        <Btn onClick={() => submitAnswer(q)} disabled={savingId===q.id} variant="gold">{savingId===q.id ? "送出中..." : "送出回覆"}</Btn>
                      </div>
                    </div>
                  )}

                  {/* 操作列：分享 / 取消分享 / 刪除 */}
                  <div style={{ display:"flex", alignItems:"center", gap:8, marginTop:12, paddingTop:12, borderTop:`1px solid ${C.border}`, flexWrap:"wrap" }}>
                    {q.status === "answered" && (
                      <>
                        <button
                          onClick={() => handleShare(q, true)}
                          disabled={q.shared}
                          style={{ padding:"6px 14px", borderRadius:7, border:"none", fontSize:12, fontWeight:600, cursor: q.shared ? "default" : "pointer", background: q.shared ? `${C.textLight}30` : C.success, color: q.shared ? C.textLight : "#FFF", display:"flex", alignItems:"center", gap:5 }}
                        >
                          {q.shared ? "✓ 已分享至課後交流" : "🟢 分享至課後交流"}
                        </button>
                        <button
                          onClick={() => handleShare(q, false)}
                          disabled={!q.shared}
                          style={{ padding:"6px 14px", borderRadius:7, border:`1px solid ${!q.shared ? C.textLight+"30" : C.warning}`, fontSize:12, fontWeight:600, cursor: !q.shared ? "default" : "pointer", background:"#FFF", color: !q.shared ? C.textLight : C.warning }}
                        >
                          取消分享
                        </button>
                      </>
                    )}
                    <button onClick={() => handleDelete(q)} style={{ marginLeft:"auto", padding:"6px 12px", borderRadius:7, border:`1px solid ${C.danger}40`, background:"#FFF", color:C.danger, fontSize:12, cursor:"pointer", fontWeight:500 }}>🗑️ 刪除</button>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
