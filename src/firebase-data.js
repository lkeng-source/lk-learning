// Firebase 資料層 - 所有資料庫操作集中在此
import {
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged,
  updatePassword,
  EmailAuthProvider,
  reauthenticateWithCredential
} from "firebase/auth";
import {
  collection,
  doc,
  setDoc,
  getDoc,
  getDocs,
  addDoc,
  updateDoc,
  deleteDoc,
  query,
  where,
  orderBy,
  onSnapshot,
  serverTimestamp,
  writeBatch
} from "firebase/firestore";
import { auth, db } from "./firebase";

// ───────────────────────────────────────────
// 預設資料（首次使用時建立）
// ───────────────────────────────────────────
export const DEFAULT_CATEGORIES = [
  { id: "sys1", name: "生產系統", icon: "⚙️", color: "#0066CC", order: 1 },
  { id: "sys2", name: "品質系統", icon: "✅", color: "#00875A", order: 2 },
  { id: "sys3", name: "安全系統", icon: "🛡️", color: "#E65100", order: 3 },
  { id: "sys4", name: "環境系統", icon: "🌿", color: "#2E7D32", order: 4 },
  { id: "sys5", name: "資訊系統", icon: "💻", color: "#5B21B6", order: 5 },
  { id: "mgmt", name: "管理類", icon: "📊", color: "#B45309", order: 6 },
];

export const DEFAULT_COURSES = [
  { id:"c1", title:"5S 管理實務", category:"sys1", instructor:"王大明", duration:45, thumbnail:"📘", description:"學習5S管理方法，提升生產效率與工作環境品質。", views:0, publishDate:"2026-05-10", status:"published", files:[], chapters:[{title:"5S 概論",duration:15,youtubeUrl:""},{title:"整理與整頓",duration:15,youtubeUrl:""},{title:"清掃、清潔、素養",duration:15,youtubeUrl:""}], quiz:[{q:"5S 中的「整理」指的是什麼？",options:["區分需要與不需要的物品","將物品擺放整齊","打掃環境","維持好習慣"],answer:0},{q:"5S 實施的正確順序為何？",options:["整理→整頓→清掃→清潔→素養","清掃→整理→整頓→清潔→素養","素養→整理→整頓→清掃→清潔","整頓→整理→清掃→清潔→素養"],answer:0}] },
  { id:"c2", title:"ISO 9001 品質管理入門", category:"sys2", instructor:"李美玲", duration:60, thumbnail:"📗", description:"深入了解 ISO 9001 品質管理系統的核心要求與實施方法。", views:0, publishDate:"2026-05-08", status:"published", files:[], chapters:[{title:"ISO 9001 概述",duration:20,youtubeUrl:""},{title:"品質管理原則",duration:20,youtubeUrl:""},{title:"文件化要求",duration:20,youtubeUrl:""}], quiz:[{q:"ISO 9001 最新版本是哪一年發布？",options:["2008","2012","2015","2020"],answer:2}] },
  { id:"c3", title:"職業安全衛生基礎", category:"sys3", instructor:"張志強", duration:50, thumbnail:"📕", description:"了解職場安全衛生法規與防災基本知識。", views:0, publishDate:"2026-05-01", status:"published", files:[], chapters:[{title:"安全衛生法規",duration:20,youtubeUrl:""},{title:"危害辨識",duration:15,youtubeUrl:""},{title:"事故預防",duration:15,youtubeUrl:""}], quiz:[{q:"以下何者屬於物理性危害？",options:["噪音","化學溶劑","病毒","心理壓力"],answer:0}] },
];

// ───────────────────────────────────────────
// 認證相關
// ───────────────────────────────────────────
export const loginWithEmail = async (email, password) => {
  const result = await signInWithEmailAndPassword(auth, email, password);
  return result.user;
};

export const logoutUser = async () => {
  await signOut(auth);
};

export const watchAuthState = (callback) => {
  return onAuthStateChanged(auth, callback);
};

export const changePassword = async (oldPassword, newPassword) => {
  const user = auth.currentUser;
  if (!user) throw new Error("未登入");
  const credential = EmailAuthProvider.credential(user.email, oldPassword);
  await reauthenticateWithCredential(user, credential);
  await updatePassword(user, newPassword);
  // 同步更新 Firestore 的 mustChangePw
  await updateDoc(doc(db, "users", user.uid), { mustChangePw: false });
};

// ───────────────────────────────────────────
// 使用者資料（users collection）
// 文件 ID = Firebase Auth UID
// ───────────────────────────────────────────
export const getCurrentUserData = async (uid) => {
  const snap = await getDoc(doc(db, "users", uid));
  return snap.exists() ? { id: snap.id, ...snap.data() } : null;
};

export const getAllUsers = async () => {
  const snap = await getDocs(collection(db, "users"));
  return snap.docs.map(d => ({ id: d.id, ...d.data() }));
};

export const watchAllUsers = (callback) => {
  return onSnapshot(collection(db, "users"), (snap) => {
    const users = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    callback(users);
  });
};

// 注意：建立使用者帳號需要 Admin SDK，前端無法直接做
// 這裡的「新增使用者」只在 Firestore 建紀錄，實際登入帳號要管理員在 Firebase Console 建立
// 或者用 Cloud Functions（需付費）
// 暫時的折衷方案：用 createUserWithEmailAndPassword，但會自動登入新使用者
// 所以我們用一個變通方法：建立後請管理員重新登入
export const createUserAccount = async (empNo, name, email, password, role, department) => {
  // 用 Auth 建立帳號（會自動登入這個新帳號，需要再讓 admin 重新登入）
  const result = await createUserWithEmailAndPassword(auth, email, password || empNo);
  const uid = result.user.uid;
  // 在 Firestore 建立使用者資料
  await setDoc(doc(db, "users", uid), {
    empNo,
    name,
    email,
    role: role || "user",
    department: department || "",
    mustChangePw: !password, // 沒給密碼就是用預設員工編號當密碼，需要改
    createdAt: serverTimestamp(),
  });
  return uid;
};

export const updateUserData = async (uid, data) => {
  await updateDoc(doc(db, "users", uid), data);
};

export const deleteUserData = async (uid) => {
  // 注意：這只刪 Firestore 資料，Auth 帳號要在 Firebase Console 刪
  await deleteDoc(doc(db, "users", uid));
};

// ───────────────────────────────────────────
// 課程資料（courses collection）
// ───────────────────────────────────────────
export const watchCourses = (callback) => {
  return onSnapshot(collection(db, "courses"), (snap) => {
    const courses = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    callback(courses);
  });
};

export const addCourse = async (courseData) => {
  const ref = await addDoc(collection(db, "courses"), {
    ...courseData,
    createdAt: serverTimestamp(),
  });
  return ref.id;
};

export const updateCourse = async (id, data) => {
  await updateDoc(doc(db, "courses", id), data);
};

export const deleteCourse = async (id) => {
  await deleteDoc(doc(db, "courses", id));
};

export const incrementViews = async (id, currentViews) => {
  await updateDoc(doc(db, "courses", id), { views: (currentViews || 0) + 1 });
};

// ───────────────────────────────────────────
// 分類資料（categories collection）
// ───────────────────────────────────────────
export const watchCategories = (callback) => {
  return onSnapshot(collection(db, "categories"), (snap) => {
    const cats = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    callback(cats);
  });
};

export const addCategory = async (categoryData) => {
  const ref = await addDoc(collection(db, "categories"), categoryData);
  return ref.id;
};

export const updateCategory = async (id, data) => {
  await updateDoc(doc(db, "categories", id), data);
};

export const deleteCategory = async (id) => {
  await deleteDoc(doc(db, "categories", id));
};

// ───────────────────────────────────────────
// 學習紀錄（watchHistory collection）
// 文件 ID = {userId}_{courseId}
// ───────────────────────────────────────────
export const watchUserHistory = (userId, callback) => {
  const q = query(collection(db, "watchHistory"), where("userId", "==", userId));
  return onSnapshot(q, (snap) => {
    const history = {};
    snap.docs.forEach(d => {
      history[d.id] = { id: d.id, ...d.data() };
    });
    callback(history);
  });
};

// 管理員：訂閱所有人的觀看紀錄
export const watchAllWatchHistory = (callback) => {
  return onSnapshot(collection(db, "watchHistory"), (snap) => {
    const records = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    callback(records);
  });
};

export const recordWatchProgress = async (userId, courseId, chapterIndex, progress) => {
  const docId = `${userId}_${courseId}`;
  const docRef = doc(db, "watchHistory", docId);
  const snap = await getDoc(docRef);
  const currentData = snap.exists() ? snap.data() : { totalTime: 0 };
  await setDoc(docRef, {
    userId,
    courseId,
    chapterIndex,
    progress,
    lastWatched: serverTimestamp(),
    totalTime: (currentData.totalTime || 0) + 1,
  }, { merge: true });
};

// ───────────────────────────────────────────
// 測驗結果（quizResults collection）
// ───────────────────────────────────────────
export const watchAllQuizResults = (callback) => {
  return onSnapshot(collection(db, "quizResults"), (snap) => {
    const results = {};
    snap.docs.forEach(d => {
      results[d.id] = { id: d.id, ...d.data() };
    });
    callback(results);
  });
};

export const saveQuizResult = async (userId, courseId, score, total, userName, answers, questionsSnapshot) => {
  const docId = `${userId}_${courseId}`;
  await setDoc(doc(db, "quizResults", docId), {
    userId,
    courseId,
    score,
    total,
    userName,
    answers: answers || {},  // 學員的作答：{ "0": 2, "1": 0, ... } 題號→選的選項
    questionsSnapshot: questionsSnapshot || [],  // 當時測驗的題目快照（避免題目後來被改動）
    date: serverTimestamp(),
  });
};

// ───────────────────────────────────────────
// 初始化預設資料（第一次使用時呼叫）
// ───────────────────────────────────────────
export const initializeDefaultData = async () => {
  // 檢查是否已經有分類
  const catsSnap = await getDocs(collection(db, "categories"));
  if (catsSnap.empty) {
    const batch = writeBatch(db);
    DEFAULT_CATEGORIES.forEach(cat => {
      const { id, ...data } = cat;
      batch.set(doc(db, "categories", id), data);
    });
    await batch.commit();
    console.log("Default categories created");
  }

  // 檢查是否已經有課程
  const coursesSnap = await getDocs(collection(db, "courses"));
  if (coursesSnap.empty) {
    const batch = writeBatch(db);
    DEFAULT_COURSES.forEach(course => {
      const { id, ...data } = course;
      batch.set(doc(db, "courses", id), data);
    });
    await batch.commit();
    console.log("Default courses created");
  }
};
