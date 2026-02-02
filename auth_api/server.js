/**
 * HanQ 로그인/회원 API (MongoDB Atlas)
 * - POST /login.php   : 로그인
 * - GET  /list_users.php : 회원 목록
 * - POST /add_user.php  : 회원 추가
 */
require("dotenv").config();

const express = require("express");
const cors = require("cors");
const { MongoClient } = require("mongodb");

const app = express();
const PORT = process.env.PORT || 5000;
const MONGODB_URI = process.env.MONGODB_URI || "";
const DB_NAME = "hanq";
const COLLECTION_NAME = "users";

app.use(cors());
app.use(express.json({ limit: "1mb" }));

let db = null;

async function getDb() {
  if (db) return db;
  if (!MONGODB_URI) throw new Error("MONGODB_URI 환경 변수가 필요합니다.");
  const client = new MongoClient(MONGODB_URI);
  await client.connect();
  db = client.db(DB_NAME);
  return db;
}

/** POST /login.php - HanQ 로그인 (기존 PHP API와 동일 응답) */
app.post("/login.php", async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  const { user_id: rawId, password } = req.body || {};
  const user_id = typeof rawId === "string" ? rawId.trim() : "";
  if (!user_id) {
    return res.json({ success: false, message: "아이디를 입력해 주세요." });
  }
  try {
    const col = (await getDb()).collection(COLLECTION_NAME);
    const user = await col.findOne({ user_id });
    if (!user) {
      return res.json({ success: false, message: "아이디 또는 비밀번호가 맞지 않습니다." });
    }
    if (user.password !== password) {
      return res.json({ success: false, message: "아이디 또는 비밀번호가 맞지 않습니다." });
    }
    return res.json({
      success: true,
      user_id: user.user_id,
      name: user.name || user.user_id,
    });
  } catch (err) {
    console.error("login err", err);
    return res.json({ success: false, message: "DB 연결 실패" });
  }
});

/** GET /list_users.php - 회원 목록 (기존 PHP API와 동일 응답) */
app.get("/list_users.php", async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  try {
    const col = (await getDb()).collection(COLLECTION_NAME);
    const list = await col
      .find({}, { projection: { user_id: 1, name: 1, created_at: 1, _id: 0 } })
      .sort({ created_at: -1 })
      .toArray();
    const users = list.map((u) => ({
      user_id: u.user_id,
      name: u.name || u.user_id,
      created_at: u.created_at || null,
    }));
    return res.json({ success: true, users });
  } catch (err) {
    console.error("list_users err", err);
    return res.json({ success: false, message: "DB 연결 실패" });
  }
});

/** POST /add_user.php - 회원 추가 (기존 PHP API와 동일 응답) */
app.post("/add_user.php", async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  const { user_id: rawId, password, name: rawName } = req.body || {};
  const user_id = typeof rawId === "string" ? rawId.trim() : "";
  const passwordStr = typeof password === "string" ? password : "";
  const name = typeof rawName === "string" ? rawName.trim() : "";
  if (!user_id) {
    return res.json({ success: false, message: "아이디를 입력해 주세요." });
  }
  try {
    const col = (await getDb()).collection(COLLECTION_NAME);
    const existing = await col.findOne({ user_id });
    if (existing) {
      return res.json({ success: false, message: "이미 사용 중인 아이디입니다." });
    }
    const created_at = new Date().toISOString();
    await col.insertOne({
      user_id,
      password: passwordStr,
      name: name || user_id,
      created_at,
    });
    return res.json({ success: true });
  } catch (err) {
    console.error("add_user err", err);
    return res.json({ success: false, message: "저장 중 오류가 발생했습니다." });
  }
});

/** POST /update_user.php - 회원 정보 수정 (이름, 아이디, 비밀번호) */
app.post("/update_user.php", async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  const body = req.body || {};
  const current_user_id = typeof body.current_user_id === "string" ? body.current_user_id.trim() : "";
  const password = typeof body.password === "string" ? body.password : "";
  const new_user_id = typeof body.new_user_id === "string" ? body.new_user_id.trim() : "";
  const new_name = typeof body.new_name === "string" ? body.new_name.trim() : "";
  const new_password = typeof body.new_password === "string" ? body.new_password : "";

  if (!current_user_id || !password) {
    return res.json({ success: false, message: "현재 아이디와 비밀번호를 입력해 주세요." });
  }

  try {
    const col = (await getDb()).collection(COLLECTION_NAME);
    const user = await col.findOne({ user_id: current_user_id });
    if (!user) {
      return res.json({ success: false, message: "아이디 또는 비밀번호가 맞지 않습니다." });
    }
    if (user.password !== password) {
      return res.json({ success: false, message: "아이디 또는 비밀번호가 맞지 않습니다." });
    }

    const updates = {};
    if (new_name !== "") updates.name = new_name;
    if (new_password !== "") updates.password = new_password;
    if (new_user_id !== "" && new_user_id !== current_user_id) {
      const taken = await col.findOne({ user_id: new_user_id });
      if (taken) {
        return res.json({ success: false, message: "이미 사용 중인 아이디입니다." });
      }
      updates.user_id = new_user_id;
    }

    if (Object.keys(updates).length === 0) {
      return res.json({ success: true, user_id: current_user_id, name: user.name || current_user_id });
    }

    await col.updateOne(
      { user_id: current_user_id },
      { $set: updates }
    );

    const final_user_id = updates.user_id !== undefined ? updates.user_id : current_user_id;
    const final_name = updates.name !== undefined ? updates.name : (user.name || current_user_id);
    return res.json({ success: true, user_id: final_user_id, name: final_name });
  } catch (err) {
    console.error("update_user err", err);
    return res.json({ success: false, message: "저장 중 오류가 발생했습니다." });
  }
});

const ADMIN_USER_ID = "admin";

/** POST /admin_update_user.php - 관리자용 회원 수정 (admin만 호출 가능) */
app.post("/admin_update_user.php", async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  const body = req.body || {};
  const admin_user_id = typeof body.admin_user_id === "string" ? body.admin_user_id.trim() : "";
  const admin_password = typeof body.admin_password === "string" ? body.admin_password : "";
  const target_user_id = typeof body.target_user_id === "string" ? body.target_user_id.trim() : "";
  const new_user_id = typeof body.new_user_id === "string" ? body.new_user_id.trim() : "";
  const new_name = typeof body.new_name === "string" ? body.new_name.trim() : "";
  const new_password = typeof body.new_password === "string" ? body.new_password : "";

  if (admin_user_id !== ADMIN_USER_ID || !admin_password || !target_user_id) {
    return res.json({ success: false, message: "권한이 없습니다." });
  }

  try {
    const col = (await getDb()).collection(COLLECTION_NAME);
    const admin = await col.findOne({ user_id: ADMIN_USER_ID });
    if (!admin || admin.password !== admin_password) {
      return res.json({ success: false, message: "관리자 비밀번호가 맞지 않습니다." });
    }
    const target = await col.findOne({ user_id: target_user_id });
    if (!target) {
      return res.json({ success: false, message: "대상 회원을 찾을 수 없습니다." });
    }

    const updates = {};
    if (new_name !== "") updates.name = new_name;
    if (new_password !== "") updates.password = new_password;
    if (new_user_id !== "" && new_user_id !== target_user_id) {
      const taken = await col.findOne({ user_id: new_user_id });
      if (taken) {
        return res.json({ success: false, message: "이미 사용 중인 아이디입니다." });
      }
      updates.user_id = new_user_id;
    }

    if (Object.keys(updates).length === 0) {
      return res.json({ success: true });
    }

    await col.updateOne({ user_id: target_user_id }, { $set: updates });
    return res.json({ success: true });
  } catch (err) {
    console.error("admin_update_user err", err);
    return res.json({ success: false, message: "저장 중 오류가 발생했습니다." });
  }
});

/** POST /delete_user.php - 관리자용 회원 삭제 (admin만 호출 가능) */
app.post("/delete_user.php", async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  const body = req.body || {};
  const admin_user_id = typeof body.admin_user_id === "string" ? body.admin_user_id.trim() : "";
  const admin_password = typeof body.admin_password === "string" ? body.admin_password : "";
  const target_user_id = typeof body.target_user_id === "string" ? body.target_user_id.trim() : "";

  if (admin_user_id !== ADMIN_USER_ID || !admin_password || !target_user_id) {
    return res.json({ success: false, message: "권한이 없습니다." });
  }
  if (target_user_id === ADMIN_USER_ID) {
    return res.json({ success: false, message: "관리자 계정은 삭제할 수 없습니다." });
  }

  try {
    const col = (await getDb()).collection(COLLECTION_NAME);
    const admin = await col.findOne({ user_id: ADMIN_USER_ID });
    if (!admin || admin.password !== admin_password) {
      return res.json({ success: false, message: "관리자 비밀번호가 맞지 않습니다." });
    }
    const result = await col.deleteOne({ user_id: target_user_id });
    if (result.deletedCount === 0) {
      return res.json({ success: false, message: "대상 회원을 찾을 수 없습니다." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("delete_user err", err);
    return res.json({ success: false, message: "삭제 중 오류가 발생했습니다." });
  }
});

/** 헬스체크 (Heroku 등에서 사용) */
app.get("/", (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.json({ ok: true, service: "ch-lms-auth-api" });
});

async function start() {
  if (!MONGODB_URI) {
    console.warn("MONGODB_URI가 없습니다. 배포 시 Heroku Config Var에 설정하세요.");
  }
  app.listen(PORT, () => {
    console.log(`Auth API listening on port ${PORT}`);
  });
}

start().catch((err) => {
  console.error("Start failed:", err);
  process.exit(1);
});
