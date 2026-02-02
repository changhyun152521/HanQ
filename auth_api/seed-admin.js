/**
 * 관리자 계정 한 번만 생성 (이미 있으면 건너뜀)
 * 실행: node seed-admin.js (auth_api 폴더에서) 또는 node auth_api/seed-admin.js (프로젝트 루트에서)
 */
const path = require("path");
require("dotenv").config({ path: path.join(__dirname, ".env") });
const { MongoClient } = require("mongodb");

const MONGODB_URI = process.env.MONGODB_URI || "";
const DB_NAME = "hanq";
const COLLECTION_NAME = "users";

const ADMIN_USER = {
  user_id: "admin",
  password: "admin",
  name: "관리자",
  created_at: new Date().toISOString(),
};

async function main() {
  if (!MONGODB_URI) {
    console.error("MONGODB_URI가 없습니다. .env 파일을 확인하세요.");
    process.exit(1);
  }
  const client = new MongoClient(MONGODB_URI);
  try {
    await client.connect();
    const col = client.db(DB_NAME).collection(COLLECTION_NAME);
    const existing = await col.findOne({ user_id: "admin" });
    if (existing) {
      console.log("관리자 계정이 이미 있습니다. (user_id: admin)");
      return;
    }
    await col.insertOne(ADMIN_USER);
    console.log("관리자 계정이 생성되었습니다.");
    console.log("  아이디: admin");
    console.log("  비밀번호: admin");
    console.log("  (로그인 후 정보수정에서 비밀번호를 변경하세요.)");
  } catch (err) {
    console.error("오류:", err.message);
    process.exit(1);
  } finally {
    await client.close();
  }
}

main();
