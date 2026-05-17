import dotenv from "dotenv";
import { fileURLToPath } from "node:url";

dotenv.config();
dotenv.config({
  path: fileURLToPath(new URL("../.env", import.meta.url)),
  override: false
});

export const config = {
  mongodbUri: process.env.MONGODB_URI ?? "mongodb://127.0.0.1:27017/krishiv_seed",
  port: Number(process.env.PORT ?? 4000),
  enableDiscrepancyWorkflow: process.env.ENABLE_DISCREPANCY_WORKFLOW !== "false",
  importAdminPassword: process.env.IMPORT_ADMIN_PASSWORD ?? "krishiv-import-lock",
  jwtSecret:
    process.env.JWT_SECRET ??
    "krishiv-local-development-jwt-secret-change-in-production",
  frontendOrigins: (process.env.FRONTEND_ORIGINS ?? "http://localhost:3000,http://127.0.0.1:3000,https://krishivseed.vercel.app")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean),
  requestBodyLimitBytes: Number(process.env.REQUEST_BODY_LIMIT_BYTES ?? 1_048_576),
  loginRateLimitMax: Number(process.env.LOGIN_RATE_LIMIT_MAX ?? 30),
  backupRoot: process.env.BACKUP_ROOT ?? "mongo-backups"
};
