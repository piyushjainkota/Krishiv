import dotenv from "dotenv";

dotenv.config();

export const config = {
  mongodbUri: process.env.MONGODB_URI ?? "mongodb://127.0.0.1:27017/krishiv_seed",
  port: Number(process.env.PORT ?? 4000),
  enableDiscrepancyWorkflow: process.env.ENABLE_DISCREPANCY_WORKFLOW !== "false",
  importAdminPassword: process.env.IMPORT_ADMIN_PASSWORD ?? "krishiv-import-lock"
};
