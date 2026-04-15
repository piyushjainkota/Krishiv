import mongoose from "mongoose";
import { config } from "./config";

let connected = false;

export async function connectDatabase() {
  if (connected) {
    return mongoose.connection;
  }

  mongoose.set("strictQuery", true);
  await mongoose.connect(config.mongodbUri, {
    dbName: "krishiv_seed"
  });
  connected = true;
  return mongoose.connection;
}
