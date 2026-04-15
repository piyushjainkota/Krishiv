import { buildApp } from "./server";
import { config } from "./config";

const port = config.port;

async function start() {
  const app = await buildApp();
  await app.listen({ port, host: "0.0.0.0" });
}

start().catch((error) => {
  console.error(error);
  process.exit(1);
});
