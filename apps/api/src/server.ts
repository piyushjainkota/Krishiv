import cors from "@fastify/cors";
import sensible from "@fastify/sensible";
import Fastify from "fastify";
import { connectDatabase } from "./db";
import { registerHealthRoutes } from "./modules/health/health.routes";
import { registerIntakeRoutes } from "./modules/intake/intake.routes";
import { registerSeedRoutes } from "./modules/seed/seed.routes";

export async function buildApp() {
  await connectDatabase();

  const app = Fastify({
    logger: true
  });

  await app.register(cors, {
    origin: true
  });

  await app.register(sensible);
  await registerHealthRoutes(app);
  await registerIntakeRoutes(app);
  await registerSeedRoutes(app);

  return app;
}
