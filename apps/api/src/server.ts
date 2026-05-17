import cors from "@fastify/cors";
import jwt from "@fastify/jwt";
import rateLimit from "@fastify/rate-limit";
import sensible from "@fastify/sensible";
import Fastify from "fastify";
import { config } from "./config";
import { connectDatabase } from "./db";
import { registerHealthRoutes } from "./modules/health/health.routes";
import { registerIntakeRoutes } from "./modules/intake/intake.routes";
import { registerSeedRoutes } from "./modules/seed/seed.routes";

export async function buildApp() {
  await connectDatabase();

  const app = Fastify({
    logger: true,
    bodyLimit: config.requestBodyLimitBytes
  });

  await app.register(cors, {
    origin(origin, callback) {
      if (!origin || config.frontendOrigins.includes(origin)) {
        callback(null, true);
        return;
      }
      callback(new Error("Origin is not allowed."), false);
    },
    credentials: true
  });

  await app.register(jwt, {
    secret: config.jwtSecret
  });
  await app.register(rateLimit, {
    global: false
  });
  await app.register(sensible);

  app.setErrorHandler((error, _request, reply) => {
    const requestError = error as {
      validation?: unknown;
      statusCode?: number;
      message?: string;
    };
    const validationError = Boolean(requestError.validation);
    const statusCode = validationError
      ? 400
      : requestError.statusCode && requestError.statusCode < 500
        ? requestError.statusCode
        : 400;
    const rawMessage = error instanceof Error ? error.message : "Request failed.";
    const safeMessage =
      rawMessage.includes("ENOENT") || rawMessage.includes("\\") || rawMessage.includes("/")
        ? "Request failed. Please check the submitted details and try again."
        : rawMessage;

    reply.status(statusCode).send({
      message: safeMessage
    });
  });

  await registerHealthRoutes(app);
  await registerIntakeRoutes(app);
  await registerSeedRoutes(app);

  return app;
}
