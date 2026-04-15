import type { FastifyInstance } from "fastify";
import { createIntakeReceiptSchema } from "./intake.schemas";
import { IntakeService } from "./intake.service";

const intakeService = new IntakeService();

export async function registerIntakeRoutes(app: FastifyInstance) {
  app.post("/api/intake/preview-allocation", async (request, reply) => {
    const parsed = createIntakeReceiptSchema.safeParse(request.body);

    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }

    const result = await intakeService.previewAllocation(parsed.data);
    return {
      receiptNo: parsed.data.receiptNo,
      lines: result
    };
  });
}
