import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import {
  addOrganizerPaymentSchema,
  addFinancialVoucherPaymentSchema,
  assignRegistrationOrganizerSchema,
  backupDatabaseSchema,
  createDiscrepancyShiftSchema,
  createStackAccommodationSchema,
  createFinancialVoucherSchema,
  createGodownSchema,
  createOrganizerSchema,
  createReceiptSchema,
  createStackSchema,
  financialVoucherActionSchema,
  importRegistrationsSchema,
  loginSchema,
  reportFilterSchema,
  restoreDatabaseSchema,
  stackAccommodationActionSchema,
  updateOrganizerPaymentSchema,
  updateOrganizerSchema,
  updateFinancialVoucherPaymentSchema,
  updateFinancialVoucherSchema,
  updateStackAccommodationSchema,
  updateReceiptSchema
} from "./seed.schemas";
import { SeedService } from "./seed.service";

const seedService = new SeedService();

type AppRole = "ADMIN" | "MANAGER" | "USER";

async function requireRoles(
  request: FastifyRequest,
  reply: FastifyReply,
  allowedRoles: AppRole[]
) {
  const email = String(request.headers["x-user-email"] ?? "");
  const role = String(request.headers["x-user-role"] ?? "");
  try {
    return await seedService.authorizeUser(email, role, allowedRoles);
  } catch (error) {
    return reply.status(403).send({
      message: error instanceof Error ? error.message : "Access denied."
    });
  }
}

export async function registerSeedRoutes(app: FastifyInstance) {
  app.post("/api/auth/login", async (request, reply) => {
    const parsed = loginSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    try {
      return await seedService.loginUser(parsed.data);
    } catch (error) {
      return reply.status(401).send({
        message: error instanceof Error ? error.message : "Invalid login."
      });
    }
  });

  app.get("/api/seed/bootstrap", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER", "USER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    return seedService.bootstrap();
  });

  app.get("/api/seed/bootstrap/core", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER", "USER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    return seedService.bootstrapCore();
  });

  app.get("/api/seed/bootstrap/operations", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER", "USER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    return seedService.bootstrapOperational();
  });

  app.post("/api/seed/import/registrations", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = importRegistrationsSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.importRegistrations(parsed.data);
  });

  app.post("/api/seed/masters/godowns", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = createGodownSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.createGodown(parsed.data);
  });

  app.post("/api/seed/masters/stacks", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = createStackSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.createStack(parsed.data);
  });

  app.post("/api/seed/masters/organizers", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = createOrganizerSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.createOrganizer(parsed.data);
  });

  app.put("/api/seed/masters/organizers/:organizerId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = updateOrganizerSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const organizerId = String((request.params as { organizerId: string }).organizerId);
    return seedService.updateOrganizer(organizerId, parsed.data);
  });

  app.delete("/api/seed/masters/organizers/:organizerId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const organizerId = String((request.params as { organizerId: string }).organizerId);
    return seedService.deleteOrganizer(organizerId);
  });

  app.put("/api/seed/registrations/:registrationId/organizer", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = assignRegistrationOrganizerSchema.safeParse(request.body ?? {});
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const registrationId = String((request.params as { registrationId: string }).registrationId);
    return seedService.assignOrganizerToRegistration(registrationId, parsed.data);
  });

  app.post("/api/seed/intake/receipts", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER", "USER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = createReceiptSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.saveReceipt(parsed.data);
  });

  app.put("/api/seed/intake/receipts/:receiptNo", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = updateReceiptSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const receiptNo = String((request.params as { receiptNo: string }).receiptNo);
    return seedService.updateReceipt(receiptNo, parsed.data);
  });

  app.delete("/api/seed/intake/receipts/:receiptNo", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const receiptNo = String((request.params as { receiptNo: string }).receiptNo);
    return seedService.deleteReceipt(receiptNo);
  });

  app.post("/api/seed/discrepancies/shifts", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = createDiscrepancyShiftSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.createDiscrepancyShift(parsed.data);
  });

  app.post("/api/seed/discrepancies/accommodations", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = createStackAccommodationSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const actor = denied as {
      user: { email: string; role: AppRole };
      permissions: { canShift: boolean };
    };
    if (!actor.permissions.canShift) {
      return reply.status(403).send({ message: "Your role cannot save stack accommodation." });
    }
    return seedService.createStackAccommodation(parsed.data, actor.user);
  });

  app.put("/api/seed/discrepancies/accommodations/:accommodationId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = updateStackAccommodationSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const actor = denied as {
      user: { email: string; role: AppRole };
      permissions: { canShift: boolean };
    };
    if (!actor.permissions.canShift) {
      return reply.status(403).send({ message: "Your role cannot edit stack accommodation." });
    }
    const accommodationId = String((request.params as { accommodationId: string }).accommodationId);
    return seedService.updateStackAccommodation(accommodationId, parsed.data, actor.user);
  });

  app.delete("/api/seed/discrepancies/accommodations/:accommodationId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = stackAccommodationActionSchema.safeParse(request.body ?? {});
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const actor = denied as {
      user: { email: string; role: AppRole };
      permissions: { canShift: boolean };
    };
    if (!actor.permissions.canShift) {
      return reply.status(403).send({ message: "Your role cannot delete stack accommodation." });
    }
    const accommodationId = String((request.params as { accommodationId: string }).accommodationId);
    return seedService.deleteStackAccommodation(accommodationId, actor.user);
  });

  app.post("/api/seed/financial-vouchers", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = createFinancialVoucherSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.createFinancialVoucher(parsed.data);
  });

  app.put("/api/seed/financial-vouchers/:voucherId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = updateFinancialVoucherSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const voucherId = String((request.params as { voucherId: string }).voucherId);
    return seedService.updateFinancialVoucher(voucherId, parsed.data);
  });

  app.post("/api/seed/financial-vouchers/:voucherId/paid", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = financialVoucherActionSchema.safeParse(request.body ?? {});
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const voucherId = String((request.params as { voucherId: string }).voucherId);
    return seedService.markFinancialVoucherPaid(voucherId);
  });

  app.delete("/api/seed/financial-vouchers/:voucherId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = financialVoucherActionSchema.safeParse(request.body ?? {});
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const voucherId = String((request.params as { voucherId: string }).voucherId);
    return seedService.deleteFinancialVoucher(voucherId, parsed.data);
  });

  app.post("/api/seed/financial-vouchers/:voucherId/payments", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = addFinancialVoucherPaymentSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const voucherId = String((request.params as { voucherId: string }).voucherId);
    const actor = denied as {
      user: { email: string; role: AppRole };
    };
    return seedService.addFinancialVoucherPayment(voucherId, parsed.data, actor.user);
  });

  app.put("/api/seed/financial-vouchers/:voucherId/payments/:paymentId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = updateFinancialVoucherPaymentSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const params = request.params as { voucherId: string; paymentId: string };
    const actor = denied as {
      user: { email: string; role: AppRole };
    };
    return seedService.updateFinancialVoucherPayment(params.voucherId, params.paymentId, parsed.data, actor.user);
  });

  app.post("/api/seed/organizer-payments", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = addOrganizerPaymentSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const actor = denied as {
      user: { email: string; role: AppRole };
    };
    return seedService.addOrganizerPayment(parsed.data, actor.user);
  });

  app.put("/api/seed/organizer-payments/:paymentId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = updateOrganizerPaymentSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    const actor = denied as {
      user: { email: string; role: AppRole };
    };
    const paymentId = String((request.params as { paymentId: string }).paymentId);
    return seedService.updateOrganizerPayment(paymentId, parsed.data, actor.user);
  });

  app.delete("/api/seed/organizer-payments/:paymentId", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const actor = denied as {
      user: { email: string; role: AppRole };
    };
    const paymentId = String((request.params as { paymentId: string }).paymentId);
    return seedService.deleteOrganizerPayment(paymentId, actor.user);
  });

  app.post("/api/seed/reports/preview", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER", "USER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = reportFilterSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.previewReport(parsed.data);
  });

  app.post("/api/seed/reports/export", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN", "MANAGER", "USER"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = reportFilterSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }

    const exportFile = await seedService.exportReports(parsed.data);
    reply.header(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    reply.header("Content-Disposition", `attachment; filename=\"${exportFile.fileName}\"`);
    return reply.send(exportFile.content);
  });

  app.post("/api/seed/validations/discrepancies", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    return seedService.validateAllDiscrepancies();
  });

  app.post("/api/seed/validations/lots", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    return seedService.validateLots();
  });

  app.post("/api/seed/validations/lots/reindex", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    return seedService.reindexLots();
  });

  app.post("/api/seed/database/backup", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = backupDatabaseSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.backupDatabase(parsed.data);
  });

  app.post("/api/seed/database/restore", async (request, reply) => {
    const denied = await requireRoles(request, reply, ["ADMIN"]);
    if ((denied as { statusCode?: number } | undefined)?.statusCode) {
      return denied;
    }
    const parsed = restoreDatabaseSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send(parsed.error.flatten());
    }
    return seedService.restoreDatabase(parsed.data);
  });
}
