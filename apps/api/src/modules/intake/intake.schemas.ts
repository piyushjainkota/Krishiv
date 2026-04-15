import { z } from "zod";

export const intakeLineSchema = z.object({
  cropRegistrationId: z.string().uuid(),
  farmerId: z.string().uuid(),
  crop: z.string().min(1),
  variety: z.string().min(1),
  classStage: z.string().min(1),
  godownId: z.string().uuid(),
  stackId: z.string().uuid(),
  qtyQtl: z.number().positive(),
  noOfBags: z.number().positive(),
  weightPerBagKg: z.number().positive().optional(),
  moisturePercent: z.number().min(0).max(100),
  qualityStatus: z.enum(["ACCEPTED", "HOLD", "REJECTED"]),
  remarks: z.string().max(500).optional()
});

export const createIntakeReceiptSchema = z.object({
  receiptNo: z.string().min(1),
  receiptDatetime: z.string().datetime(),
  vehicleNo: z.string().min(1),
  remarks: z.string().max(500).optional(),
  lines: z.array(intakeLineSchema).min(1)
});

export type CreateIntakeReceiptInput = z.infer<typeof createIntakeReceiptSchema>;
