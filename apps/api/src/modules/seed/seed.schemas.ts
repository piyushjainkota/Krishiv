import { z } from "zod";

export const registrationSchema = z.object({
  id: z.string(),
  season: z.string(),
  year: z.string(),
  ro: z.string(),
  cropRegistrationCode: z.string(),
  farmerName: z.string(),
  fatherName: z.string(),
  village: z.string(),
  block: z.string(),
  district: z.string(),
  crop: z.string(),
  variety: z.string(),
  classStage: z.string(),
  registeredAreaHa: z.number(),
  inspectedAreaHa: z.number(),
  rejectedAreaHa: z.number(),
  certifiedAreaHa: z.number(),
  expectedYieldQtl: z.number(),
  allowedIntakeQtl: z.number(),
  totalReceivedQtl: z.number(),
  balanceQtl: z.number(),
  status: z.enum(["ACTIVE", "BLOCKED", "CLOSED", "EXHAUSTED"]),
  sourceRowNumber: z.number()
});

export const importRegistrationsSchema = z.object({
  fileName: z.string(),
  seasonLabel: z.string().optional(),
  importPassword: z.string().min(1),
  registrations: z.array(registrationSchema)
});

export const createGodownSchema = z.object({
  name: z.string().min(1)
});

export const createStackSchema = z.object({
  godownId: z.string().min(1),
  stackNo: z.string().min(1)
});

export const receiptLineSchema = z.object({
  id: z.string(),
  godownId: z.string(),
  stackNo: z.string().min(1),
  grossWeightQtl: z.number().positive(),
  noOfBags: z.number().nonnegative(),
  weightPerBagKg: z.number().nonnegative(),
  netWeightQtl: z.number().nonnegative(),
  moisturePercent: z.number().nonnegative(),
  vehicleNo: z.string().min(1),
  remarks: z.string().optional().default("")
});

export const createReceiptSchema = z.object({
  receiptNo: z.string().optional(),
  receiptDate: z.string().min(1),
  cropRegistrationId: z.string().min(1),
  lines: z.array(receiptLineSchema).min(1)
});

export const updateReceiptSchema = createReceiptSchema.extend({
  receiptNo: z.string().min(1)
});

export const createDiscrepancyShiftSchema = z.object({
  discrepancyId: z.string().min(1),
  toGodownId: z.string().min(1),
  toStackNo: z.string().min(1),
  shiftedQtyQtl: z.number().positive(),
  shiftedBags: z.number().int().positive(),
  shiftDate: z.string().min(1),
  approvedBy: z.string().optional().default(""),
  remarks: z.string().optional().default("")
});

export const createFinancialVoucherSchema = z.object({
  voucherDate: z.string().min(1),
  cropRegistrationId: z.string().min(1),
  certifiedRatePerQtl: z.number().nonnegative(),
  discrepancyRatePerQtl: z.number().nonnegative().default(0),
  deductionAmount: z.number().nonnegative().default(0),
  remarks: z.string().optional().default("")
});

export const updateFinancialVoucherSchema = createFinancialVoucherSchema;

export const loginSchema = z.object({
  email: z.string().trim().min(1),
  password: z.string().min(1)
});

export const backupDatabaseSchema = z.object({
  backupDirectory: z.string().trim().min(1)
});

export const restoreDatabaseSchema = z.object({
  restoreDirectory: z.string().trim().min(1)
});

export const reportFilterSchema = z.object({
  reportType: z.enum([
    "GODOWN_WISE_DETAIL",
    "FARMER_WISE_DETAIL",
    "SUMMARY",
    "DAILY_INTAKE_REGISTER",
    "REGISTRATION_PENDING_RECEIVED",
    "LOT_WISE_STOCK_LEDGER",
    "STACK_WISE_STOCK_POSITION",
    "DISCREPANCY_REGISTER"
  ]),
  season: z.string().optional().default(""),
  fromDate: z.string().optional().default(""),
  toDate: z.string().optional().default(""),
  crop: z.string().optional().default(""),
  variety: z.string().optional().default(""),
  classStage: z.string().optional().default(""),
  district: z.string().optional().default(""),
  godownId: z.string().optional().default(""),
  stackNo: z.string().optional().default(""),
  cropRegistrationCode: z.string().optional().default(""),
  farmerName: z.string().optional().default(""),
  reportMode: z
    .enum(["ALL", "ACCEPTED_ONLY", "DISCREPANCY_ONLY"])
    .optional()
    .default("ALL"),
  includeDiscrepancy: z.boolean().optional().default(true)
});

export type ImportRegistrationsInput = z.infer<typeof importRegistrationsSchema>;
export type CreateGodownInput = z.infer<typeof createGodownSchema>;
export type CreateStackInput = z.infer<typeof createStackSchema>;
export type CreateReceiptInput = z.infer<typeof createReceiptSchema>;
export type UpdateReceiptInput = z.infer<typeof updateReceiptSchema>;
export type CreateDiscrepancyShiftInput = z.infer<typeof createDiscrepancyShiftSchema>;
export type CreateFinancialVoucherInput = z.infer<typeof createFinancialVoucherSchema>;
export type UpdateFinancialVoucherInput = z.infer<typeof updateFinancialVoucherSchema>;
export type LoginInput = z.infer<typeof loginSchema>;
export type BackupDatabaseInput = z.infer<typeof backupDatabaseSchema>;
export type RestoreDatabaseInput = z.infer<typeof restoreDatabaseSchema>;
export type ReportFilterInput = z.infer<typeof reportFilterSchema>;
