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
  organizerId: z.string().optional().default(""),
  organizerName: z.string().optional().default(""),
  organizerCommissionRatePerQtl: z.number().optional().default(0),
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

export const createOrganizerSchema = z.object({
  name: z.string().trim().min(1),
  mobile: z.string().trim().optional().default(""),
  village: z.string().trim().optional().default(""),
  district: z.string().trim().optional().default(""),
  commissionRatePerQtl: z.number().nonnegative(),
  deductionAmount: z.number().nonnegative().optional().default(0),
  isActive: z.boolean().optional().default(true)
});

export const updateOrganizerSchema = createOrganizerSchema;

export const assignRegistrationOrganizerSchema = z.object({
  organizerId: z.string().trim().optional().default("")
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

export const createStackAccommodationSchema = z.object({
  discrepancyId: z.string().min(1),
  targetRegistrationId: z.string().min(1),
  adjustedQtyQtl: z.number().positive(),
  adjustedBags: z.number().int().nonnegative().optional().default(0),
  adjustmentDate: z.string().min(1),
  remarks: z.string().optional().default("")
});

export const updateStackAccommodationSchema = createStackAccommodationSchema.extend({
  adminPassword: z.string().optional().default("")
});

export const stackAccommodationActionSchema = z.object({
  adminPassword: z.string().optional().default("")
});

export const createFinancialVoucherSchema = z.object({
  voucherDate: z.string().min(1),
  cropRegistrationId: z.string().min(1),
  certifiedRatePerQtl: z.number().nonnegative(),
  discrepancyRatePerQtl: z.number().nonnegative().default(0),
  deductionAmount: z.number().nonnegative().default(0),
  remarks: z.string().optional().default("")
});

export const updateFinancialVoucherSchema = createFinancialVoucherSchema.extend({
  adminPassword: z.string().optional().default("")
});

export const financialVoucherActionSchema = z.object({
  adminPassword: z.string().optional().default("")
});

export const addFinancialVoucherPaymentSchema = z.object({
  paymentDate: z.string().min(1),
  amount: z.number().nonnegative(),
  transactionNo: z.string().trim().min(1),
  remarks: z.string().optional().default("")
});

export const updateFinancialVoucherPaymentSchema = addFinancialVoucherPaymentSchema.extend({
  adminPassword: z.string().optional().default("")
});

export const addOrganizerPaymentSchema = z.object({
  organizerId: z.string().trim().min(1),
  paymentDate: z.string().min(1),
  amount: z.number().positive(),
  transactionNo: z.string().trim().min(1),
  remarks: z.string().optional().default("")
});

export const updateOrganizerPaymentSchema = z.object({
  paymentDate: z.string().min(1),
  amount: z.number().positive(),
  transactionNo: z.string().trim().min(1),
  remarks: z.string().optional().default("")
});

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
    "DISTRICT_WISE_DETAIL",
    "FARMER_WISE_DETAIL",
    "OVERALL_INTAKE",
    "SUMMARY",
    "DAILY_INTAKE_REGISTER",
    "REGISTRATION_PENDING_RECEIVED",
    "LOT_WISE_STOCK_LEDGER",
    "STACK_WISE_STOCK_POSITION",
    "STACK_CARD_REGISTER",
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
export type CreateOrganizerInput = z.infer<typeof createOrganizerSchema>;
export type UpdateOrganizerInput = z.infer<typeof updateOrganizerSchema>;
export type AssignRegistrationOrganizerInput = z.infer<typeof assignRegistrationOrganizerSchema>;
export type CreateReceiptInput = z.infer<typeof createReceiptSchema>;
export type UpdateReceiptInput = z.infer<typeof updateReceiptSchema>;
export type CreateDiscrepancyShiftInput = z.infer<typeof createDiscrepancyShiftSchema>;
export type CreateStackAccommodationInput = z.infer<typeof createStackAccommodationSchema>;
export type UpdateStackAccommodationInput = z.infer<typeof updateStackAccommodationSchema>;
export type StackAccommodationActionInput = z.infer<typeof stackAccommodationActionSchema>;
export type CreateFinancialVoucherInput = z.infer<typeof createFinancialVoucherSchema>;
export type UpdateFinancialVoucherInput = z.infer<typeof updateFinancialVoucherSchema>;
export type FinancialVoucherActionInput = z.infer<typeof financialVoucherActionSchema>;
export type AddFinancialVoucherPaymentInput = z.infer<typeof addFinancialVoucherPaymentSchema>;
export type UpdateFinancialVoucherPaymentInput = z.infer<typeof updateFinancialVoucherPaymentSchema>;
export type AddOrganizerPaymentInput = z.infer<typeof addOrganizerPaymentSchema>;
export type UpdateOrganizerPaymentInput = z.infer<typeof updateOrganizerPaymentSchema>;
export type LoginInput = z.infer<typeof loginSchema>;
export type BackupDatabaseInput = z.infer<typeof backupDatabaseSchema>;
export type RestoreDatabaseInput = z.infer<typeof restoreDatabaseSchema>;
export type ReportFilterInput = z.infer<typeof reportFilterSchema>;
