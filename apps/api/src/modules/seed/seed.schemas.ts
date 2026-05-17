import { z } from "zod";

const MAX_MONEY_AMOUNT = 100_000_000;
const MAX_RATE_PER_QTL = 100_000;

function stripHtml(value: string) {
  return value.replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim();
}

const safeText = z.string().transform(stripHtml);
const optionalSafeText = z.string().optional().default("").transform(stripHtml);
const moneyAmount = z.number().min(0).max(MAX_MONEY_AMOUNT);
const positiveMoneyAmount = z.number().positive().max(MAX_MONEY_AMOUNT);
const ratePerQtl = z.number().min(0).max(MAX_RATE_PER_QTL);

function isValidIsoDate(value: string) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return false;
  }
  const [year, month, day] = value.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}

function isNotFutureIsoDate(value: string) {
  if (!isValidIsoDate(value)) {
    return false;
  }
  const today = new Date();
  const todayKey = [
    today.getFullYear(),
    String(today.getMonth() + 1).padStart(2, "0"),
    String(today.getDate()).padStart(2, "0")
  ].join("-");
  return value <= todayKey;
}

const isoDate = z.string().trim().refine(isValidIsoDate, "Date must be in YYYY-MM-DD format.");
const notFutureIsoDate = z
  .string()
  .trim()
  .refine(isNotFutureIsoDate, "Date must be valid and cannot be in the future.");

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
  status: z.enum(["ACTIVE", "BLOCKED", "CLOSED", "EXHAUSTED", "REJECTED"]),
  sourceRowNumber: z.number()
});

export const importRegistrationsSchema = z.object({
  fileName: z.string(),
  seasonLabel: z.string().optional(),
  importPassword: z.string().min(1),
  registrations: z.array(registrationSchema)
});

export const createGodownSchema = z.object({
  name: safeText.pipe(z.string().min(1))
});

export const createStackSchema = z.object({
  godownId: z.string().min(1),
  stackNo: safeText.pipe(z.string().min(1))
});

export const createOrganizerSchema = z.object({
  name: safeText.pipe(z.string().min(1)),
  mobile: optionalSafeText,
  village: optionalSafeText,
  district: optionalSafeText,
  commissionRatePerQtl: ratePerQtl,
  deductionAmount: moneyAmount.optional().default(0),
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
  moisturePercent: z.number().min(0).max(100),
  vehicleNo: safeText.pipe(z.string().min(1)),
  remarks: optionalSafeText
});

export const createReceiptSchema = z.object({
  receiptNo: z.string().optional(),
  receiptDate: notFutureIsoDate,
  cropRegistrationId: z.string().min(1),
  lines: z.array(receiptLineSchema).min(1)
});

export const updateReceiptSchema = createReceiptSchema.extend({
  receiptNo: z.string().min(1)
});

export const createDiscrepancyShiftSchema = z.object({
  discrepancyId: z.string().min(1),
  toGodownId: z.string().min(1),
  toStackNo: safeText.pipe(z.string().min(1)),
  shiftedQtyQtl: z.number().positive(),
  shiftedBags: z.number().int().positive(),
  shiftDate: notFutureIsoDate,
  approvedBy: optionalSafeText,
  remarks: optionalSafeText
});

export const createStackAccommodationSchema = z.object({
  discrepancyId: z.string().min(1),
  targetRegistrationId: z.string().min(1),
  adjustedQtyQtl: z.number().positive(),
  adjustedBags: z.number().int().nonnegative().optional().default(0),
  adjustmentDate: notFutureIsoDate,
  remarks: optionalSafeText
});

export const updateStackAccommodationSchema = createStackAccommodationSchema.extend({
  adminPassword: z.string().optional().default("")
});

export const stackAccommodationActionSchema = z.object({
  adminPassword: z.string().optional().default("")
});

export const createFinancialVoucherSchema = z.object({
  voucherDate: notFutureIsoDate,
  cropRegistrationId: z.string().min(1),
  certifiedRatePerQtl: ratePerQtl,
  discrepancyRatePerQtl: ratePerQtl.default(0),
  deductionAmount: moneyAmount.default(0),
  remarks: optionalSafeText
});

export const updateFinancialVoucherSchema = createFinancialVoucherSchema.extend({
  adminPassword: z.string().optional().default("")
});

export const financialVoucherActionSchema = z.object({
  adminPassword: z.string().optional().default("")
});

export const addFinancialVoucherPaymentSchema = z.object({
  paymentDate: notFutureIsoDate,
  amount: positiveMoneyAmount,
  transactionNo: safeText.pipe(z.string().min(1)),
  remarks: optionalSafeText
});

export const updateFinancialVoucherPaymentSchema = addFinancialVoucherPaymentSchema.extend({
  adminPassword: z.string().optional().default("")
});

export const addOrganizerPaymentSchema = z.object({
  organizerId: z.string().trim().min(1),
  paymentDate: notFutureIsoDate,
  amount: positiveMoneyAmount,
  transactionNo: safeText.pipe(z.string().min(1)),
  remarks: optionalSafeText
});

export const updateOrganizerPaymentSchema = z.object({
  paymentDate: notFutureIsoDate,
  amount: positiveMoneyAmount,
  transactionNo: safeText.pipe(z.string().min(1)),
  remarks: optionalSafeText
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
    "CUSTOM_DATE_PAYMENT_REGISTER",
    "ORGANIZER_FARMER_PAYMENT_REGISTER",
    "ORGANIZER_PAYMENT_TRANSACTION_REPORT",
    "OVERPAID_FARMER_REPORT",
    "RECEIPT_VOUCHER_TRACEABILITY_REPORT",
    "ORGANIZER_INTAKE_PAYMENT_COMMISSION_REPORT",
    "REGISTRATION_PENDING_RECEIVED",
    "LOT_WISE_STOCK_LEDGER",
    "ADJUSTED_LOT_FORMATION_REGISTER",
    "ADJUSTED_LOT_LEDGER_FARMER_WISE",
    "STACK_WISE_STOCK_POSITION",
    "STACK_CARD_REGISTER",
    "DISCREPANCY_REGISTER"
  ]),
  season: z.string().optional().default(""),
  fromDate: z.union([isoDate, z.literal("")]).optional().default(""),
  toDate: z.union([isoDate, z.literal("")]).optional().default(""),
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
