import mongoose, { Schema } from "mongoose";

const registrationStatusValues = ["ACTIVE", "BLOCKED", "CLOSED", "EXHAUSTED"] as const;
const lotStatusValues = ["OPEN", "FULL", "VOID", "CANCELLED"] as const;
const discrepancyStatusValues = ["OPEN", "SHIFT_PENDING", "RESOLVED"] as const;

const UserSchema = new Schema(
  {
    name: { type: String, required: true },
    email: { type: String, required: true, unique: true },
    role: {
      type: String,
      enum: ["ADMIN", "MANAGER", "USER"],
      required: true
    },
    passwordHash: { type: String },
    isActive: { type: Boolean, default: true }
  },
  { timestamps: true }
);

const SeasonSchema = new Schema(
  {
    code: { type: String, required: true, unique: true },
    name: { type: String, required: true },
    startDate: { type: Date },
    endDate: { type: Date },
    isActive: { type: Boolean, default: false }
  },
  { timestamps: true }
);

const FarmerSchema = new Schema(
  {
    farmerCode: { type: String },
    name: { type: String, required: true },
    fatherHusbandName: { type: String },
    village: { type: String },
    block: { type: String },
    district: { type: String },
    mobile: { type: String }
  },
  { timestamps: true }
);

const ImportSchema = new Schema(
  {
    fileName: { type: String, required: true },
    importedAt: { type: Date, default: Date.now },
    seasonLabel: { type: String },
    rowCount: { type: Number, default: 0 },
    successCount: { type: Number, default: 0 },
    failedCount: { type: Number, default: 0 }
  },
  { timestamps: true }
);

const RegistrationSchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    season: { type: String, required: true },
    year: { type: String, required: true },
    seasonKey: { type: String, required: true, index: true },
    ro: { type: String },
    cropRegistrationCode: { type: String, required: true },
    farmerName: { type: String, required: true },
    fatherName: { type: String },
    village: { type: String },
    block: { type: String },
    district: { type: String },
    crop: { type: String, required: true },
    variety: { type: String, required: true },
    classStage: { type: String, required: true },
    registeredAreaHa: { type: Number, default: 0 },
    inspectedAreaHa: { type: Number, default: 0 },
    rejectedAreaHa: { type: Number, default: 0 },
    certifiedAreaHa: { type: Number, default: 0 },
    expectedYieldQtl: { type: Number, default: 0 },
    allowedIntakeQtl: { type: Number, default: 0 },
    totalReceivedQtl: { type: Number, default: 0 },
    balanceQtl: { type: Number, default: 0 },
    status: { type: String, enum: registrationStatusValues, required: true },
    sourceRowNumber: { type: Number, required: true },
    sourceImportId: { type: Schema.Types.ObjectId, ref: "ImportBatch" }
  },
  { timestamps: true }
);

const GodownSchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    name: { type: String, required: true, unique: true }
  },
  { timestamps: true }
);

const StackSchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    godownId: { type: String, required: true },
    stackNo: { type: String, required: true }
  },
  { timestamps: true }
);

const LotSchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    cropRegistrationId: { type: String, required: true, index: true },
    cropRegistrationCode: { type: String, required: true },
    lotNo: { type: Number, required: true },
    lotCode: { type: String, required: true, unique: true },
    godownId: { type: String, required: true },
    godownName: { type: String, required: true },
    stackId: { type: String, required: true },
    stackNo: { type: String, required: true },
    qualityStatus: { type: String, default: "ACCEPTED" },
    currentQtyQtl: { type: Number, required: true },
    maxAllowedQtyQtl: { type: Number, default: 200 },
    status: { type: String, enum: lotStatusValues, required: true },
    voidReason: { type: String, default: "" },
    createdAt: { type: String, required: true }
  },
  { timestamps: true }
);

const ReceiptLineAllocationSchema = new Schema(
  {
    lotId: { type: String, required: true },
    lotCode: { type: String, required: true },
    qtyQtl: { type: Number, required: true }
  },
  { _id: false }
);

const ReceiptLineSchema = new Schema(
  {
    id: { type: String, required: true },
    godownId: { type: String, required: true },
    stackId: { type: String, required: true },
    stackNo: { type: String, required: true },
    grossWeightQtl: { type: Number, required: true },
    qtyQtl: { type: Number, required: true },
    noOfBags: { type: Number, required: true },
    weightPerBagKg: { type: Number, required: true },
    netWeightQtl: { type: Number, required: true },
    moisturePercent: { type: Number, required: true },
    vehicleNo: { type: String, required: true },
    remarks: { type: String, default: "" },
    qualityStatus: { type: String, default: "ACCEPTED" },
    allocations: { type: [ReceiptLineAllocationSchema], default: [] }
  },
  { _id: false }
);

const ReceiptSchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    seasonKey: { type: String, required: true, index: true },
    season: { type: String, required: true },
    year: { type: String, required: true },
    receiptNo: { type: String, required: true },
    receiptSequenceNo: { type: Number, required: true, default: 0 },
    receiptDate: { type: String, required: true },
    cropRegistrationId: { type: String, required: true, index: true },
    cropRegistrationCode: { type: String, required: true },
    farmerName: { type: String, required: true },
    lines: { type: [ReceiptLineSchema], default: [] }
  },
  { timestamps: true }
);

const AuditLogSchema = new Schema(
  {
    entityType: { type: String, required: true },
    entityId: { type: String, required: true },
    action: { type: String, required: true },
    payload: { type: Schema.Types.Mixed }
  },
  { timestamps: true }
);

const DiscrepancySchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    discrepancyNo: { type: String, required: true, unique: true },
    season: { type: String, required: true },
    cropRegistrationId: { type: String, required: true, index: true },
    cropRegistrationCode: { type: String, required: true, index: true },
    farmerName: { type: String, required: true },
    receiptNo: { type: String, required: true, index: true },
    receiptDate: { type: String, required: true },
    godownId: { type: String, required: true },
    godownName: { type: String, required: true },
    stackId: { type: String, required: true },
    stackNo: { type: String, required: true },
    expectedQtyQtl: { type: Number, required: true },
    receiptNetQtyQtl: { type: Number, required: true },
    totalReceivedAfterReceiptQtl: { type: Number, required: true },
    excessQtyQtl: { type: Number, required: true },
    estimatedExcessBags: { type: Number, required: true },
    handlingMode: { type: String, default: "MARK_STACK_FOR_SHIFT" },
    status: { type: String, enum: discrepancyStatusValues, default: "OPEN" },
    remarks: { type: String, default: "" }
  },
  { timestamps: true }
);

const DiscrepancyShiftSchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    discrepancyId: { type: String, required: true, index: true },
    discrepancyNo: { type: String, required: true, index: true },
    cropRegistrationCode: { type: String, required: true },
    farmerName: { type: String, required: true },
    fromGodownName: { type: String, required: true },
    fromStackNo: { type: String, required: true },
    toGodownId: { type: String, required: true },
    toGodownName: { type: String, required: true },
    toStackId: { type: String, required: true },
    toStackNo: { type: String, required: true },
    shiftedQtyQtl: { type: Number, required: true },
    shiftedBags: { type: Number, required: true },
    shiftDate: { type: String, required: true },
    approvedBy: { type: String, default: "" },
    remarks: { type: String, default: "" }
  },
  { timestamps: true }
);

const FinancialVoucherLineSchema = new Schema(
  {
    receiptId: { type: String, required: true },
    receiptNo: { type: String, required: true },
    receiptDate: { type: String, required: true },
    vehicleNo: { type: String, default: "" },
    stackNo: { type: String, required: true },
    bags: { type: Number, required: true },
    grossQtyQtl: { type: Number, required: true },
    netQtyQtl: { type: Number, required: true }
  },
  { _id: false }
);

const FinancialVoucherSchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    voucherNo: { type: String, required: true, unique: true },
    voucherDate: { type: String, required: true },
    cropRegistrationId: { type: String, required: true, index: true },
    cropRegistrationCode: { type: String, required: true, index: true },
    farmerName: { type: String, required: true },
    fatherName: { type: String, default: "" },
    village: { type: String, default: "" },
    block: { type: String, default: "" },
    district: { type: String, default: "" },
    season: { type: String, required: true },
    year: { type: String, required: true },
    crop: { type: String, required: true },
    variety: { type: String, required: true },
    classStage: { type: String, required: true },
    expectedYieldQtl: { type: Number, required: true },
    totalGrossQtyQtl: { type: Number, required: true },
    totalNetQtyQtl: { type: Number, required: true },
    totalBags: { type: Number, required: true },
    certifiedQtyQtl: { type: Number, required: true },
    discrepancyQtyQtl: { type: Number, required: true },
    discrepancyBags: { type: Number, required: true },
    certifiedRatePerQtl: { type: Number, required: true },
    discrepancyRatePerQtl: { type: Number, required: true, default: 0 },
    certifiedAmount: { type: Number, required: true },
    discrepancyAmount: { type: Number, required: true },
    grossPayableAmount: { type: Number, required: true },
    deductionAmount: { type: Number, required: true, default: 0 },
    netPayableAmount: { type: Number, required: true },
    status: { type: String, required: true, default: "DRAFT" },
    remarks: { type: String, default: "" },
    lines: { type: [FinancialVoucherLineSchema], default: [] }
  },
  { timestamps: true }
);
FinancialVoucherSchema.index(
  { cropRegistrationId: 1 },
  { unique: true, name: "cropRegistrationId_1_unique_voucher" }
);

const NonCertificationStockMovementSchema = new Schema(
  {
    id: { type: String, required: true, unique: true },
    movementType: { type: String, required: true, default: "DISCREPANCY_SHIFT_OUT" },
    referenceType: { type: String, required: true, default: "DISCREPANCY" },
    referenceId: { type: String, required: true, index: true },
    cropRegistrationCode: { type: String, required: true, index: true },
    farmerName: { type: String, required: true },
    fromGodownId: { type: String, required: true },
    fromGodownName: { type: String, required: true },
    fromStackId: { type: String, required: true },
    fromStackNo: { type: String, required: true },
    toGodownId: { type: String, required: true },
    toGodownName: { type: String, required: true },
    toStackId: { type: String, required: true },
    toStackNo: { type: String, required: true },
    qtyQtl: { type: Number, required: true },
    bags: { type: Number, required: true },
    movementDate: { type: String, required: true },
    remarks: { type: String, default: "" }
  },
  { timestamps: true }
);

RegistrationSchema.index({ seasonKey: 1, cropRegistrationCode: 1 }, { unique: true });
ReceiptSchema.index({ seasonKey: 1, receiptNo: 1 }, { unique: true });
StackSchema.index({ godownId: 1, stackNo: 1 }, { unique: true });

export const UserModel = mongoose.models.User ?? mongoose.model("User", UserSchema);
export const SeasonModel = mongoose.models.Season ?? mongoose.model("Season", SeasonSchema);
export const FarmerModel = mongoose.models.Farmer ?? mongoose.model("Farmer", FarmerSchema);
export const ImportBatchModel =
  mongoose.models.ImportBatch ?? mongoose.model("ImportBatch", ImportSchema);
export const RegistrationModel =
  mongoose.models.Registration ?? mongoose.model("Registration", RegistrationSchema);
export const GodownModel = mongoose.models.Godown ?? mongoose.model("Godown", GodownSchema);
export const StackModel = mongoose.models.Stack ?? mongoose.model("Stack", StackSchema);
export const LotModel = mongoose.models.Lot ?? mongoose.model("Lot", LotSchema);
export const ReceiptModel = mongoose.models.Receipt ?? mongoose.model("Receipt", ReceiptSchema);
export const AuditLogModel =
  mongoose.models.AuditLog ?? mongoose.model("AuditLog", AuditLogSchema);
export const DiscrepancyModel =
  mongoose.models.Discrepancy ?? mongoose.model("Discrepancy", DiscrepancySchema);
export const DiscrepancyShiftModel =
  mongoose.models.DiscrepancyShift ??
  mongoose.model("DiscrepancyShift", DiscrepancyShiftSchema);
export const NonCertificationStockMovementModel =
  mongoose.models.NonCertificationStockMovement ??
  mongoose.model("NonCertificationStockMovement", NonCertificationStockMovementSchema);
export const FinancialVoucherModel =
  mongoose.models.FinancialVoucher ??
  mongoose.model("FinancialVoucher", FinancialVoucherSchema);
