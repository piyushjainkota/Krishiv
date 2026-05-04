import { randomUUID } from "node:crypto";
import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, access, stat } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import * as XLSX from "xlsx";
import type {
  AddOrganizerPaymentInput,
  AddFinancialVoucherPaymentInput,
  AssignRegistrationOrganizerInput,
  BackupDatabaseInput,
  CreateOrganizerInput,
  CreateDiscrepancyShiftInput,
  CreateStackAccommodationInput,
  CreateFinancialVoucherInput,
  CreateGodownInput,
  CreateReceiptInput,
  CreateStackInput,
  FinancialVoucherActionInput,
  ImportRegistrationsInput,
  LoginInput,
  ReportFilterInput,
  RestoreDatabaseInput,
  UpdateOrganizerInput,
  UpdateOrganizerPaymentInput,
  UpdateFinancialVoucherPaymentInput,
  UpdateFinancialVoucherInput,
  UpdateStackAccommodationInput,
  UpdateReceiptInput
} from "./seed.schemas";
import { config } from "../../config";
import {
  AuditLogModel,
  CounterModel,
  DiscrepancyModel,
  DiscrepancyShiftModel,
  FinancialVoucherModel,
  GodownModel,
  ImportBatchModel,
  LotModel,
  NonCertificationStockMovementModel,
  OrganizerModel,
  OrganizerPaymentModel,
  ReceiptModel,
  RegistrationModel,
  StackAccommodationModel,
  StackModel,
  UserModel
} from "../../models/seed.models";
import mongoose, { type ClientSession } from "mongoose";

const defaultGodowns = [
  { id: "godown-1", name: "Main Godown" },
  { id: "godown-2", name: "North Godown" }
];

const defaultStacks = [
  { id: "stack-1", godownId: "godown-1", stackNo: "A-01" },
  { id: "stack-2", godownId: "godown-1", stackNo: "A-02" },
  { id: "stack-3", godownId: "godown-2", stackNo: "B-01" }
];

const defaultUsers = [
  {
    name: "System Admin",
    email: "admin",
    role: "ADMIN",
    passwordHash: "admin123"
  },
  {
    name: "Operations Manager",
    email: "manager",
    role: "MANAGER",
    passwordHash: "manager123"
  },
  {
    name: "Intake User",
    email: "user",
    role: "USER",
    passwordHash: "user123"
  }
] as const;

const execFileAsync = promisify(execFile);

type AppRole = "ADMIN" | "MANAGER" | "USER";
type ActorUser = { email: string; role: AppRole };

const permissionMatrix: Record<
  AppRole,
  {
    canEntry: boolean;
    canEdit: boolean;
    canDelete: boolean;
    canImport: boolean;
    canValidate: boolean;
    canVoucher: boolean;
    canShift: boolean;
    canMaintenance: boolean;
  }
> = {
  ADMIN: {
    canEntry: true,
    canEdit: true,
    canDelete: true,
    canImport: true,
    canValidate: true,
    canVoucher: true,
    canShift: true,
    canMaintenance: true
  },
  MANAGER: {
    canEntry: true,
    canEdit: true,
    canDelete: false,
    canImport: false,
    canValidate: false,
    canVoucher: true,
    canShift: true,
    canMaintenance: false
  },
  USER: {
    canEntry: true,
    canEdit: false,
    canDelete: false,
    canImport: false,
    canValidate: false,
    canVoucher: false,
    canShift: false,
    canMaintenance: false
  }
};

type RegistrationRecord = ImportRegistrationsInput["registrations"][number];

function roundQtl(value: number): number {
  return Number(value.toFixed(2));
}

function normalizeOrganizerName(value: string): string {
  return value.trim().replace(/\s+/g, " ");
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function buildCaseInsensitiveExactRegex(value: string): RegExp {
  return new RegExp(`^${escapeRegex(value)}$`, "i");
}

function calculateNetWeightQtl(grossWeightQtl: number, noOfBags: number) {
  return roundQtl(grossWeightQtl - grossWeightQtl * 0.003 - noOfBags * 0.007);
}

function calculateWeightPerBagKg(netWeightQtl: number, noOfBags: number) {
  if (noOfBags <= 0) {
    return 0;
  }

  return roundQtl((netWeightQtl * 100) / noOfBags);
}

function buildSeasonKey(season: string, year: string) {
  return `${String(season).trim().toUpperCase()}::${String(year).trim().toUpperCase()}`;
}

function nextDiscrepancyNumber(discrepancyNos: string[]) {
  const max = discrepancyNos.reduce((current, value) => {
    const parsed = Number(String(value).replace(/\D/g, ""));
    return Number.isFinite(parsed) ? Math.max(current, parsed) : current;
  }, 0);

  return `D${String(max + 1).padStart(4, "0")}`;
}

function formatVoucherSeasonPart(season: string, year: string) {
  const normalizedSeason = String(season).trim().toUpperCase();
  const normalizedYear = String(year).trim();
  const match = normalizedYear.match(/^(\d{4})-(\d{2,4})$/);
  if (match) {
    return `${normalizedSeason}${match[1].slice(2)}-${match[2].slice(-2)}`;
  }
  return `${normalizedSeason}${normalizedYear}`;
}

function timestampForFolder() {
  const now = new Date();
  const pad = (value: number) => String(value).padStart(2, "0");
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

function resolveMongoTool(toolName: "mongodump" | "mongorestore") {
  const configuredBin = process.env.MONGODB_BIN;
  if (configuredBin) {
    return path.join(configuredBin, process.platform === "win32" ? `${toolName}.exe` : toolName);
  }
  const knownBinPaths = [
    "C:\\Users\\piyus\\Downloads\\mongodb-database-tools-windows-x86_64-100.16.0\\mongodb-database-tools-windows-x86_64-100.16.0\\bin"
  ];
  for (const binPath of knownBinPaths) {
    const candidate = path.join(
      binPath,
      process.platform === "win32" ? `${toolName}.exe` : toolName
    );
    if (existsSync(candidate)) {
      return candidate;
    }
  }
  return process.platform === "win32" ? `${toolName}.exe` : toolName;
}

function deriveRegistrationStatus(params: {
  existingStatus?: RegistrationRecord["status"];
  certifiedAreaHa: number;
  expectedYieldQtl: number;
  balanceQtl: number;
}): RegistrationRecord["status"] {
  const { existingStatus, certifiedAreaHa, expectedYieldQtl, balanceQtl } = params;

  if (existingStatus === "CLOSED") {
    return "CLOSED";
  }

  if (certifiedAreaHa <= 0 || expectedYieldQtl <= 0) {
    return "BLOCKED";
  }

  if (balanceQtl <= 0) {
    return "EXHAUSTED";
  }

  return "ACTIVE";
}

function deriveVoucherStatus(finalPayableAmount: number, totalPaidAmount: number) {
  if (finalPayableAmount < 0) {
    return "OVERPAID";
  }
  if (totalPaidAmount <= 0) {
    return "DRAFT";
  }
  if (totalPaidAmount > finalPayableAmount) {
    return "OVERPAID";
  }
  if (totalPaidAmount === finalPayableAmount) {
    return "PAID";
  }
  return "PART PAID";
}

function isVoucherLockedStatus(status: string) {
  return status === "PAID" || status === "OVERPAID";
}

function nextRegistrationState(registration: RegistrationRecord, addedQtyQtl: number) {
  const totalReceivedQtl = roundQtl(registration.totalReceivedQtl + addedQtyQtl);
  const balanceQtl = roundQtl(Math.max(registration.allowedIntakeQtl - totalReceivedQtl, 0));

  return {
    totalReceivedQtl,
    balanceQtl,
    status: deriveRegistrationStatus({
      existingStatus: registration.status,
      certifiedAreaHa: registration.certifiedAreaHa,
      expectedYieldQtl: registration.expectedYieldQtl,
      balanceQtl
    })
  };
}

function estimateExcessBags(excessQtyQtl: number, weightPerBagKg: number) {
  if (excessQtyQtl <= 0 || weightPerBagKg <= 0) {
    return 0;
  }

  return Math.ceil((excessQtyQtl * 100) / weightPerBagKg);
}

function appendShiftRemark(existingRemark: string, shiftRemark: string, shiftedQtyQtl: number, shiftedBags: number) {
  const parts = [existingRemark?.trim(), `Shifted ${shiftedQtyQtl} QTL / ${shiftedBags} bags.`];
  if (shiftRemark?.trim()) {
    parts.push(shiftRemark.trim());
  }

  return parts.filter(Boolean).join(" ");
}

function sumReceiptNetQty(receipt: { lines: { qtyQtl?: number }[] }) {
  return roundQtl(
    receipt.lines.reduce((sum, line) => sum + Number(line.qtyQtl ?? 0), 0)
  );
}

function averageWeightPerBagKg(receipt: { lines: { netWeightQtl?: number; noOfBags?: number }[] }) {
  const totalNetQty = receipt.lines.reduce((sum, line) => sum + Number(line.netWeightQtl ?? 0), 0);
  const totalBags = receipt.lines.reduce((sum, line) => sum + Number(line.noOfBags ?? 0), 0);

  return calculateWeightPerBagKg(roundQtl(totalNetQty), totalBags);
}

type ReportType =
  | "GODOWN_WISE_DETAIL"
  | "DISTRICT_WISE_DETAIL"
  | "FARMER_WISE_DETAIL"
  | "OVERALL_INTAKE"
  | "SUMMARY"
  | "DAILY_INTAKE_REGISTER"
  | "REGISTRATION_PENDING_RECEIVED"
  | "LOT_WISE_STOCK_LEDGER"
  | "STACK_WISE_STOCK_POSITION"
  | "STACK_CARD_REGISTER"
  | "DISCREPANCY_REGISTER";

type ReportPreview = {
  reportType: ReportType;
  title: string;
  columns: string[];
  rows: Record<string, string | number>[];
  totals: Record<string, string | number>;
  generatedAt: string;
  fileName: string;
};

type BaseReceiptRow = {
  seasonLabel: string;
  receiptNo: string;
  receiptDate: string;
  cropRegistrationId: string;
  cropRegistrationCode: string;
  farmerName: string;
  fatherName: string;
  village: string;
  block: string;
  district: string;
  crop: string;
  variety: string;
  classStage: string;
  certifiedAreaHa: number;
  expectedYieldQtl: number;
  allowedIntakeQtl: number;
  registrationReceivedQtl: number;
  registrationBalanceQtl: number;
  godownId: string;
  godownName: string;
  stackId: string;
  stackNo: string;
  vehicleNo: string;
  moisturePercent: number;
  grossWeightQtl: number;
  netWeightQtl: number;
  noOfBags: number;
  weightPerBagKg: number;
  acceptedGrossQtl: number;
  acceptedNetQtl: number;
  acceptedBags: number;
  discrepancyGrossQtl: number;
  discrepancyNetQtl: number;
  discrepancyBags: number;
  discrepancyStatus: string;
  discrepancyQtyQtl: number;
  discrepancyShiftedQtyQtl: number;
  remarks: string;
  lotCodes: string[];
};

type FinancialVoucherPreview = {
  id: string;
  voucherNo: string;
  voucherDate: string;
  cropRegistrationId: string;
  cropRegistrationCode: string;
  farmerName: string;
  fatherName: string;
  village: string;
  block: string;
  district: string;
  season: string;
  year: string;
  crop: string;
  variety: string;
  classStage: string;
  expectedYieldQtl: number;
  totalGrossQtyQtl: number;
  totalNetQtyQtl: number;
  totalBags: number;
  certifiedQtyQtl: number;
  discrepancyQtyQtl: number;
  discrepancyBags: number;
  certifiedRatePerQtl: number;
  discrepancyRatePerQtl: number;
  certifiedAmount: number;
  discrepancyAmount: number;
  grossPayableAmount: number;
  deductionAmount: number;
  netPayableAmount: number;
  roundedOffAmount: number;
  finalPayableAmount: number;
  totalPaidAmount: number;
  balanceAmount: number;
  lastPaymentDate: string;
  status: string;
  remarks: string;
  lines: {
    receiptId: string;
    receiptNo: string;
    receiptDate: string;
    vehicleNo: string;
    stackNo: string;
    bags: number;
    grossQtyQtl: number;
    netQtyQtl: number;
  }[];
  payments: {
    id: string;
    paymentDate: string;
    amount: number;
    transactionNo: string;
    mode: string;
    remarks: string;
    createdBy?: string;
    updatedBy?: string;
  }[];
};

type FinancialVoucherRecord = FinancialVoucherPreview;

type OrganizerRecord = {
  id: string;
  name: string;
  mobile?: string;
  village?: string;
  district?: string;
  commissionRatePerQtl: number;
  deductionAmount?: number;
  isActive: boolean;
};

type OrganizerPaymentRecord = {
  id: string;
  organizerId: string;
  organizerName: string;
  paymentDate: string;
  amount: number;
  transactionNo: string;
  remarks: string;
  createdBy?: string;
  updatedBy?: string;
};

function containsText(value: string | undefined, filterValue: string | undefined) {
  if (!filterValue?.trim()) {
    return true;
  }

  return String(value ?? "").toLowerCase().includes(filterValue.trim().toLowerCase());
}

function dateWithinRange(dateValue: string, fromDate?: string, toDate?: string) {
  if (fromDate && dateValue < fromDate) {
    return false;
  }

  if (toDate && dateValue > toDate) {
    return false;
  }

  return true;
}

function roundInteger(value: number) {
  return Math.round(value);
}

function formatReportFileName(season: string, suffix: string) {
  return `${suffix}_${season.replace(/\s+/g, "_")}.xlsx`;
}

export class SeedService {
  private compatibilityReady = false;
  private usersReady = false;

  private async ensureDefaultUsers() {
    if (this.usersReady) {
      return;
    }

    for (const user of defaultUsers) {
      await UserModel.updateOne(
        { email: user.email },
        {
          $setOnInsert: {
            name: user.name,
            email: user.email,
            role: user.role,
            passwordHash: user.passwordHash,
            isActive: true
          }
        },
        { upsert: true }
      );
    }

    this.usersReady = true;
  }

  private async ensureCompatibility() {
    if (this.compatibilityReady) {
      await this.ensureDefaultUsers();
      return;
    }

    const registrationCursor = RegistrationModel.find({}, {
      id: 1,
      season: 1,
      year: 1,
      cropRegistrationCode: 1
    }).lean();
    for (const registration of await registrationCursor) {
      const seasonKey = buildSeasonKey(String(registration.season ?? ""), String(registration.year ?? ""));
      await RegistrationModel.updateOne(
        { id: registration.id },
        {
          $set: {
            seasonKey
          }
        }
      );
    }

    const registrations = await RegistrationModel.find({}, {
      id: 1,
      season: 1,
      year: 1
    }).lean();
    const registrationById = new Map(
      registrations.map((item) => [String(item.id), item])
    );
    const receipts = await ReceiptModel.find({}, {
      id: 1,
      cropRegistrationId: 1,
      receiptNo: 1
    }).lean();
    const seasonSequenceMap = new Map<string, number>();
    for (const receipt of receipts) {
      const registration = registrationById.get(String(receipt.cropRegistrationId));
      const season = String(registration?.season ?? "");
      const year = String(registration?.year ?? "");
      const seasonKey = buildSeasonKey(season, year);
      const parsedReceiptNo = Number(String(receipt.receiptNo ?? ""));
      const nextSequenceNo = Number.isFinite(parsedReceiptNo) && parsedReceiptNo > 0
        ? parsedReceiptNo
        : (seasonSequenceMap.get(seasonKey) ?? 0) + 1;
      seasonSequenceMap.set(seasonKey, Math.max(seasonSequenceMap.get(seasonKey) ?? 0, nextSequenceNo));
      await ReceiptModel.updateOne(
        { id: receipt.id },
        {
          $set: {
            seasonKey,
            season,
            year,
            receiptSequenceNo: nextSequenceNo
          }
        }
      );
    }

    const connection = mongoose.connection;
    const registrationCollection = connection.collection("registrations");
    const receiptCollection = connection.collection("receipts");
    const stackCollection = connection.collection("stacks");
    const financialVoucherCollection = connection.collection("financialvouchers");

    const dropIndexIfExists = async (collection: mongoose.mongo.Collection, indexName: string) => {
      try {
        const indexes = await collection.indexes();
        if (indexes.some((item) => item.name === indexName)) {
          await collection.dropIndex(indexName);
        }
      } catch {
        // Keep startup resilient; index may not exist yet.
      }
    };

    await dropIndexIfExists(registrationCollection, "cropRegistrationCode_1");
    await dropIndexIfExists(receiptCollection, "receiptNo_1");

    const safeCreateIndex = async (
      collection: mongoose.mongo.Collection,
      spec: Record<string, 1 | -1>,
      options: mongoose.mongo.CreateIndexesOptions
    ) => {
      try {
        await collection.createIndex(spec, options);
      } catch {
        // Do not block startup; legacy duplicates can be cleaned through validation later.
      }
    };

    await safeCreateIndex(
      registrationCollection,
      { seasonKey: 1, cropRegistrationCode: 1 },
      { unique: true, name: "seasonKey_1_cropRegistrationCode_1" }
    );
    await safeCreateIndex(
      receiptCollection,
      { seasonKey: 1, receiptNo: 1 },
      { unique: true, name: "seasonKey_1_receiptNo_1" }
    );
    await safeCreateIndex(
      stackCollection,
      { godownId: 1, stackNo: 1 },
      { unique: true, name: "godownId_1_stackNo_1" }
    );
    await safeCreateIndex(
      financialVoucherCollection,
      { cropRegistrationId: 1 },
      { unique: true, name: "cropRegistrationId_1_unique_voucher" }
    );

    this.compatibilityReady = true;
    await this.ensureDefaultUsers();
  }

  private async runInTransaction<T>(work: (session: ClientSession | null) => Promise<T>) {
    const session = await mongoose.startSession();
    try {
      let result: T | undefined;
      try {
        await session.withTransaction(async () => {
          result = await work(session);
        });
        return result as T;
      } catch (error) {
        const message = error instanceof Error ? error.message : "";
        const unsupported =
          message.includes("Transaction numbers are only allowed") ||
          message.includes("replica set") ||
          message.includes("Transaction");
        if (!unsupported) {
          throw error;
        }
        return work(null);
      }
    } finally {
      await session.endSession();
    }
  }

  private async findActiveUser(email: string, role?: AppRole) {
    const query: { email: string; isActive: boolean; role?: AppRole } = {
      email: email.trim(),
      isActive: true
    };
    if (role) {
      query.role = role;
    }
    return UserModel.findOne(query).lean<{
      name: string;
      email: string;
      role: AppRole;
      passwordHash?: string;
      isActive: boolean;
    } | null>();
  }

  private async verifyAdminPassword(adminPassword?: string) {
    const password = String(adminPassword ?? "").trim();
    if (!password) {
      throw new Error("Admin password is required for this action.");
    }

    const adminUser = await this.findActiveUser("admin", "ADMIN");
    if (!adminUser || String(adminUser.passwordHash ?? "") !== password) {
      throw new Error("Admin password is invalid.");
    }
  }

  private buildVoucherPaymentSummary(
    payments: {
      paymentDate: string;
      amount: number;
    }[],
    finalPayableAmount: number
  ) {
    const totalPaidAmount = roundQtl(
      payments.reduce((sum, item) => sum + Number(item.amount ?? 0), 0)
    );
    const sortedDates = payments
      .map((item) => String(item.paymentDate ?? ""))
      .filter(Boolean)
      .sort();
    return {
      totalPaidAmount,
      balanceAmount: roundQtl(finalPayableAmount - totalPaidAmount),
      lastPaymentDate: sortedDates[sortedDates.length - 1] ?? "",
      status: deriveVoucherStatus(finalPayableAmount, totalPaidAmount)
    };
  }

  private async assertVoucherPaymentEditable(
    voucher: { status?: string },
    payment: { createdBy?: string },
    actor: ActorUser,
    adminPassword?: string
  ) {
    if (isVoucherLockedStatus(String(voucher.status ?? "DRAFT"))) {
      await this.verifyAdminPassword(adminPassword);
    }

    const paymentCreator = String(payment.createdBy ?? "").trim().toLowerCase();
    if (actor.role !== "ADMIN" && paymentCreator && paymentCreator !== actor.email.trim().toLowerCase()) {
      throw new Error("Only the original payment entry user or admin can edit this ledger entry.");
    }
  }

  private assertOrganizerPaymentEditable(
    payment: { createdBy?: string },
    actor: ActorUser
  ) {
    const paymentCreator = String(payment.createdBy ?? "").trim().toLowerCase();
    if (actor.role !== "ADMIN" && paymentCreator && paymentCreator !== actor.email.trim().toLowerCase()) {
      throw new Error("Only the original commission entry user or admin can edit this commission payment.");
    }
  }

  private async assertStackAccommodationEditable(
    accommodation: { createdBy?: string },
    actor: ActorUser,
    adminPassword?: string
  ) {
    const creator = String(accommodation.createdBy ?? "").trim().toLowerCase();
    if (actor.role !== "ADMIN" && creator && creator !== actor.email.trim().toLowerCase()) {
      if (!adminPassword) {
        throw new Error("Only the original mapping user or admin can edit this accommodation entry.");
      }
      await this.verifyAdminPassword(adminPassword);
      return;
    }
    if (actor.role === "ADMIN" && adminPassword) {
      await this.verifyAdminPassword(adminPassword);
    }
  }

  private async getOrganizerCommissionPosition(organizerId: string, excludedPaymentId = "") {
    const [organizer, linkedRegistrations, organizerPayments] = await Promise.all([
      OrganizerModel.findOne({ id: organizerId })
        .select({ deductionAmount: 1 })
        .lean<{ deductionAmount?: number } | null>(),
      RegistrationModel.find({ organizerId }).select({
        totalReceivedQtl: 1,
        organizerCommissionRatePerQtl: 1
      }).lean<Array<{ totalReceivedQtl?: number; organizerCommissionRatePerQtl?: number }>>(),
      OrganizerPaymentModel.find(
        excludedPaymentId ? { organizerId, id: { $ne: excludedPaymentId } } : { organizerId }
      )
        .select({ amount: 1 })
        .lean<Array<{ amount?: number }>>()
    ]);

    const payableAmount = roundQtl(
      linkedRegistrations.reduce(
        (sum, item) =>
          sum +
          Number(item.totalReceivedQtl ?? 0) * Number(item.organizerCommissionRatePerQtl ?? 0),
        0
      )
    );
    const deductionAmount = roundQtl(Number(organizer?.deductionAmount ?? 0));
    const netPayableAmount = roundQtl(payableAmount - deductionAmount);
    const paidAmount = roundQtl(
      organizerPayments.reduce((sum, item) => sum + Number(item.amount ?? 0), 0)
    );

    return {
      payableAmount,
      deductionAmount,
      netPayableAmount,
      paidAmount,
      balanceAmount: roundQtl(netPayableAmount - paidAmount)
    };
  }

  async loginUser(input: LoginInput) {
    await this.ensureCompatibility();
    const user = await this.findActiveUser(input.email);
    if (!user || String(user.passwordHash ?? "") !== input.password) {
      throw new Error("Invalid login credentials.");
    }

    return {
      user: {
        name: user.name,
        email: user.email,
        role: user.role
      },
      permissions: permissionMatrix[user.role]
    };
  }

  async backupDatabase(input: BackupDatabaseInput) {
    await this.ensureCompatibility();
    const rootDirectory = path.resolve(input.backupDirectory.trim());
    await mkdir(rootDirectory, { recursive: true });
    const backupDirectory = path.join(rootDirectory, `krishiv_seed_backup_${timestampForFolder()}`);
    await mkdir(backupDirectory, { recursive: true });

    try {
      await execFileAsync(resolveMongoTool("mongodump"), [
        `--uri=${config.mongodbUri}`,
        `--out=${backupDirectory}`
      ]);
    } catch (error) {
      throw new Error(
        `Database backup failed. Ensure MongoDB Database Tools are installed and mongodump is available. ${error instanceof Error ? error.message : ""}`.trim()
      );
    }

    await AuditLogModel.create({
      entityType: "DATABASE_BACKUP",
      entityId: backupDirectory,
      action: "CREATED",
      payload: {
        backupDirectory
      }
    });

    return {
      backupDirectory,
      message: "Database backup completed."
    };
  }

  async restoreDatabase(input: RestoreDatabaseInput) {
    await this.ensureCompatibility();
    const providedDirectory = path.resolve(input.restoreDirectory.trim());
    await access(providedDirectory);
    const providedStats = await stat(providedDirectory);
    if (!providedStats.isDirectory()) {
      throw new Error("Restore directory must be a valid folder.");
    }

    let restoreDirectory = providedDirectory;
    const basename = path.basename(providedDirectory).toLowerCase();
    if (basename === "krishiv_seed") {
      restoreDirectory = path.dirname(providedDirectory);
    }

    try {
      await execFileAsync(resolveMongoTool("mongorestore"), [
        `--uri=${config.mongodbUri}`,
        "--drop",
        `--dir=${restoreDirectory}`
      ]);
    } catch (error) {
      throw new Error(
        `Database restore failed. Ensure MongoDB Database Tools are installed and the restore folder is correct. ${error instanceof Error ? error.message : ""}`.trim()
      );
    }

    this.compatibilityReady = false;
    this.usersReady = false;
    await this.ensureCompatibility();

    await AuditLogModel.create({
      entityType: "DATABASE_RESTORE",
      entityId: restoreDirectory,
      action: "RESTORED",
      payload: {
        restoreDirectory
      }
    });

    return {
      restoreDirectory,
      message: "Database restore completed."
    };
  }

  async authorizeUser(email: string, role: string, allowedRoles: AppRole[]) {
    await this.ensureCompatibility();
    if (!email || !role) {
      throw new Error("Login required.");
    }
    if (!allowedRoles.includes(role as AppRole)) {
      throw new Error("You do not have permission for this action.");
    }
    const user = await this.findActiveUser(email, role as AppRole);
    if (!user) {
      throw new Error("User session is invalid.");
    }
    return {
      user: {
        name: user.name,
        email: user.email,
        role: user.role
      },
      permissions: permissionMatrix[user.role]
    };
  }

  private async findReceiptByReference(
    receiptRef: string,
    cropRegistrationId?: string,
    session?: ClientSession | null
  ) {
    const query = ReceiptModel.findOne({
      $or: [
        { id: receiptRef },
        ...(cropRegistrationId
          ? [{ receiptNo: receiptRef, cropRegistrationId }]
          : [{ receiptNo: receiptRef }])
      ]
    });
    if (session) {
      query.session(session);
    }
    return query.lean();
  }

  async bootstrap() {
    const [core, operations] = await Promise.all([
      this.bootstrapCore(),
      this.bootstrapOperational()
    ]);

    return {
      ...core,
      ...operations
    };
  }

  async bootstrapCore() {
    await this.ensureCompatibility();
    if ((await GodownModel.countDocuments()) === 0) {
      await GodownModel.insertMany(defaultGodowns);
    }

    if ((await StackModel.countDocuments()) === 0) {
      await StackModel.insertMany(defaultStacks);
    }

    const [registrations, godowns, stacks, organizers] = await Promise.all([
      RegistrationModel.find().sort({ cropRegistrationCode: 1 }).lean(),
      GodownModel.find().sort({ name: 1 }).lean(),
      StackModel.find().sort({ stackNo: 1 }).lean(),
      OrganizerModel.find().sort({ name: 1 }).lean()
    ]);

    return {
      registrations,
      godowns,
      stacks,
      organizers,
      features: {
        discrepancyWorkflow: config.enableDiscrepancyWorkflow
      }
    };
  }

  async bootstrapOperational() {
    await this.ensureCompatibility();
    const [
      lots,
      receipts,
      discrepancies,
      discrepancyShifts,
      stackAccommodations,
      financialVouchers,
      organizerPayments
    ] = await Promise.all([
      LotModel.find().sort({ lotCode: 1 }).lean(),
      ReceiptModel.find().sort({ createdAt: -1 }).lean(),
      DiscrepancyModel.find().sort({ createdAt: -1 }).lean(),
      DiscrepancyShiftModel.find().sort({ createdAt: -1 }).lean(),
      StackAccommodationModel.find().sort({ adjustmentDate: -1, createdAt: -1 }).lean(),
      FinancialVoucherModel.find().sort({ createdAt: -1 }).lean(),
      OrganizerPaymentModel.find().sort({ paymentDate: -1, createdAt: -1 }).lean()
    ]);

    return {
      lots,
      receipts,
      discrepancies,
      discrepancyShifts,
      stackAccommodations,
      financialVouchers,
      organizerPayments,
      features: {
        discrepancyWorkflow: config.enableDiscrepancyWorkflow
      }
    };
  }

  async importRegistrations(input: ImportRegistrationsInput) {
    if (input.importPassword !== config.importAdminPassword) {
      throw new Error("Import authorization password is invalid.");
    }
    const importBatch = await ImportBatchModel.create({
      fileName: input.fileName,
      seasonLabel: input.seasonLabel,
      rowCount: input.registrations.length,
      successCount: input.registrations.length,
      failedCount: 0
    });

    const existingRegistrations = await RegistrationModel.find(
      {},
      {
        id: 1,
        seasonKey: 1,
        cropRegistrationCode: 1,
        totalReceivedQtl: 1,
        status: 1,
        organizerId: 1,
        organizerName: 1,
        organizerCommissionRatePerQtl: 1
      }
    ).lean();

    const existingByCode = new Map(
      existingRegistrations.map((registration) => [
        `${registration.seasonKey}::${registration.cropRegistrationCode}`,
        registration
      ])
    );

    let insertedCount = 0;
    let updatedCount = 0;

    await RegistrationModel.bulkWrite(
      input.registrations.map((registration) => {
        const seasonKey = buildSeasonKey(registration.season, registration.year);
        const existing = existingByCode.get(`${seasonKey}::${registration.cropRegistrationCode}`);
        const preservedReceivedQtl = roundQtl(Number(existing?.totalReceivedQtl ?? 0));
        const allowedIntakeQtl = roundQtl(registration.allowedIntakeQtl);
        const balanceQtl = roundQtl(Math.max(allowedIntakeQtl - preservedReceivedQtl, 0));
        const status = deriveRegistrationStatus({
          existingStatus: existing?.status as RegistrationRecord["status"] | undefined,
          certifiedAreaHa: registration.certifiedAreaHa,
          expectedYieldQtl: registration.expectedYieldQtl,
          balanceQtl
        });

        if (existing) {
          updatedCount += 1;
        } else {
          insertedCount += 1;
        }

        return {
          updateOne: {
            filter: { seasonKey, cropRegistrationCode: registration.cropRegistrationCode },
            update: {
              $set: {
                ...registration,
                seasonKey,
                id: existing?.id ?? registration.id,
                totalReceivedQtl: preservedReceivedQtl,
                balanceQtl,
                organizerId: String(existing?.organizerId ?? registration.organizerId ?? ""),
                organizerName: String(existing?.organizerName ?? registration.organizerName ?? ""),
                organizerCommissionRatePerQtl: roundQtl(
                  Number(existing?.organizerCommissionRatePerQtl ?? registration.organizerCommissionRatePerQtl ?? 0)
                ),
                status,
                sourceImportId: importBatch._id
              }
            },
            upsert: true
          }
        };
      })
    );

    await AuditLogModel.create({
      entityType: "IMPORT",
      entityId: String(importBatch._id),
      action: "REGISTRATIONS_UPSERTED",
      payload: {
        fileName: input.fileName,
        count: input.registrations.length,
        insertedCount,
        updatedCount,
        note: "Registrations were upserted. Existing receipts and lots were preserved."
      }
    });

    return this.bootstrap();
  }

  async createGodown(input: CreateGodownInput) {
    const godown = await GodownModel.create({
      id: randomUUID(),
      name: input.name.trim()
    });

    await AuditLogModel.create({
      entityType: "GODOWN",
      entityId: godown.id,
      action: "CREATED",
      payload: godown.toObject()
    });

    return godown.toObject();
  }

  async createStack(input: CreateStackInput) {
    const normalizedStackNo = input.stackNo.trim();
    const existing = await StackModel.findOne({
      godownId: input.godownId,
      stackNo: normalizedStackNo
    }).lean();
    if (existing) {
      return existing;
    }

    const stack = await StackModel.create({
      id: randomUUID(),
      godownId: input.godownId,
      stackNo: normalizedStackNo
    });

    await AuditLogModel.create({
      entityType: "STACK",
      entityId: stack.id,
      action: "CREATED",
      payload: stack.toObject()
    });

    return stack.toObject();
  }

  async createOrganizer(input: CreateOrganizerInput) {
    await this.ensureCompatibility();
    const normalizedName = normalizeOrganizerName(input.name);
    const existing = await OrganizerModel.findOne({
      name: buildCaseInsensitiveExactRegex(normalizedName)
    }).lean();
    if (existing) {
      throw new Error(`Organizer ${normalizedName} already exists.`);
    }

    const organizer = await OrganizerModel.create({
      id: randomUUID(),
      name: normalizedName,
      mobile: input.mobile.trim(),
      village: input.village.trim(),
      district: input.district.trim(),
      commissionRatePerQtl: roundQtl(Number(input.commissionRatePerQtl ?? 0)),
      deductionAmount: roundQtl(Number(input.deductionAmount ?? 0)),
      isActive: Boolean(input.isActive)
    });

    await AuditLogModel.create({
      entityType: "ORGANIZER",
      entityId: organizer.id,
      action: "CREATED",
      payload: organizer.toObject()
    });

    return organizer.toObject();
  }

  async updateOrganizer(organizerId: string, input: UpdateOrganizerInput) {
    await this.ensureCompatibility();
    const existing = await OrganizerModel.findOne({ id: organizerId }).lean<OrganizerRecord | null>();
    if (!existing) {
      throw new Error("Organizer not found.");
    }

    const normalizedName = normalizeOrganizerName(input.name);
    const duplicate = await OrganizerModel.findOne({
      id: { $ne: organizerId },
      name: buildCaseInsensitiveExactRegex(normalizedName)
    }).lean();
    if (duplicate) {
      throw new Error(`Organizer ${normalizedName} already exists.`);
    }

    const nextOrganizer = {
      name: normalizedName,
      mobile: input.mobile.trim(),
      village: input.village.trim(),
      district: input.district.trim(),
      commissionRatePerQtl: roundQtl(Number(input.commissionRatePerQtl ?? 0)),
      deductionAmount: roundQtl(Number(input.deductionAmount ?? 0)),
      isActive: Boolean(input.isActive)
    };

    await OrganizerModel.updateOne(
      { id: organizerId },
      {
        $set: nextOrganizer
      }
    );

    await RegistrationModel.updateMany(
      { organizerId },
      {
        $set: {
          organizerName: nextOrganizer.name,
          organizerCommissionRatePerQtl: nextOrganizer.commissionRatePerQtl
        }
      }
    );

    await AuditLogModel.create({
      entityType: "ORGANIZER",
      entityId: organizerId,
      action: "UPDATED",
      payload: {
        before: existing,
        after: {
          ...existing,
          ...nextOrganizer
        }
      }
    });

    return {
      id: organizerId,
      ...nextOrganizer
    };
  }

  async deleteOrganizer(organizerId: string) {
    await this.ensureCompatibility();
    const existing = await OrganizerModel.findOne({ id: organizerId }).lean<OrganizerRecord | null>();
    if (!existing) {
      throw new Error("Organizer not found.");
    }

    const linkedRegistrations = await RegistrationModel.countDocuments({ organizerId });
    if (linkedRegistrations > 0) {
      throw new Error("Organizer is linked with farmer registrations. Remove mapping first.");
    }

    const paymentCount = await OrganizerPaymentModel.countDocuments({ organizerId });
    if (paymentCount > 0) {
      throw new Error("Organizer payment entries exist for this organizer. Delete is blocked.");
    }

    await OrganizerModel.deleteOne({ id: organizerId });

    await AuditLogModel.create({
      entityType: "ORGANIZER",
      entityId: organizerId,
      action: "DELETED",
      payload: existing
    });

    return { success: true };
  }

  async assignOrganizerToRegistration(registrationId: string, input: AssignRegistrationOrganizerInput) {
    await this.ensureCompatibility();
    const registration = await RegistrationModel.findOne({ id: registrationId }).lean<
      (RegistrationRecord & {
        organizerId?: string;
        organizerName?: string;
        organizerCommissionRatePerQtl?: number;
      }) | null
    >();
    if (!registration) {
      throw new Error("Registration not found.");
    }

    const organizerId = String(input.organizerId ?? "").trim();
    const existingSnapshot = {
      organizerId: String(registration.organizerId ?? ""),
      organizerName: String(registration.organizerName ?? ""),
      organizerCommissionRatePerQtl: roundQtl(Number(registration.organizerCommissionRatePerQtl ?? 0))
    };

    if (!organizerId) {
      await RegistrationModel.updateOne(
        { id: registrationId },
        {
          $set: {
            organizerId: "",
            organizerName: "",
            organizerCommissionRatePerQtl: 0
          }
        }
      );

      await AuditLogModel.create({
        entityType: "REGISTRATION_ORGANIZER",
        entityId: registrationId,
        action: "CLEARED",
        payload: {
          cropRegistrationCode: registration.cropRegistrationCode,
          before: existingSnapshot
        }
      });

      return { success: true };
    }

    const organizer = await OrganizerModel.findOne({ id: organizerId }).lean<OrganizerRecord | null>();
    if (!organizer) {
      throw new Error("Selected organizer not found.");
    }
    if (!organizer.isActive) {
      throw new Error("Selected organizer is inactive.");
    }

    await RegistrationModel.updateOne(
      { id: registrationId },
      {
        $set: {
          organizerId: organizer.id,
          organizerName: organizer.name,
          organizerCommissionRatePerQtl: roundQtl(Number(organizer.commissionRatePerQtl ?? 0))
        }
      }
    );

    await AuditLogModel.create({
      entityType: "REGISTRATION_ORGANIZER",
      entityId: registrationId,
      action: "ASSIGNED",
      payload: {
        cropRegistrationCode: registration.cropRegistrationCode,
        before: existingSnapshot,
        after: {
          organizerId: organizer.id,
          organizerName: organizer.name,
          organizerCommissionRatePerQtl: roundQtl(Number(organizer.commissionRatePerQtl ?? 0))
        }
      }
    });

    return { success: true };
  }

  private async ensureStack(godownId: string, stackNo: string, session?: ClientSession | null) {
    const normalizedStackNo = stackNo.trim();
    const existingQuery = StackModel.findOne({
      godownId,
      stackNo: normalizedStackNo
    });
    if (session) {
      existingQuery.session(session);
    }
    let stack = await existingQuery.lean();

    if (!stack) {
      try {
        stack = (
          await StackModel.create([{
            id: randomUUID(),
            godownId,
            stackNo: normalizedStackNo
          }], { session: session ?? undefined })
        )[0].toObject();
      } catch {
        const retryQuery = StackModel.findOne({
          godownId,
          stackNo: normalizedStackNo
        });
        if (session) {
          retryQuery.session(session);
        }
        stack = await retryQuery.lean();
      }
    }

    return stack;
  }

  private async reserveNextReceiptNumber(
    seasonKey: string,
    session?: ClientSession | null
  ) {
    const counterId = `RECEIPT_SEQ:${seasonKey}`;

    for (let attempt = 0; attempt < 5; attempt += 1) {
      const existingCounterQuery = CounterModel.findOne({ id: counterId });
      if (session) {
        existingCounterQuery.session(session);
      }
      const existingCounter = await existingCounterQuery.lean();

      if (!existingCounter) {
        const receiptMaxQuery = ReceiptModel.find({ seasonKey }, { receiptSequenceNo: 1, receiptNo: 1, _id: 0 })
          .sort({ receiptSequenceNo: -1, receiptNo: -1 })
          .limit(1);
        if (session) {
          receiptMaxQuery.session(session);
        }
        const currentMaxReceipt = await receiptMaxQuery.lean();
        const currentMaxValue = currentMaxReceipt.reduce((max, item) => {
          const sequenceNo = Number(item.receiptSequenceNo ?? 0);
          if (Number.isFinite(sequenceNo) && sequenceNo > 0) {
            return Math.max(max, sequenceNo);
          }
          const parsedReceiptNo = Number(String(item.receiptNo ?? ""));
          return Number.isFinite(parsedReceiptNo) ? Math.max(max, parsedReceiptNo) : max;
        }, 0);

        try {
          await CounterModel.create([{
            id: counterId,
            value: currentMaxValue
          }], { session: session ?? undefined });
        } catch {
          continue;
        }
      }

      const counter = await CounterModel.findOneAndUpdate(
        { id: counterId },
        { $inc: { value: 1 } },
        {
          new: true,
          session: session ?? undefined
        }
      ).lean();

      const nextValue = Number(counter?.value ?? 0);
      if (Number.isFinite(nextValue) && nextValue > 0) {
        return String(nextValue).padStart(3, "0");
      }
    }

    throw new Error("Unable to reserve the next receipt number.");
  }

  private async reserveNextLotNo(
    cropRegistrationId: string,
    session?: ClientSession | null
  ) {
    const counterId = `LOT_SEQ:${cropRegistrationId}`;

    for (let attempt = 0; attempt < 5; attempt += 1) {
      const existingCounterQuery = CounterModel.findOne({ id: counterId });
      if (session) {
        existingCounterQuery.session(session);
      }
      const existingCounter = await existingCounterQuery.lean();

      if (!existingCounter) {
        const maxLotQuery = LotModel.find({ cropRegistrationId }, { lotNo: 1, _id: 0 })
          .sort({ lotNo: -1 })
          .limit(1);
        if (session) {
          maxLotQuery.session(session);
        }
        const currentMaxLotNo = Number((await maxLotQuery.lean())[0]?.lotNo ?? 0);

        try {
          await CounterModel.create([{
            id: counterId,
            value: currentMaxLotNo
          }], { session: session ?? undefined });
        } catch {
          continue;
        }
      }

      const reserved = await CounterModel.findOneAndUpdate(
        { id: counterId },
        { $inc: { value: 1 } },
        {
          new: true,
          session: session ?? undefined
        }
      ).lean();

      const nextValue = Number(reserved?.value ?? 0);
      if (Number.isFinite(nextValue) && nextValue > 0) {
        return nextValue;
      }
    }

    throw new Error("Unable to reserve the next lot number.");
  }

  private async reserveNextVoucherNumber(
    season: string,
    year: string,
    session?: ClientSession | null
  ) {
    const seasonKey = buildSeasonKey(season, year);
    const counterId = `VOUCHER_SEQ:${seasonKey}`;

    for (let attempt = 0; attempt < 5; attempt += 1) {
      const existingCounterQuery = CounterModel.findOne({ id: counterId });
      if (session) {
        existingCounterQuery.session(session);
      }
      const existingCounter = await existingCounterQuery.lean();

      if (!existingCounter) {
        const prefix = `${formatVoucherSeasonPart(season, year)}/`;
        const existingVoucherNosQuery = FinancialVoucherModel.find(
          { season, year },
          { voucherNo: 1, _id: 0 }
        );
        if (session) {
          existingVoucherNosQuery.session(session);
        }
        const existingVoucherNos = (await existingVoucherNosQuery.lean()).map((item) =>
          String(item.voucherNo ?? "")
        );
        const currentMax = existingVoucherNos.reduce((max, value) => {
          if (!value.startsWith(prefix)) {
            return max;
          }
          const parsed = Number(String(value).slice(prefix.length));
          return Number.isFinite(parsed) ? Math.max(max, parsed) : max;
        }, 0);

        try {
          await CounterModel.create([{
            id: counterId,
            value: currentMax
          }], { session: session ?? undefined });
        } catch {
          continue;
        }
      }

      const reserved = await CounterModel.findOneAndUpdate(
        { id: counterId },
        { $inc: { value: 1 } },
        {
          new: true,
          session: session ?? undefined
        }
      ).lean();

      const nextValue = Number(reserved?.value ?? 0);
      if (Number.isFinite(nextValue) && nextValue > 0) {
        return `${formatVoucherSeasonPart(season, year)}/${String(nextValue).padStart(2, "0")}`;
      }
    }

    throw new Error("Unable to reserve the next voucher number.");
  }

  private async assertReceiptUnlockedByVoucher(
    cropRegistrationId: string,
    operationLabel: string,
    session?: ClientSession | null
  ) {
    const voucherQuery = FinancialVoucherModel.findOne(
      { cropRegistrationId },
      { voucherNo: 1, status: 1, _id: 0 }
    );
    if (session) {
      voucherQuery.session(session);
    }
    const voucher = await voucherQuery.lean();
    if (voucher) {
      throw new Error(
        `${operationLabel} is blocked because voucher ${voucher.voucherNo} already exists for this registration. Update the voucher only after the intake source is finalized.`
      );
    }
  }

  private async rollbackReceipt(
    receiptRef: string,
    cropRegistrationId?: string,
    session?: ClientSession | null
  ) {
    const receipt = await this.findReceiptByReference(receiptRef, cropRegistrationId, session);
    if (!receipt) {
      return null;
    }

    for (const line of receipt.lines) {
      for (const allocation of line.allocations) {
        const lotQuery = LotModel.findOne({ id: allocation.lotId });
        if (session) {
          lotQuery.session(session);
        }
        const lot = await lotQuery.lean();
        if (!lot) {
          continue;
        }

        const nextQty = roundQtl(lot.currentQtyQtl - allocation.qtyQtl);
        if (nextQty <= 0) {
          await LotModel.updateOne(
            { id: lot.id },
            {
              $set: {
                currentQtyQtl: 0,
                status: "VOID",
                voidReason: `Rolled back from receipt ${receipt.receiptNo}`
              }
            },
            { session: session ?? undefined }
          );
        } else {
          await LotModel.updateOne(
            { id: lot.id },
            {
              $set: {
                currentQtyQtl: nextQty,
                status: nextQty >= lot.maxAllowedQtyQtl ? "FULL" : "OPEN",
                voidReason: ""
              }
            },
            { session: session ?? undefined }
          );
        }
      }
    }

    const totalQty = roundQtl(
      receipt.lines.reduce(
        (sum: number, line: { qtyQtl?: number }) => sum + Number(line.qtyQtl ?? 0),
        0
      )
    );
    const registrationQuery = RegistrationModel.findOne({
      id: receipt.cropRegistrationId
    });
    if (session) {
      registrationQuery.session(session);
    }
    const registration = await registrationQuery.lean<RegistrationRecord | null>();

    if (registration) {
      const totalReceivedQtl = roundQtl(Math.max(registration.totalReceivedQtl - totalQty, 0));
      const balanceQtl = roundQtl(
        Math.max(registration.allowedIntakeQtl - totalReceivedQtl, 0)
      );
      await RegistrationModel.updateOne(
        { id: registration.id },
        {
          $set: {
            totalReceivedQtl,
            balanceQtl,
            status: deriveRegistrationStatus({
              existingStatus: registration.status,
              certifiedAreaHa: registration.certifiedAreaHa,
              expectedYieldQtl: registration.expectedYieldQtl,
              balanceQtl
            })
          }
        },
        { session: session ?? undefined }
      );
    }

    await ReceiptModel.deleteOne({ id: receipt.id }, { session: session ?? undefined });
    await DiscrepancyModel.deleteMany({ receiptNo: receipt.receiptNo }, { session: session ?? undefined });
    await AuditLogModel.create([{
      entityType: "RECEIPT",
      entityId: receipt.id,
      action: "ROLLED_BACK",
      payload: { receiptNo: receipt.receiptNo }
    }], { session: session ?? undefined });

    return receipt;
  }

  private async reconcileDiscrepanciesForRegistration(
    registrationId: string,
    session?: ClientSession | null
  ) {
    if (!config.enableDiscrepancyWorkflow) {
      return;
    }

    const registrationQuery = RegistrationModel.findOne({
      id: registrationId
    });
    if (session) {
      registrationQuery.session(session);
    }
    const registration = await registrationQuery.lean<RegistrationRecord | null>();

    if (!registration) {
      return;
    }

    const receiptsQuery = ReceiptModel.find({
      cropRegistrationId: registrationId
    })
      .sort({ receiptDate: 1, createdAt: 1, receiptNo: 1 });
    if (session) {
      receiptsQuery.session(session);
    }
    const receipts = await receiptsQuery.lean();

    const discrepanciesQuery = DiscrepancyModel.find({
      cropRegistrationId: registrationId
    })
      .sort({ createdAt: 1, discrepancyNo: 1 });
    if (session) {
      discrepanciesQuery.session(session);
    }
    const existingDiscrepancies = await discrepanciesQuery.lean();
    const existingByReceiptNo = new Map<string, (typeof existingDiscrepancies)[number][]>();
    for (const item of existingDiscrepancies) {
      existingByReceiptNo.set(item.receiptNo, [
        ...(existingByReceiptNo.get(item.receiptNo) ?? []),
        item
      ]);
    }
    const discrepancyShiftSums = new Map<string, number>();

    if (existingDiscrepancies.length > 0) {
      const shiftsQuery = DiscrepancyShiftModel.find({
        discrepancyId: { $in: existingDiscrepancies.map((item) => item.id) }
      });
      if (session) {
        shiftsQuery.session(session);
      }
      const shifts = await shiftsQuery.lean();

      for (const shift of shifts) {
        discrepancyShiftSums.set(
          shift.discrepancyId,
          roundQtl((discrepancyShiftSums.get(shift.discrepancyId) ?? 0) + Number(shift.shiftedQtyQtl ?? 0))
        );
      }
    }

    const discrepancyNoQuery = DiscrepancyModel.find({}, { discrepancyNo: 1, _id: 0 });
    if (session) {
      discrepancyNoQuery.session(session);
    }
    const existingDiscrepancyNos = (await discrepancyNoQuery.lean()).map((item) => item.discrepancyNo);

    let runningTotal = 0;
    const seenReceiptNos = new Set<string>();

    for (const receipt of receipts) {
      const receiptNetQtyQtl = sumReceiptNetQty(receipt);
      const previousRunningTotal = runningTotal;
      runningTotal = roundQtl(runningTotal + receiptNetQtyQtl);

      const previousExcess = roundQtl(
        Math.max(previousRunningTotal - registration.allowedIntakeQtl, 0)
      );
      const cumulativeExcess = roundQtl(
        Math.max(runningTotal - registration.allowedIntakeQtl, 0)
      );
      const receiptExcessQtyQtl = roundQtl(Math.max(cumulativeExcess - previousExcess, 0));
      const existingGroup = existingByReceiptNo.get(receipt.receiptNo) ?? [];
      const existing = existingGroup[0];
      const shiftedQtyQtl = roundQtl(
        existing ? discrepancyShiftSums.get(existing.id) ?? 0 : 0
      );
      const pendingExcessQtyQtl = roundQtl(
        Math.max(receiptExcessQtyQtl - shiftedQtyQtl, 0)
      );

      seenReceiptNos.add(receipt.receiptNo);

      if (receiptExcessQtyQtl <= 0) {
        for (const stale of existingGroup) {
          const staleShiftedQtyQtl = roundQtl(discrepancyShiftSums.get(stale.id) ?? 0);
          if (staleShiftedQtyQtl > 0) {
            await DiscrepancyModel.updateOne(
              { id: stale.id },
              {
                $set: {
                  expectedQtyQtl: registration.allowedIntakeQtl,
                  receiptNetQtyQtl,
                  totalReceivedAfterReceiptQtl: runningTotal,
                  excessQtyQtl: 0,
                  estimatedExcessBags: 0,
                  status: "RESOLVED",
                  remarks: stale.remarks || "Auto-corrected: discrepancy no longer applicable."
                }
              },
              { session: session ?? undefined }
            );
          } else {
            await DiscrepancyModel.deleteOne({ id: stale.id }, { session: session ?? undefined });
          }
        }
        continue;
      }

      const firstLine = receipt.lines[0];
      const avgWeightPerBag = averageWeightPerBagKg(receipt);
      const estimatedPendingExcessBags = estimateExcessBags(pendingExcessQtyQtl, avgWeightPerBag);
      const nextStatus =
        pendingExcessQtyQtl <= 0 ? "RESOLVED" : shiftedQtyQtl > 0 ? "SHIFT_PENDING" : "OPEN";

      if (existing) {
        await DiscrepancyModel.updateOne(
          { id: existing.id },
          {
            $set: {
              expectedQtyQtl: registration.allowedIntakeQtl,
              receiptNetQtyQtl,
              totalReceivedAfterReceiptQtl: runningTotal,
              excessQtyQtl: pendingExcessQtyQtl,
              estimatedExcessBags: estimatedPendingExcessBags,
              status: nextStatus,
              remarks:
                existing.remarks ||
                "Receipt exceeded expected yield. Stack marked for later excess shift."
            }
          },
          { session: session ?? undefined }
        );

        for (const duplicate of existingGroup.slice(1)) {
          const duplicateShiftedQtyQtl = roundQtl(discrepancyShiftSums.get(duplicate.id) ?? 0);
          if (duplicateShiftedQtyQtl > 0) {
            await DiscrepancyModel.updateOne(
              { id: duplicate.id },
              {
                $set: {
                  excessQtyQtl: 0,
                  estimatedExcessBags: 0,
                  status: "RESOLVED",
                  remarks: duplicate.remarks || "Auto-corrected duplicate discrepancy."
                }
              },
              { session: session ?? undefined }
            );
          } else {
            await DiscrepancyModel.deleteOne({ id: duplicate.id }, { session: session ?? undefined });
          }
        }
      } else {
        const discrepancyNo = nextDiscrepancyNumber(existingDiscrepancyNos);
        existingDiscrepancyNos.push(discrepancyNo);

        const godownNameQuery = GodownModel.findOne({ id: firstLine.godownId }, { name: 1, _id: 0 });
        if (session) {
          godownNameQuery.session(session);
        }
        const godownName = (await godownNameQuery.lean())?.name ?? "";

        await DiscrepancyModel.create([{
          id: randomUUID(),
          discrepancyNo,
          season: `${registration.season} ${registration.year}`,
          cropRegistrationId: registration.id,
          cropRegistrationCode: registration.cropRegistrationCode,
          farmerName: registration.farmerName,
          receiptNo: receipt.receiptNo,
          receiptDate: receipt.receiptDate,
          godownId: firstLine.godownId,
          godownName,
          stackId: firstLine.stackId,
          stackNo: firstLine.stackNo,
          expectedQtyQtl: registration.allowedIntakeQtl,
          receiptNetQtyQtl,
          totalReceivedAfterReceiptQtl: runningTotal,
          excessQtyQtl: pendingExcessQtyQtl,
          estimatedExcessBags: estimatedPendingExcessBags,
          handlingMode: "MARK_STACK_FOR_SHIFT",
          status: nextStatus,
          remarks:
            firstLine.remarks?.trim() ||
            "Receipt exceeded expected yield. Stack marked for later excess shift."
        }], { session: session ?? undefined });
      }
    }

    for (const discrepancy of existingDiscrepancies) {
      if (seenReceiptNos.has(discrepancy.receiptNo)) {
        continue;
      }

      const shiftedQtyQtl = roundQtl(discrepancyShiftSums.get(discrepancy.id) ?? 0);
      if (shiftedQtyQtl > 0) {
        await DiscrepancyModel.updateOne(
          { id: discrepancy.id },
          {
            $set: {
              excessQtyQtl: 0,
              estimatedExcessBags: 0,
              status: "RESOLVED",
              remarks: discrepancy.remarks || "Auto-corrected after receipt change."
            }
          },
          { session: session ?? undefined }
        );
      } else {
        await DiscrepancyModel.deleteOne({ id: discrepancy.id }, { session: session ?? undefined });
      }
    }
  }

  private async persistReceipt(input: UpdateReceiptInput, session?: ClientSession | null) {
    const registrationQuery = RegistrationModel.findOne({
      id: input.cropRegistrationId
    });
    if (session) {
      registrationQuery.session(session);
    }
    const registration = await registrationQuery.lean<RegistrationRecord | null>();
    if (!registration) {
      throw new Error("Registration not found.");
    }

    if (registration.status !== "ACTIVE") {
      throw new Error("Selected registration is not active for intake.");
    }

    await this.assertReceiptUnlockedByVoucher(
      registration.id,
      input.receiptNo?.trim() ? "Receipt update" : "Receipt save",
      session
    );

    const godownQuery = GodownModel.find();
    const stackQuery = StackModel.find();
    if (session) {
      godownQuery.session(session);
      stackQuery.session(session);
    }
    const godowns = await godownQuery.lean();
    const stacks = await stackQuery.lean();
    const existingDiscrepancyNosQuery = DiscrepancyModel.find({}, { discrepancyNo: 1, _id: 0 });
    if (session) {
      existingDiscrepancyNosQuery.session(session);
    }
    const existingDiscrepancyNos = (await existingDiscrepancyNosQuery.lean()).map(
      (item) => item.discrepancyNo
    );
    const openLotsQuery = LotModel.find({
      cropRegistrationId: registration.id
    });
    if (session) {
      openLotsQuery.session(session);
    }
    let openLots = await openLotsQuery.lean();

    const lines = [];
    let registrationState = registration;

    for (const line of input.lines) {
      const godown = godowns.find((item) => item.id === line.godownId);

      if (!godown) {
        throw new Error("Invalid godown.");
      }

      const stack = await this.ensureStack(godown.id, line.stackNo, session);

        const netWeightQtl = calculateNetWeightQtl(line.grossWeightQtl, line.noOfBags);
        const weightPerBagKg = calculateWeightPerBagKg(netWeightQtl, line.noOfBags);

        const expectedAvailableBalance = roundQtl(registrationState.balanceQtl);
        const excessQtyQtl = roundQtl(Math.max(netWeightQtl - expectedAvailableBalance, 0));

        if (!config.enableDiscrepancyWorkflow && netWeightQtl > expectedAvailableBalance) {
          throw new Error("Net weight cannot exceed expected available balance.");
        }

        if (
          !config.enableDiscrepancyWorkflow &&
          roundQtl(registrationState.totalReceivedQtl + netWeightQtl) > registrationState.allowedIntakeQtl
        ) {
          throw new Error("Receipt quantity exceeds allowed intake for this registration.");
        }

      const matchingLots = openLots
        .filter(
          (lot) =>
            lot.cropRegistrationId === registrationState.id &&
            lot.godownId === godown.id &&
            lot.stackId === stack.id &&
            lot.status === "OPEN"
        )
        .sort((a, b) => a.lotNo - b.lotNo);

      let remainingQty = roundQtl(netWeightQtl);
      const allocations: { lotId: string; lotCode: string; qtyQtl: number }[] = [];
      for (const lot of matchingLots) {
        if (remainingQty <= 0) {
          break;
        }
        const available = roundQtl(lot.maxAllowedQtyQtl - lot.currentQtyQtl);
        if (available <= 0) {
          continue;
        }

        const allocatedQty = roundQtl(Math.min(available, remainingQty));
        const nextQty = roundQtl(lot.currentQtyQtl + allocatedQty);
        remainingQty = roundQtl(remainingQty - allocatedQty);

        const lotUpdate = await LotModel.updateOne(
          { id: lot.id, currentQtyQtl: lot.currentQtyQtl, status: "OPEN" },
          {
            $set: {
              currentQtyQtl: nextQty,
              status: nextQty >= lot.maxAllowedQtyQtl ? "FULL" : "OPEN"
            }
          },
          { session: session ?? undefined }
        );
        if (lotUpdate.matchedCount === 0) {
          throw new Error(
            "Another online entry updated the same lot while this receipt was being saved. Refresh and try again."
          );
        }

        allocations.push({
          lotId: lot.id,
          lotCode: lot.lotCode,
          qtyQtl: allocatedQty
        });
      }

      while (remainingQty > 0) {
        const allocatedQty = roundQtl(Math.min(remainingQty, 200));
        const lotId = randomUUID();
        const reservedLotNo = await this.reserveNextLotNo(registrationState.id, session);
        const lotCode = `${registrationState.year}/${registrationState.cropRegistrationCode}/L${reservedLotNo}`;
        const lot = await LotModel.create([{
          id: lotId,
          cropRegistrationId: registrationState.id,
          cropRegistrationCode: registrationState.cropRegistrationCode,
          lotNo: reservedLotNo,
          lotCode,
          godownId: godown.id,
          godownName: godown.name,
          stackId: stack.id,
          stackNo: stack.stackNo,
          qualityStatus: "ACCEPTED",
          currentQtyQtl: allocatedQty,
          maxAllowedQtyQtl: 200,
          status: allocatedQty >= 200 ? "FULL" : "OPEN",
          createdAt: new Date().toISOString()
        }], { session: session ?? undefined });

        allocations.push({
          lotId: lot[0].id,
          lotCode: lot[0].lotCode,
          qtyQtl: allocatedQty
        });

        openLots = [
          ...openLots,
          {
            ...lot[0].toObject(),
            lotNo: reservedLotNo
          }
        ];
        remainingQty = roundQtl(remainingQty - allocatedQty);
      }

      registrationState = {
        ...registrationState,
        ...nextRegistrationState(registrationState, netWeightQtl)
      };

      lines.push({
        ...line,
        stackId: stack.id,
        stackNo: stack.stackNo,
        netWeightQtl,
        weightPerBagKg,
        qtyQtl: netWeightQtl,
        qualityStatus: "ACCEPTED",
        allocations
      });

      const refreshLotsQuery = LotModel.find({
        cropRegistrationId: registration.id
      });
      if (session) {
        refreshLotsQuery.session(session);
      }
      openLots = await refreshLotsQuery.lean();
    }

    const seasonKey = buildSeasonKey(registrationState.season, registrationState.year);
    const receiptNo = input.receiptNo?.trim()
      ? input.receiptNo.trim()
      : await this.reserveNextReceiptNumber(seasonKey, session);
    const numericReceiptNo = Number(receiptNo);
    const receipt = await ReceiptModel.create([{
      id: randomUUID(),
      seasonKey,
      season: registrationState.season,
      year: registrationState.year,
      receiptNo,
      receiptSequenceNo: Number.isFinite(numericReceiptNo) ? numericReceiptNo : 0,
      receiptDate: input.receiptDate,
      cropRegistrationId: registrationState.id,
      cropRegistrationCode: registrationState.cropRegistrationCode,
      farmerName: registrationState.farmerName,
      lines
    }], { session: session ?? undefined });

    const registrationUpdate = await RegistrationModel.updateOne(
      {
        id: registrationState.id,
        totalReceivedQtl: registration.totalReceivedQtl,
        balanceQtl: registration.balanceQtl,
        status: registration.status
      },
      {
        $set: {
          totalReceivedQtl: registrationState.totalReceivedQtl,
          balanceQtl: registrationState.balanceQtl,
          status: registrationState.status
        }
      },
      { session: session ?? undefined }
    );
    if (registrationUpdate.matchedCount === 0) {
      throw new Error(
        "Another online entry changed this registration while the receipt was being saved. Refresh and try again."
      );
    }

    await AuditLogModel.create([{
      entityType: "RECEIPT",
      entityId: receipt[0].id,
      action: "CREATED",
      payload: receipt[0].toObject()
    }], { session: session ?? undefined });

    await this.reconcileDiscrepanciesForRegistration(registrationState.id, session);

    return this.bootstrap();
  }

  async saveReceipt(input: CreateReceiptInput) {
    await this.ensureCompatibility();
    return this.runInTransaction((session) =>
      this.persistReceipt(
        {
          ...input,
          receiptNo: input.receiptNo?.trim() || ""
        },
        session
      )
    );
  }

  async updateReceipt(receiptRef: string, input: UpdateReceiptInput) {
    return this.runInTransaction(async (session) => {
      await this.assertReceiptUnlockedByVoucher(input.cropRegistrationId, "Receipt update", session);
      await this.rollbackReceipt(receiptRef, input.cropRegistrationId, session);
      return this.persistReceipt(input, session);
    });
  }

  async deleteReceipt(receiptRef: string) {
    return this.runInTransaction(async (session) => {
      const existingReceipt = await this.findReceiptByReference(receiptRef, undefined, session);
      if (!existingReceipt) {
        throw new Error("Receipt not found.");
      }
      await this.assertReceiptUnlockedByVoucher(
        existingReceipt.cropRegistrationId,
        "Receipt delete",
        session
      );
      const rolledBack = await this.rollbackReceipt(receiptRef, undefined, session);
      if (!rolledBack) {
        throw new Error("Receipt not found.");
      }

      await AuditLogModel.create([{
        entityType: "RECEIPT",
        entityId: rolledBack.id,
        action: "DELETED",
        payload: { receiptNo: rolledBack.receiptNo }
      }], { session: session ?? undefined });

      await this.reconcileDiscrepanciesForRegistration(rolledBack.cropRegistrationId, session);

      return this.bootstrap();
    });
  }

  async createDiscrepancyShift(input: CreateDiscrepancyShiftInput) {
    return this.runInTransaction(async (session) => {
      const discrepancyQuery = DiscrepancyModel.findOne({ id: input.discrepancyId });
      if (session) {
        discrepancyQuery.session(session);
      }
      const discrepancy = await discrepancyQuery.lean();
      if (!discrepancy) {
        throw new Error("Discrepancy not found.");
      }

      if (discrepancy.status === "RESOLVED") {
        throw new Error("Discrepancy is already resolved.");
      }

      if (input.shiftedQtyQtl - discrepancy.excessQtyQtl > 0.0001) {
        throw new Error("Shift quantity cannot exceed pending excess quantity.");
      }

      if (input.shiftedBags > discrepancy.estimatedExcessBags && discrepancy.estimatedExcessBags > 0) {
        throw new Error("Shift bags cannot exceed pending excess bags.");
      }

      const godownQuery = GodownModel.findOne({ id: input.toGodownId });
      if (session) {
        godownQuery.session(session);
      }
      const godown = await godownQuery.lean();
      if (!godown) {
        throw new Error("Target godown not found.");
      }

      const stack = await this.ensureStack(godown.id, input.toStackNo, session);
      const nextExcessQtyQtl = roundQtl(Math.max(discrepancy.excessQtyQtl - input.shiftedQtyQtl, 0));
      const nextExcessBags = Math.max(discrepancy.estimatedExcessBags - input.shiftedBags, 0);
      const nextStatus = nextExcessQtyQtl <= 0 ? "RESOLVED" : "SHIFT_PENDING";

      await DiscrepancyShiftModel.create([{
        id: randomUUID(),
        discrepancyId: discrepancy.id,
        discrepancyNo: discrepancy.discrepancyNo,
        cropRegistrationCode: discrepancy.cropRegistrationCode,
        farmerName: discrepancy.farmerName,
        fromGodownName: discrepancy.godownName,
        fromStackNo: discrepancy.stackNo,
        toGodownId: godown.id,
        toGodownName: godown.name,
        toStackId: stack.id,
        toStackNo: stack.stackNo,
        shiftedQtyQtl: roundQtl(input.shiftedQtyQtl),
        shiftedBags: input.shiftedBags,
        shiftDate: input.shiftDate,
        approvedBy: input.approvedBy?.trim() || "",
        remarks: input.remarks?.trim() || ""
      }], { session: session ?? undefined });

      await NonCertificationStockMovementModel.create([{
        id: randomUUID(),
        movementType: "DISCREPANCY_SHIFT_OUT",
        referenceType: "DISCREPANCY",
        referenceId: discrepancy.id,
        cropRegistrationCode: discrepancy.cropRegistrationCode,
        farmerName: discrepancy.farmerName,
        fromGodownId: discrepancy.godownId,
        fromGodownName: discrepancy.godownName,
        fromStackId: discrepancy.stackId,
        fromStackNo: discrepancy.stackNo,
        toGodownId: godown.id,
        toGodownName: godown.name,
        toStackId: stack.id,
        toStackNo: stack.stackNo,
        qtyQtl: roundQtl(input.shiftedQtyQtl),
        bags: input.shiftedBags,
        movementDate: input.shiftDate,
        remarks: input.remarks?.trim() || ""
      }], { session: session ?? undefined });

      await DiscrepancyModel.updateOne(
        { id: discrepancy.id },
        {
          $set: {
            excessQtyQtl: nextExcessQtyQtl,
            estimatedExcessBags: nextExcessBags,
            status: nextStatus,
            remarks: appendShiftRemark(
              String(discrepancy.remarks ?? ""),
              input.remarks ?? "",
              roundQtl(input.shiftedQtyQtl),
              input.shiftedBags
            )
          }
        },
        { session: session ?? undefined }
      );

      await AuditLogModel.create([{
        entityType: "DISCREPANCY_SHIFT",
        entityId: discrepancy.id,
        action: "SHIFT_CREATED",
        payload: {
          discrepancyId: discrepancy.id,
          shiftedQtyQtl: roundQtl(input.shiftedQtyQtl),
          shiftedBags: input.shiftedBags,
          shiftDate: input.shiftDate,
          toGodownId: godown.id,
          toStackNo: stack.stackNo
        }
      }], { session: session ?? undefined });

      return this.bootstrap();
    });
  }

  async createStackAccommodation(input: CreateStackAccommodationInput, actor?: ActorUser) {
    await this.ensureCompatibility();

    const discrepancy = await DiscrepancyModel.findOne({ id: input.discrepancyId }).lean();
    if (!discrepancy) {
      throw new Error("Discrepancy record not found.");
    }
    if (discrepancy.status === "RESOLVED") {
      throw new Error("Resolved discrepancy cannot be accommodated.");
    }

    const targetRegistration = await RegistrationModel.findOne({ id: input.targetRegistrationId }).lean();
    if (!targetRegistration) {
      throw new Error("Target registration not found.");
    }
    if (targetRegistration.id === discrepancy.cropRegistrationId) {
      throw new Error("Source and target farmer cannot be the same.");
    }

    const targetReceipts = await ReceiptModel.find({
      cropRegistrationId: targetRegistration.id,
      "lines.godownId": discrepancy.godownId,
      "lines.stackNo": discrepancy.stackNo
    }).lean();
    if (targetReceipts.length === 0) {
      throw new Error("Target farmer has no intake record in the same source stack.");
    }

    const existingAccommodations = await StackAccommodationModel.find({
      discrepancyId: discrepancy.id
    }).lean();
    const alreadyMappedQty = roundQtl(
      existingAccommodations.reduce((sum, item) => sum + Number(item.adjustedQtyQtl ?? 0), 0)
    );
    const alreadyMappedBags = existingAccommodations.reduce(
      (sum, item) => sum + Number(item.adjustedBags ?? 0),
      0
    );
    const remainingQty = roundQtl(Number(discrepancy.excessQtyQtl ?? 0) - alreadyMappedQty);
    const remainingBags = Math.max(Number(discrepancy.estimatedExcessBags ?? 0) - alreadyMappedBags, 0);

    if (input.adjustedQtyQtl > remainingQty) {
      throw new Error(`Accommodation qty exceeds remaining discrepancy qty of ${remainingQty.toFixed(2)} QTL.`);
    }
    if (input.adjustedBags > remainingBags) {
      throw new Error(`Accommodation bags exceed remaining discrepancy bags of ${remainingBags}.`);
    }

    await StackAccommodationModel.create({
      id: randomUUID(),
      discrepancyId: discrepancy.id,
      discrepancyNo: discrepancy.discrepancyNo,
      sourceRegistrationId: discrepancy.cropRegistrationId,
      sourceRegistrationCode: discrepancy.cropRegistrationCode,
      sourceFarmerName: discrepancy.farmerName,
      targetRegistrationId: targetRegistration.id,
      targetRegistrationCode: targetRegistration.cropRegistrationCode,
      targetFarmerName: targetRegistration.farmerName,
      godownId: discrepancy.godownId,
      godownName: discrepancy.godownName,
      stackId: discrepancy.stackId,
      stackNo: discrepancy.stackNo,
      adjustedQtyQtl: roundQtl(input.adjustedQtyQtl),
      adjustedBags: Number(input.adjustedBags ?? 0),
      adjustmentDate: input.adjustmentDate,
      remarks: input.remarks?.trim() || "",
      createdBy: actor?.email ?? ""
    });

    await AuditLogModel.create({
      entityType: "STACK_ACCOMMODATION",
      entityId: discrepancy.id,
      action: "CREATED",
      payload: {
        discrepancyId: discrepancy.id,
        discrepancyNo: discrepancy.discrepancyNo,
        sourceRegistrationCode: discrepancy.cropRegistrationCode,
        targetRegistrationCode: targetRegistration.cropRegistrationCode,
        adjustedQtyQtl: roundQtl(input.adjustedQtyQtl),
        adjustedBags: Number(input.adjustedBags ?? 0),
        adjustmentDate: input.adjustmentDate,
        remarks: input.remarks?.trim() || "",
        createdBy: actor?.email ?? ""
      }
    });

    return this.bootstrap();
  }

  async updateStackAccommodation(
    accommodationId: string,
    input: UpdateStackAccommodationInput,
    actor: ActorUser
  ) {
    await this.ensureCompatibility();

    const accommodation = await StackAccommodationModel.findOne({ id: accommodationId }).lean();
    if (!accommodation) {
      throw new Error("Accommodation entry not found.");
    }
    await this.assertStackAccommodationEditable(accommodation, actor, input.adminPassword);

    const discrepancy = await DiscrepancyModel.findOne({ id: input.discrepancyId }).lean();
    if (!discrepancy) {
      throw new Error("Discrepancy record not found.");
    }
    if (discrepancy.status === "RESOLVED") {
      throw new Error("Resolved discrepancy cannot be accommodated.");
    }

    const targetRegistration = await RegistrationModel.findOne({ id: input.targetRegistrationId }).lean();
    if (!targetRegistration) {
      throw new Error("Target registration not found.");
    }
    if (targetRegistration.id === discrepancy.cropRegistrationId) {
      throw new Error("Source and target farmer cannot be the same.");
    }

    const targetReceipts = await ReceiptModel.find({
      cropRegistrationId: targetRegistration.id,
      "lines.godownId": discrepancy.godownId,
      "lines.stackNo": discrepancy.stackNo
    }).lean();
    if (targetReceipts.length === 0) {
      throw new Error("Target farmer has no intake record in the same source stack.");
    }

    const existingAccommodations = await StackAccommodationModel.find({
      discrepancyId: discrepancy.id,
      id: { $ne: accommodationId }
    }).lean();
    const alreadyMappedQty = roundQtl(
      existingAccommodations.reduce((sum, item) => sum + Number(item.adjustedQtyQtl ?? 0), 0)
    );
    const alreadyMappedBags = existingAccommodations.reduce(
      (sum, item) => sum + Number(item.adjustedBags ?? 0),
      0
    );
    const remainingQty = roundQtl(Number(discrepancy.excessQtyQtl ?? 0) - alreadyMappedQty);
    const remainingBags = Math.max(Number(discrepancy.estimatedExcessBags ?? 0) - alreadyMappedBags, 0);

    if (input.adjustedQtyQtl > remainingQty) {
      throw new Error(`Accommodation qty exceeds remaining discrepancy qty of ${remainingQty.toFixed(2)} QTL.`);
    }
    if (input.adjustedBags > remainingBags) {
      throw new Error(`Accommodation bags exceed remaining discrepancy bags of ${remainingBags}.`);
    }

    await StackAccommodationModel.updateOne(
      { id: accommodationId },
      {
        $set: {
          discrepancyId: discrepancy.id,
          discrepancyNo: discrepancy.discrepancyNo,
          sourceRegistrationId: discrepancy.cropRegistrationId,
          sourceRegistrationCode: discrepancy.cropRegistrationCode,
          sourceFarmerName: discrepancy.farmerName,
          targetRegistrationId: targetRegistration.id,
          targetRegistrationCode: targetRegistration.cropRegistrationCode,
          targetFarmerName: targetRegistration.farmerName,
          godownId: discrepancy.godownId,
          godownName: discrepancy.godownName,
          stackId: discrepancy.stackId,
          stackNo: discrepancy.stackNo,
          adjustedQtyQtl: roundQtl(input.adjustedQtyQtl),
          adjustedBags: Number(input.adjustedBags ?? 0),
          adjustmentDate: input.adjustmentDate,
          remarks: input.remarks?.trim() || ""
        }
      }
    );

    await AuditLogModel.create({
      entityType: "STACK_ACCOMMODATION",
      entityId: accommodationId,
      action: "UPDATED",
      payload: {
        discrepancyId: discrepancy.id,
        discrepancyNo: discrepancy.discrepancyNo,
        sourceRegistrationCode: discrepancy.cropRegistrationCode,
        targetRegistrationCode: targetRegistration.cropRegistrationCode,
        adjustedQtyQtl: roundQtl(input.adjustedQtyQtl),
        adjustedBags: Number(input.adjustedBags ?? 0),
        adjustmentDate: input.adjustmentDate,
        remarks: input.remarks?.trim() || "",
        updatedBy: actor.email
      }
    });

    return this.bootstrap();
  }

  async deleteStackAccommodation(accommodationId: string, actor: ActorUser) {
    await this.ensureCompatibility();
    const accommodation = await StackAccommodationModel.findOne({ id: accommodationId }).lean();
    if (!accommodation) {
      throw new Error("Accommodation entry not found.");
    }
    await this.assertStackAccommodationEditable(accommodation, actor);

    await StackAccommodationModel.deleteOne({ id: accommodationId });
    await AuditLogModel.create({
      entityType: "STACK_ACCOMMODATION",
      entityId: accommodationId,
      action: "DELETED",
      payload: {
        discrepancyId: accommodation.discrepancyId,
        discrepancyNo: accommodation.discrepancyNo,
        sourceRegistrationCode: accommodation.sourceRegistrationCode,
        targetRegistrationCode: accommodation.targetRegistrationCode,
        adjustedQtyQtl: roundQtl(Number(accommodation.adjustedQtyQtl ?? 0)),
        adjustedBags: Number(accommodation.adjustedBags ?? 0),
        deletedBy: actor.email
      }
    });

    return this.bootstrap();
  }

  private async buildFinancialVoucherPreview(
    input: CreateFinancialVoucherInput | UpdateFinancialVoucherInput,
    existingVoucher?: FinancialVoucherRecord | null
  ): Promise<FinancialVoucherPreview> {
    const registration = await RegistrationModel.findOne({
      id: input.cropRegistrationId
    }).lean<RegistrationRecord | null>();
    if (!registration) {
      throw new Error("Registration not found.");
    }

    const receipts = await ReceiptModel.find({
      cropRegistrationId: registration.id
    })
      .sort({ receiptDate: 1, createdAt: 1, receiptNo: 1 })
      .lean();
    if (receipts.length === 0) {
      throw new Error("No intake receipts found for this registration.");
    }

    const discrepancies = await DiscrepancyModel.find({
      cropRegistrationId: registration.id
    }).lean();

    const totalGrossQtyQtl = roundQtl(
      receipts.reduce(
        (sum, receipt) =>
          sum +
          receipt.lines.reduce(
            (lineSum: number, line: { grossWeightQtl?: number }) =>
              lineSum + Number(line.grossWeightQtl ?? 0),
            0
          ),
        0
      )
    );
    const totalNetQtyQtl = roundQtl(
      receipts.reduce((sum, receipt) => sum + sumReceiptNetQty(receipt), 0)
    );
    const totalBags = receipts.reduce(
      (sum, receipt) =>
        sum +
        receipt.lines.reduce(
          (lineSum: number, line: { noOfBags?: number }) =>
            lineSum + Number(line.noOfBags ?? 0),
          0
        ),
      0
    );
    const discrepancyQtyQtl = roundQtl(
      discrepancies.reduce((sum, item) => sum + Number(item.excessQtyQtl ?? 0), 0)
    );
    const discrepancyBags = discrepancies.reduce(
      (sum, item) => sum + Number(item.estimatedExcessBags ?? 0),
      0
    );
    const existingPayments = (existingVoucher?.payments ?? []).map((payment) => ({
      id: String(payment.id),
      paymentDate: String(payment.paymentDate),
      amount: roundQtl(Number(payment.amount ?? 0)),
      transactionNo: String(payment.transactionNo ?? ""),
      mode: String(payment.mode ?? "RTGS/NEFT"),
      remarks: String(payment.remarks ?? ""),
      createdBy: String(payment.createdBy ?? ""),
      updatedBy: String(payment.updatedBy ?? "")
    }));
    const certifiedQtyQtl = roundQtl(Math.max(totalNetQtyQtl - discrepancyQtyQtl, 0));
    const certifiedAmount = roundQtl(certifiedQtyQtl * input.certifiedRatePerQtl);
    const discrepancyAmount = roundQtl(discrepancyQtyQtl * input.discrepancyRatePerQtl);
    const grossPayableAmount = roundQtl(certifiedAmount + discrepancyAmount);
    const netPayableAmount = roundQtl(grossPayableAmount - Number(input.deductionAmount ?? 0));
    const finalPayableAmount = Math.round(netPayableAmount);
    const roundedOffAmount = roundQtl(finalPayableAmount - netPayableAmount);
    const paymentSummary = this.buildVoucherPaymentSummary(existingPayments, finalPayableAmount);
    return {
      id: existingVoucher?.id ?? randomUUID(),
      voucherNo:
        existingVoucher?.voucherNo ??
        await this.reserveNextVoucherNumber(registration.season, registration.year),
      voucherDate: input.voucherDate,
      cropRegistrationId: registration.id,
      cropRegistrationCode: registration.cropRegistrationCode,
      farmerName: registration.farmerName,
      fatherName: registration.fatherName,
      village: registration.village,
      block: registration.block,
      district: registration.district,
      season: registration.season,
      year: registration.year,
      crop: registration.crop,
      variety: registration.variety,
      classStage: registration.classStage,
      expectedYieldQtl: roundQtl(Number(registration.expectedYieldQtl ?? 0)),
      totalGrossQtyQtl,
      totalNetQtyQtl,
      totalBags,
      certifiedQtyQtl,
      discrepancyQtyQtl,
      discrepancyBags,
      certifiedRatePerQtl: roundQtl(input.certifiedRatePerQtl),
      discrepancyRatePerQtl: roundQtl(input.discrepancyRatePerQtl),
      certifiedAmount,
      discrepancyAmount,
      grossPayableAmount,
      deductionAmount: roundQtl(Number(input.deductionAmount ?? 0)),
      netPayableAmount,
      roundedOffAmount,
      finalPayableAmount,
      totalPaidAmount: paymentSummary.totalPaidAmount,
      balanceAmount: paymentSummary.balanceAmount,
      lastPaymentDate: paymentSummary.lastPaymentDate,
      status: paymentSummary.status,
      remarks: input.remarks?.trim() || "",
      lines: receipts.map((receipt) => ({
        receiptId: String(receipt.id),
        receiptNo: String(receipt.receiptNo),
        receiptDate: String(receipt.receiptDate),
        vehicleNo:
          Array.from(
            new Set(
              receipt.lines
                .map((line: { vehicleNo?: string }) => String(line.vehicleNo ?? "").trim())
                .filter(Boolean)
            )
          ).join(", ") || "-",
        stackNo:
          Array.from(
            new Set(
              receipt.lines
                .map((line: { stackNo?: string }) => String(line.stackNo ?? "").trim())
                .filter(Boolean)
            )
          ).join(", ") || "-",
        bags: receipt.lines.reduce(
          (sum: number, line: { noOfBags?: number }) => sum + Number(line.noOfBags ?? 0),
          0
        ),
        grossQtyQtl: roundQtl(
          receipt.lines.reduce(
            (sum: number, line: { grossWeightQtl?: number }) =>
              sum + Number(line.grossWeightQtl ?? 0),
            0
          )
        ),
        netQtyQtl: sumReceiptNetQty(receipt)
      })),
      payments: existingPayments
    };
  }

  async createFinancialVoucher(input: CreateFinancialVoucherInput) {
    await this.ensureCompatibility();
    const existingVoucher = await FinancialVoucherModel.findOne({
      cropRegistrationId: input.cropRegistrationId
    }).lean();
    if (existingVoucher) {
      throw new Error(
        `Voucher ${existingVoucher.voucherNo} already exists for this registration. Open the existing voucher and edit it instead of creating a duplicate.`
      );
    }
    const preview = await this.buildFinancialVoucherPreview(input);

    await FinancialVoucherModel.create({
      ...preview
    });

    await AuditLogModel.create({
      entityType: "FINANCIAL_VOUCHER",
      entityId: preview.id,
      action: "CREATED",
      payload: {
        voucherNo: preview.voucherNo,
        cropRegistrationCode: preview.cropRegistrationCode,
        grossPayableAmount: preview.grossPayableAmount,
        netPayableAmount: preview.netPayableAmount
      }
    });

    return this.bootstrap();
  }

  async updateFinancialVoucher(voucherId: string, input: UpdateFinancialVoucherInput) {
    await this.ensureCompatibility();
    const existingVoucher = await FinancialVoucherModel.findOne({ id: voucherId }).lean();
    if (!existingVoucher) {
      throw new Error("Financial voucher not found.");
    }
    if (isVoucherLockedStatus(String(existingVoucher.status ?? "DRAFT"))) {
      await this.verifyAdminPassword(input.adminPassword);
    }
    if (existingVoucher.cropRegistrationId !== input.cropRegistrationId) {
      throw new Error("Voucher registration cannot be changed.");
    }

    const preview = await this.buildFinancialVoucherPreview(
      input,
      existingVoucher as FinancialVoucherRecord
    );

    await FinancialVoucherModel.updateOne(
      { id: voucherId },
      {
        $set: {
          ...preview
        }
      }
    );

    await AuditLogModel.create({
      entityType: "FINANCIAL_VOUCHER",
      entityId: preview.id,
      action: "UPDATED",
      payload: {
        voucherNo: preview.voucherNo,
        cropRegistrationCode: preview.cropRegistrationCode,
        grossPayableAmount: preview.grossPayableAmount,
        netPayableAmount: preview.netPayableAmount
      }
    });

    return this.bootstrap();
  }

  async markFinancialVoucherPaid(voucherId: string) {
    await this.ensureCompatibility();
    const existingVoucher = await FinancialVoucherModel.findOne({ id: voucherId }).lean();
    if (!existingVoucher) {
      throw new Error("Financial voucher not found.");
    }
    if (isVoucherLockedStatus(String(existingVoucher.status ?? "DRAFT"))) {
      throw new Error("Voucher is already marked as paid.");
    }

    const targetFinalPayable = roundQtl(
      Number(existingVoucher.finalPayableAmount ?? existingVoucher.netPayableAmount ?? 0)
    );
    const paymentSummary = this.buildVoucherPaymentSummary(
      [
        ...((existingVoucher.payments ?? []).map((payment: { paymentDate?: unknown; amount?: unknown }) => ({
          paymentDate: String(payment.paymentDate ?? ""),
          amount: roundQtl(Number(payment.amount ?? 0))
        }))),
        {
          paymentDate: new Date().toISOString().slice(0, 10),
          amount: targetFinalPayable
        }
      ],
      targetFinalPayable
    );

    await FinancialVoucherModel.updateOne(
      { id: voucherId },
      {
        $set: {
          status: paymentSummary.status
        }
      }
    );

    await AuditLogModel.create({
      entityType: "FINANCIAL_VOUCHER",
      entityId: voucherId,
      action: "MARKED_PAID",
      payload: {
        voucherNo: existingVoucher.voucherNo,
        cropRegistrationCode: existingVoucher.cropRegistrationCode
      }
    });

    return this.bootstrap();
  }

  async addFinancialVoucherPayment(
    voucherId: string,
    input: AddFinancialVoucherPaymentInput,
    actor: ActorUser
  ) {
    await this.ensureCompatibility();
    await this.runInTransaction(async (session) => {
      const voucherQuery = FinancialVoucherModel.findOne({ id: voucherId });
      if (session) {
        voucherQuery.session(session);
      }
      const existingVoucher = await voucherQuery.lean();
      if (!existingVoucher) {
        throw new Error("Financial voucher not found.");
      }

      const normalizedTransactionNo = input.transactionNo.trim().toUpperCase();
      if (!normalizedTransactionNo) {
        throw new Error("Transaction number is required.");
      }

      const duplicateInVoucher = (existingVoucher.payments ?? []).some(
        (payment: { transactionNo?: unknown }) =>
          String(payment.transactionNo ?? "").trim().toUpperCase() === normalizedTransactionNo
      );
      if (duplicateInVoucher) {
        throw new Error("This transaction number is already recorded in the voucher ledger.");
      }

      const duplicateAcrossVouchersQuery = FinancialVoucherModel.findOne({
        id: { $ne: voucherId },
        "payments.transactionNo": normalizedTransactionNo
      });
      if (session) {
        duplicateAcrossVouchersQuery.session(session);
      }
      const duplicateAcrossVouchers = await duplicateAcrossVouchersQuery.lean();
      if (duplicateAcrossVouchers) {
        throw new Error(
          `Transaction number already exists in voucher ${duplicateAcrossVouchers.voucherNo}.`
        );
      }

      const payment = {
        id: randomUUID(),
        paymentDate: input.paymentDate,
        amount: roundQtl(Number(input.amount ?? 0)),
        transactionNo: normalizedTransactionNo,
        mode: "RTGS/NEFT",
        remarks: input.remarks?.trim() || "",
        createdBy: actor.email,
        updatedBy: actor.email
      };

      const payments = [...(existingVoucher.payments ?? []), payment];
      const finalPayableAmount = roundQtl(
        Number(existingVoucher.finalPayableAmount ?? existingVoucher.netPayableAmount ?? 0)
      );
      const paymentSummary = this.buildVoucherPaymentSummary(
        payments.map((item) => ({
          paymentDate: String(item.paymentDate ?? ""),
          amount: roundQtl(Number(item.amount ?? 0))
        })),
        finalPayableAmount
      );

      const voucherUpdate = await FinancialVoucherModel.updateOne(
        {
          id: voucherId,
          totalPaidAmount: roundQtl(Number(existingVoucher.totalPaidAmount ?? 0)),
          balanceAmount: roundQtl(Number(existingVoucher.balanceAmount ?? 0)),
          status: String(existingVoucher.status ?? "DRAFT")
        },
        {
          $set: {
            payments,
            totalPaidAmount: paymentSummary.totalPaidAmount,
            balanceAmount: paymentSummary.balanceAmount,
            lastPaymentDate: paymentSummary.lastPaymentDate,
            status: paymentSummary.status
          }
        },
        { session: session ?? undefined }
      );
      if (voucherUpdate.matchedCount === 0) {
        throw new Error(
          "Another online change updated this voucher while payment was being recorded. Refresh and try again."
        );
      }

      await AuditLogModel.create([{
        entityType: "FINANCIAL_VOUCHER",
        entityId: voucherId,
        action: "PAYMENT_ADDED",
        payload: {
          voucherNo: existingVoucher.voucherNo,
          paymentDate: payment.paymentDate,
          amount: payment.amount,
          transactionNo: payment.transactionNo
        }
      }], { session: session ?? undefined });
    });
    return this.bootstrap();
  }

  async updateFinancialVoucherPayment(
    voucherId: string,
    paymentId: string,
    input: UpdateFinancialVoucherPaymentInput,
    actor: ActorUser
  ) {
    await this.ensureCompatibility();
    await this.runInTransaction(async (session) => {
      const voucherQuery = FinancialVoucherModel.findOne({ id: voucherId });
      if (session) {
        voucherQuery.session(session);
      }
      const existingVoucher = await voucherQuery.lean();
      if (!existingVoucher) {
        throw new Error("Financial voucher not found.");
      }

      const existingPayments: FinancialVoucherRecord["payments"] = (existingVoucher.payments ?? []).map((payment: {
        id?: unknown;
        paymentDate?: unknown;
        amount?: unknown;
        transactionNo?: unknown;
        mode?: unknown;
        remarks?: unknown;
        createdBy?: unknown;
        updatedBy?: unknown;
      }) => ({
        id: String(payment.id ?? ""),
        paymentDate: String(payment.paymentDate ?? ""),
        amount: roundQtl(Number(payment.amount ?? 0)),
        transactionNo: String(payment.transactionNo ?? "").trim().toUpperCase(),
        mode: String(payment.mode ?? "RTGS/NEFT"),
        remarks: String(payment.remarks ?? ""),
        createdBy: String(payment.createdBy ?? ""),
        updatedBy: String(payment.updatedBy ?? "")
      }));

      const targetPayment = existingPayments.find((payment) => payment.id === paymentId);
      if (!targetPayment) {
        throw new Error("Voucher payment entry not found.");
      }

      await this.assertVoucherPaymentEditable(existingVoucher, targetPayment, actor, input.adminPassword);

      const normalizedTransactionNo = input.transactionNo.trim().toUpperCase();
      if (!normalizedTransactionNo) {
        throw new Error("Transaction number is required.");
      }

      const duplicateInVoucher = existingPayments.some(
        (payment) => payment.id !== paymentId && payment.transactionNo === normalizedTransactionNo
      );
      if (duplicateInVoucher) {
        throw new Error("This transaction number is already recorded in the voucher ledger.");
      }

      const duplicateAcrossVouchersQuery = FinancialVoucherModel.findOne({
        id: { $ne: voucherId },
        "payments.transactionNo": normalizedTransactionNo
      });
      if (session) {
        duplicateAcrossVouchersQuery.session(session);
      }
      const duplicateAcrossVouchers = await duplicateAcrossVouchersQuery.lean();
      if (duplicateAcrossVouchers) {
        throw new Error(
          `Transaction number already exists in voucher ${duplicateAcrossVouchers.voucherNo}.`
        );
      }

      const payments = existingPayments.map((payment) =>
        payment.id === paymentId
          ? {
              ...payment,
              paymentDate: input.paymentDate,
              amount: roundQtl(Number(input.amount ?? 0)),
              transactionNo: normalizedTransactionNo,
              mode: "RTGS/NEFT",
              remarks: input.remarks?.trim() || "",
              updatedBy: actor.email
            }
          : payment
      );

      const finalPayableAmount = roundQtl(
        Number(existingVoucher.finalPayableAmount ?? existingVoucher.netPayableAmount ?? 0)
      );
      const paymentSummary = this.buildVoucherPaymentSummary(
        payments.map((item) => ({
          paymentDate: String(item.paymentDate ?? ""),
          amount: roundQtl(Number(item.amount ?? 0))
        })),
        finalPayableAmount
      );

      const voucherUpdate = await FinancialVoucherModel.updateOne(
        {
          id: voucherId,
          totalPaidAmount: roundQtl(Number(existingVoucher.totalPaidAmount ?? 0)),
          balanceAmount: roundQtl(Number(existingVoucher.balanceAmount ?? 0)),
          status: String(existingVoucher.status ?? "DRAFT")
        },
        {
          $set: {
            payments,
            totalPaidAmount: paymentSummary.totalPaidAmount,
            balanceAmount: paymentSummary.balanceAmount,
            lastPaymentDate: paymentSummary.lastPaymentDate,
            status: paymentSummary.status
          }
        },
        { session: session ?? undefined }
      );
      if (voucherUpdate.matchedCount === 0) {
        throw new Error(
          "Another online change updated this voucher while payment was being edited. Refresh and try again."
        );
      }

      await AuditLogModel.create([{
        entityType: "FINANCIAL_VOUCHER",
        entityId: voucherId,
        action: "PAYMENT_UPDATED",
        payload: {
          voucherNo: existingVoucher.voucherNo,
          paymentId,
          paymentDate: input.paymentDate,
          amount: roundQtl(Number(input.amount ?? 0)),
          transactionNo: normalizedTransactionNo,
          updatedBy: actor.email
        }
      }], { session: session ?? undefined });
    });
    return this.bootstrap();
  }

  async deleteFinancialVoucher(voucherId: string, input: FinancialVoucherActionInput) {
    await this.ensureCompatibility();
    const existingVoucher = await FinancialVoucherModel.findOne({ id: voucherId }).lean();
    if (!existingVoucher) {
      throw new Error("Financial voucher not found.");
    }
    if (isVoucherLockedStatus(String(existingVoucher.status ?? "DRAFT"))) {
      await this.verifyAdminPassword(input.adminPassword);
    }

    await FinancialVoucherModel.deleteOne({ id: voucherId });

    await AuditLogModel.create({
      entityType: "FINANCIAL_VOUCHER",
      entityId: voucherId,
      action: "DELETED",
      payload: {
        voucherNo: existingVoucher.voucherNo,
        cropRegistrationCode: existingVoucher.cropRegistrationCode,
        previousStatus: existingVoucher.status
      }
    });

    return this.bootstrap();
  }

  async addOrganizerPayment(input: AddOrganizerPaymentInput, actor: ActorUser) {
    await this.ensureCompatibility();
    const organizer = await OrganizerModel.findOne({ id: input.organizerId }).lean<OrganizerRecord | null>();
    if (!organizer) {
      throw new Error("Organizer not found.");
    }

    const transactionNo = input.transactionNo.trim().toUpperCase();
    if (!transactionNo) {
      throw new Error("Transaction number is required.");
    }

    const duplicatePayment = await OrganizerPaymentModel.findOne({ transactionNo }).lean();
    if (duplicatePayment) {
      throw new Error(`Transaction number ${transactionNo} already exists in organizer commission payments.`);
    }

    const commissionPosition = await this.getOrganizerCommissionPosition(organizer.id);
    if (commissionPosition.balanceAmount <= 0) {
      throw new Error("No organizer commission balance is pending for payment.");
    }
    const requestedAmount = roundQtl(Number(input.amount ?? 0));
    if (requestedAmount > commissionPosition.balanceAmount) {
      throw new Error(
        `Commission payment exceeds pending balance. Pending balance is ${commissionPosition.balanceAmount.toFixed(2)}.`
      );
    }

    const payment = await OrganizerPaymentModel.create({
      id: randomUUID(),
      organizerId: organizer.id,
      organizerName: organizer.name,
      paymentDate: input.paymentDate,
      amount: requestedAmount,
      transactionNo,
      remarks: input.remarks?.trim() ?? "",
      createdBy: actor.email,
      updatedBy: actor.email
    });

    await AuditLogModel.create({
      entityType: "ORGANIZER_COMMISSION_PAYMENT",
      entityId: payment.id,
      action: "CREATED",
      payload: payment.toObject()
    });

    return payment.toObject();
  }

  async updateOrganizerPayment(paymentId: string, input: UpdateOrganizerPaymentInput, actor: ActorUser) {
    await this.ensureCompatibility();
    const existingPayment = await OrganizerPaymentModel.findOne({ id: paymentId }).lean<OrganizerPaymentRecord | null>();
    if (!existingPayment) {
      throw new Error("Organizer commission payment not found.");
    }

    this.assertOrganizerPaymentEditable(existingPayment, actor);

    const transactionNo = input.transactionNo.trim().toUpperCase();
    if (!transactionNo) {
      throw new Error("Transaction number is required.");
    }

    const duplicatePayment = await OrganizerPaymentModel.findOne({
      id: { $ne: paymentId },
      transactionNo
    }).lean();
    if (duplicatePayment) {
      throw new Error(`Transaction number ${transactionNo} already exists in organizer commission payments.`);
    }

    const commissionPosition = await this.getOrganizerCommissionPosition(existingPayment.organizerId, paymentId);
    const requestedAmount = roundQtl(Number(input.amount ?? 0));
    if (requestedAmount > commissionPosition.balanceAmount) {
      throw new Error(
        `Commission payment exceeds pending balance. Pending balance is ${commissionPosition.balanceAmount.toFixed(2)}.`
      );
    }

    const nextPayment = {
      paymentDate: input.paymentDate,
      amount: requestedAmount,
      transactionNo,
      remarks: input.remarks?.trim() ?? "",
      updatedBy: actor.email
    };

    await OrganizerPaymentModel.updateOne(
      { id: paymentId },
      {
        $set: nextPayment
      }
    );

    await AuditLogModel.create({
      entityType: "ORGANIZER_COMMISSION_PAYMENT",
      entityId: paymentId,
      action: "UPDATED",
      payload: {
        before: existingPayment,
        after: {
          ...existingPayment,
          ...nextPayment
        }
      }
    });

    return {
      ...existingPayment,
      ...nextPayment
    };
  }

  async deleteOrganizerPayment(paymentId: string, actor: ActorUser) {
    await this.ensureCompatibility();
    const existingPayment = await OrganizerPaymentModel.findOne({ id: paymentId }).lean<OrganizerPaymentRecord | null>();
    if (!existingPayment) {
      throw new Error("Organizer commission payment not found.");
    }

    this.assertOrganizerPaymentEditable(existingPayment, actor);

    await OrganizerPaymentModel.deleteOne({ id: paymentId });

    await AuditLogModel.create({
      entityType: "ORGANIZER_COMMISSION_PAYMENT",
      entityId: paymentId,
      action: "DELETED",
      payload: existingPayment
    });

    return { success: true };
  }

  private async buildBaseReceiptRows(filters: ReportFilterInput) {
    const [registrations, receipts, lots, godowns, discrepancies, discrepancyShifts] =
      await Promise.all([
        RegistrationModel.find().lean<RegistrationRecord[]>(),
        ReceiptModel.find().lean(),
        LotModel.find().lean(),
        GodownModel.find().lean(),
        DiscrepancyModel.find().lean(),
        DiscrepancyShiftModel.find().lean()
      ]);

    const registrationById = new Map(registrations.map((item) => [item.id, item]));
    const lotCodesById = new Map(lots.map((lot) => [lot.id, lot.lotCode]));
    const godownById = new Map(godowns.map((item) => [item.id, item.name]));
    const shiftQtyByDiscrepancyId = new Map<string, number>();

    for (const shift of discrepancyShifts) {
      shiftQtyByDiscrepancyId.set(
        shift.discrepancyId,
        roundQtl(
          (shiftQtyByDiscrepancyId.get(shift.discrepancyId) ?? 0) + Number(shift.shiftedQtyQtl ?? 0)
        )
      );
    }

    const discrepancyByReceiptNo = new Map<
      string,
      {
        pendingQtyQtl: number;
        shiftedQtyQtl: number;
        status: string;
      }
    >();

    for (const discrepancy of discrepancies) {
      const shiftedQtyQtl = roundQtl(shiftQtyByDiscrepancyId.get(discrepancy.id) ?? 0);
      const pendingQtyQtl = roundQtl(Math.max(Number(discrepancy.excessQtyQtl ?? 0), 0));
      const existing = discrepancyByReceiptNo.get(discrepancy.receiptNo) ?? {
        pendingQtyQtl: 0,
        shiftedQtyQtl: 0,
        status: "RESOLVED"
      };

      discrepancyByReceiptNo.set(discrepancy.receiptNo, {
        pendingQtyQtl: roundQtl(existing.pendingQtyQtl + pendingQtyQtl),
        shiftedQtyQtl: roundQtl(existing.shiftedQtyQtl + shiftedQtyQtl),
        status:
          existing.status === "OPEN" || discrepancy.status === "OPEN"
            ? "OPEN"
            : existing.status === "SHIFT_PENDING" || discrepancy.status === "SHIFT_PENDING"
              ? "SHIFT_PENDING"
              : "RESOLVED"
      });
    }

    const rows: BaseReceiptRow[] = [];

    for (const receipt of receipts) {
      const registration = registrationById.get(receipt.cropRegistrationId);
      if (!registration) {
        continue;
      }

      const seasonLabel = `${registration.season} ${registration.year}`;
      const receiptNetQty = roundQtl(
        receipt.lines.reduce(
          (sum: number, line: (typeof receipt.lines)[number]) =>
            sum + Number(line.netWeightQtl ?? line.qtyQtl ?? 0),
          0
        )
      );
      const discrepancyState = discrepancyByReceiptNo.get(receipt.receiptNo) ?? {
        pendingQtyQtl: 0,
        shiftedQtyQtl: 0,
        status: "RESOLVED"
      };
      const receiptDiscrepancyQty = roundQtl(discrepancyState.pendingQtyQtl);
      const acceptedRatio =
        receiptNetQty > 0 ? Math.max(roundQtl((receiptNetQty - receiptDiscrepancyQty) / receiptNetQty), 0) : 1;
      const discrepancyRatio =
        receiptNetQty > 0 ? Math.max(roundQtl(receiptDiscrepancyQty / receiptNetQty), 0) : 0;

      for (const line of receipt.lines) {
        const lineNet = roundQtl(Number(line.netWeightQtl ?? line.qtyQtl ?? 0));
        const lineGross = roundQtl(Number(line.grossWeightQtl ?? 0));
        const lineBags = roundInteger(Number(line.noOfBags ?? 0));
        const lotCodes = Array.from(
          new Set(
            line.allocations
              .map(
                (allocation: (typeof line.allocations)[number]) =>
                  lotCodesById.get(allocation.lotId) ?? allocation.lotCode
              )
              .filter(Boolean)
          )
        ) as string[];
        const acceptedNet = roundQtl(lineNet * acceptedRatio);
        const discrepancyNet = roundQtl(lineNet * discrepancyRatio);
        const acceptedGross = roundQtl(lineGross * acceptedRatio);
        const discrepancyGross = roundQtl(lineGross * discrepancyRatio);
        const acceptedBags = Math.max(0, Math.floor(lineBags * acceptedRatio));
        const discrepancyBags = Math.max(0, lineBags - acceptedBags);

        rows.push({
          seasonLabel,
          receiptNo: receipt.receiptNo,
          receiptDate: receipt.receiptDate,
          cropRegistrationId: registration.id,
          cropRegistrationCode: registration.cropRegistrationCode,
          farmerName: registration.farmerName,
          fatherName: registration.fatherName,
          village: registration.village,
          block: registration.block,
          district: registration.district,
          crop: registration.crop,
          variety: registration.variety,
          classStage: registration.classStage,
          certifiedAreaHa: roundQtl(Number(registration.certifiedAreaHa ?? 0)),
          expectedYieldQtl: roundQtl(Number(registration.expectedYieldQtl ?? 0)),
          allowedIntakeQtl: roundQtl(Number(registration.allowedIntakeQtl ?? 0)),
          registrationReceivedQtl: roundQtl(Number(registration.totalReceivedQtl ?? 0)),
          registrationBalanceQtl: roundQtl(Number(registration.balanceQtl ?? 0)),
          godownId: line.godownId,
          godownName: godownById.get(line.godownId) ?? "",
          stackId: String(line.stackId ?? ""),
          stackNo: String(line.stackNo ?? ""),
          vehicleNo: String(line.vehicleNo ?? ""),
          moisturePercent: roundQtl(Number(line.moisturePercent ?? 0)),
          grossWeightQtl: lineGross,
          netWeightQtl: lineNet,
          noOfBags: lineBags,
          weightPerBagKg: roundQtl(Number(line.weightPerBagKg ?? 0)),
          acceptedGrossQtl: acceptedGross,
          acceptedNetQtl: acceptedNet,
          acceptedBags,
          discrepancyGrossQtl: discrepancyGross,
          discrepancyNetQtl: discrepancyNet,
          discrepancyBags,
          discrepancyStatus: discrepancyState.status,
          discrepancyQtyQtl: receiptDiscrepancyQty,
          discrepancyShiftedQtyQtl: discrepancyState.shiftedQtyQtl,
          remarks: String(line.remarks ?? ""),
          lotCodes
        });
      }
    }

    return rows.filter((row) => {
      if (filters.season?.trim() && row.seasonLabel !== filters.season.trim()) {
        return false;
      }

      if (!dateWithinRange(row.receiptDate, filters.fromDate, filters.toDate)) {
        return false;
      }

      if (!containsText(row.crop, filters.crop)) {
        return false;
      }

      if (!containsText(row.variety, filters.variety)) {
        return false;
      }

      if (!containsText(row.classStage, filters.classStage)) {
        return false;
      }

      if (!containsText(row.district, filters.district)) {
        return false;
      }

      if (filters.godownId?.trim() && row.godownId !== filters.godownId.trim()) {
        return false;
      }

      if (filters.stackNo?.trim() && row.stackNo !== filters.stackNo.trim()) {
        return false;
      }

      if (!containsText(row.cropRegistrationCode, filters.cropRegistrationCode)) {
        return false;
      }

      if (!containsText(row.farmerName, filters.farmerName)) {
        return false;
      }

      if (filters.reportMode === "ACCEPTED_ONLY" && row.acceptedNetQtl <= 0) {
        return false;
      }

      if (filters.reportMode === "DISCREPANCY_ONLY" && row.discrepancyNetQtl <= 0) {
        return false;
      }

      return true;
    });
  }

  private buildPreviewFromRows(
    reportType: ReportType,
    rows: BaseReceiptRow[],
    filters: ReportFilterInput
  ): ReportPreview {
    const generatedAt = new Date().toISOString();
    const season = filters.season?.trim() || "ALL_SEASONS";

    if (reportType === "GODOWN_WISE_DETAIL") {
      const previewRows = rows.map((row, index) => ({
        "S.No.": index + 1,
        Crop: row.crop,
        Variety: row.variety,
        "Class/Stage": row.classStage,
        "Seed Grower": row.farmerName,
        "F/H Name": row.fatherName,
        Village: row.village,
        "Expected Yield (QTL)": row.expectedYieldQtl,
        Date: row.receiptDate,
        "Moisture (%)": row.moisturePercent,
        Bags: row.noOfBags,
        "Wt/Bag (KG)": row.weightPerBagKg,
        "Gross (QTL)": row.grossWeightQtl,
        "Net (QTL)": row.netWeightQtl,
        "Reg. Code": row.cropRegistrationCode,
        Godown: row.godownName,
        "Stack No.": row.stackNo,
        "Lot No.": row.lotCodes.join(", "),
        "Discrepancy Qty (QTL)": row.discrepancyNetQtl,
        Remarks: row.remarks || "-"
      }));

      return {
        reportType,
        title: "Godown Wise Detail",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Rows": previewRows.length,
          "Total Bags": rows.reduce((sum, row) => sum + row.noOfBags, 0),
          "Total Gross (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.grossWeightQtl, 0)),
          "Total Net (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.netWeightQtl, 0))
        },
        generatedAt,
        fileName: formatReportFileName(season, "godown-wise-detail")
      };
    }

    if (reportType === "DISTRICT_WISE_DETAIL") {
      const sortedRows = [...rows].sort((left, right) => {
        const districtCompare = String(left.district).localeCompare(String(right.district), "en", {
          sensitivity: "base"
        });
        if (districtCompare !== 0) {
          return districtCompare;
        }
        const farmerCompare = String(left.farmerName).localeCompare(String(right.farmerName), "en", {
          sensitivity: "base"
        });
        if (farmerCompare !== 0) {
          return farmerCompare;
        }
        const regCompare = String(left.cropRegistrationCode).localeCompare(
          String(right.cropRegistrationCode),
          "en",
          { sensitivity: "base", numeric: true }
        );
        if (regCompare !== 0) {
          return regCompare;
        }
        return String(left.receiptDate).localeCompare(String(right.receiptDate));
      });

      const previewRows = sortedRows.map((row, index) => ({
        "S.No.": index + 1,
        District: row.district,
        Block: row.block,
        Village: row.village,
        "Reg. Code": row.cropRegistrationCode,
        "Seed Grower": row.farmerName,
        "F/H Name": row.fatherName,
        Crop: row.crop,
        Variety: row.variety,
        "Class/Stage": row.classStage,
        "Expected Yield (QTL)": row.expectedYieldQtl,
        Date: row.receiptDate,
        "Receipt No.": row.receiptNo,
        Godown: row.godownName,
        "Stack No.": row.stackNo,
        Vehicle: row.vehicleNo,
        Bags: row.noOfBags,
        "Wt/Bag (KG)": row.weightPerBagKg,
        "Gross (QTL)": row.grossWeightQtl,
        "Net (QTL)": row.netWeightQtl,
        "Moisture (%)": row.moisturePercent,
        "Lot No.": row.lotCodes.join(", "),
        "Discrepancy Qty (QTL)": row.discrepancyNetQtl,
        Remarks: row.remarks || "-"
      }));

      return {
        reportType,
        title: "District Wise Detail",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Rows": previewRows.length,
          "Total Districts": new Set(rows.map((row) => row.district).filter(Boolean)).size,
          "Total Bags": rows.reduce((sum, row) => sum + row.noOfBags, 0),
          "Total Gross (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.grossWeightQtl, 0)),
          "Total Net (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.netWeightQtl, 0)),
          "Total Discrepancy Qty (QTL)": roundQtl(
            rows.reduce((sum, row) => sum + row.discrepancyNetQtl, 0)
          )
        },
        generatedAt,
        fileName: formatReportFileName(season, "district-wise-detail")
      };
    }

    if (reportType === "FARMER_WISE_DETAIL") {
      const previewRows = rows.map((row, index) => ({
        "S.No.": index + 1,
        Crop: row.crop,
        Variety: row.variety,
        "Class/Stage": row.classStage,
        "Grower Name": row.farmerName,
        "F/H Name": row.fatherName,
        Village: row.village,
        "Area Offered (Ha)": row.certifiedAreaHa,
        "Expected Yield (QTL)": row.expectedYieldQtl,
        Date: row.receiptDate,
        "Receipt No.": row.receiptNo,
        "Moisture (%)": row.moisturePercent,
        Bags: row.noOfBags,
        "Gross (QTL)": row.grossWeightQtl,
        "Net (QTL)": row.netWeightQtl,
        "Reg. Code": row.cropRegistrationCode,
        "Lot No.": row.lotCodes.join(", "),
        "Stack No.": row.stackNo,
        Godown: row.godownName,
        Remarks: row.remarks || "-"
      }));

      return {
        reportType,
        title: "Farmer Wise Detail",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Rows": previewRows.length,
          "Total Bags": rows.reduce((sum, row) => sum + row.noOfBags, 0),
          "Total Gross (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.grossWeightQtl, 0)),
          "Total Net (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.netWeightQtl, 0))
        },
        generatedAt,
        fileName: formatReportFileName(season, "farmer-wise-detail")
      };
    }

    if (reportType === "OVERALL_INTAKE") {
      const grouped = new Map<string, BaseReceiptRow>();
      for (const row of rows) {
        if (!grouped.has(row.cropRegistrationId)) {
          grouped.set(row.cropRegistrationId, row);
        }
      }

      const previewRows = Array.from(grouped.values())
        .sort((left, right) => {
          const districtCompare = String(left.district).localeCompare(String(right.district), "en", {
            sensitivity: "base"
          });
          if (districtCompare !== 0) {
            return districtCompare;
          }
          const farmerCompare = String(left.farmerName).localeCompare(String(right.farmerName), "en", {
            sensitivity: "base"
          });
          if (farmerCompare !== 0) {
            return farmerCompare;
          }
          return String(left.cropRegistrationCode).localeCompare(String(right.cropRegistrationCode), "en", {
            sensitivity: "base",
            numeric: true
          });
        })
        .map((row, index) => ({
          "S.No.": index + 1,
          "Farmer Name": row.farmerName,
          "Father Name": row.fatherName,
          Village: row.village,
          District: row.district,
          "Expected Yield (QTL)": row.expectedYieldQtl,
          "Net Intake Qty (QTL)": row.registrationReceivedQtl,
          "Balance Qty (QTL)": row.registrationBalanceQtl
        }));

      return {
        reportType,
        title: "Overall Intake",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Farmers": previewRows.length,
          "Total Expected Yield (QTL)": roundQtl(
            previewRows.reduce((sum, row) => sum + Number(row["Expected Yield (QTL)"] ?? 0), 0)
          ),
          "Total Net Intake Qty (QTL)": roundQtl(
            previewRows.reduce((sum, row) => sum + Number(row["Net Intake Qty (QTL)"] ?? 0), 0)
          ),
          "Total Balance Qty (QTL)": roundQtl(
            previewRows.reduce((sum, row) => sum + Number(row["Balance Qty (QTL)"] ?? 0), 0)
          )
        },
        generatedAt,
        fileName: formatReportFileName(season, "overall-intake")
      };
    }

    if (reportType === "SUMMARY") {
      const grouped = new Map<string, Record<string, string | number>>();

      for (const row of rows) {
        const key = `${row.crop}|${row.variety}|${row.classStage}`;
        const current = grouped.get(key) ?? {
          Crop: row.crop,
          Variety: row.variety,
          "Class/Stage": row.classStage,
          "Expected Yield (QTL)": 0,
          "Raw Intake Bags": 0,
          "Raw Gross (QTL)": 0,
          "Raw Net (QTL)": 0,
          "Accepted Net (QTL)": 0,
          "Discrepancy Qty (QTL)": 0,
          "Shifted Qty (QTL)": 0,
          "Lot Count": 0,
          Remarks: ""
        };

        current["Expected Yield (QTL)"] = roundQtl(
          Number(current["Expected Yield (QTL)"]) + row.expectedYieldQtl
        );
        current["Raw Intake Bags"] = Number(current["Raw Intake Bags"]) + row.noOfBags;
        current["Raw Gross (QTL)"] = roundQtl(Number(current["Raw Gross (QTL)"]) + row.grossWeightQtl);
        current["Raw Net (QTL)"] = roundQtl(Number(current["Raw Net (QTL)"]) + row.netWeightQtl);
        current["Accepted Net (QTL)"] = roundQtl(
          Number(current["Accepted Net (QTL)"]) + row.acceptedNetQtl
        );
        current["Discrepancy Qty (QTL)"] = roundQtl(
          Number(current["Discrepancy Qty (QTL)"]) + row.discrepancyNetQtl
        );
        current["Shifted Qty (QTL)"] = roundQtl(
          Number(current["Shifted Qty (QTL)"]) + row.discrepancyShiftedQtyQtl
        );
        current["Lot Count"] = Number(current["Lot Count"]) + row.lotCodes.length;
        grouped.set(key, current);
      }

      const previewRows: Record<string, string | number>[] = Array.from(grouped.values()).map((item, index) => ({
        "S.No.": index + 1,
        ...item
      }));

      return {
        reportType,
        title: "Summary",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Rows": previewRows.length,
          "Total Raw Bags": rows.reduce((sum, row) => sum + row.noOfBags, 0),
          "Total Raw Gross (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.grossWeightQtl, 0)),
          "Total Raw Net (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.netWeightQtl, 0)),
          "Total Discrepancy Qty (QTL)": roundQtl(
            rows.reduce((sum, row) => sum + row.discrepancyNetQtl, 0)
          )
        },
        generatedAt,
        fileName: formatReportFileName(season, "summary")
      };
    }

    if (reportType === "DAILY_INTAKE_REGISTER") {
      const previewRows = [...rows]
        .sort((left, right) => {
          const dateCompare = String(left.receiptDate).localeCompare(String(right.receiptDate));
          if (dateCompare !== 0) {
            return dateCompare;
          }
          const receiptCompare = String(left.receiptNo).localeCompare(String(right.receiptNo), undefined, {
            numeric: true
          });
          if (receiptCompare !== 0) {
            return receiptCompare;
          }
          return String(left.cropRegistrationCode).localeCompare(String(right.cropRegistrationCode));
        })
        .map((row, index) => ({
        "S.No.": index + 1,
        Date: row.receiptDate,
        "Receipt No.": row.receiptNo,
        "Reg. Code": row.cropRegistrationCode,
        Farmer: row.farmerName,
        Godown: row.godownName,
        Stack: row.stackNo,
        Vehicle: row.vehicleNo,
        Bags: row.noOfBags,
        "Gross (QTL)": row.grossWeightQtl,
        "Net (QTL)": row.netWeightQtl,
        "Moisture (%)": row.moisturePercent,
        "Lot No.": row.lotCodes.join(", "),
        Remarks: row.remarks || "-"
      }));

      return {
        reportType,
        title: "Daily Intake Register",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Rows": previewRows.length,
          "Total Bags": rows.reduce((sum, row) => sum + row.noOfBags, 0),
          "Total Gross (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.grossWeightQtl, 0)),
          "Total Net (QTL)": roundQtl(rows.reduce((sum, row) => sum + row.netWeightQtl, 0))
        },
        generatedAt,
        fileName: formatReportFileName(season, "daily-intake-register")
      };
    }

    if (reportType === "REGISTRATION_PENDING_RECEIVED") {
      const grouped = new Map<string, BaseReceiptRow>();
      for (const row of rows) {
        grouped.set(row.cropRegistrationId, row);
      }

      const previewRows = Array.from(grouped.values()).map((row, index) => ({
        "S.No.": index + 1,
        "Reg. Code": row.cropRegistrationCode,
        Farmer: row.farmerName,
        Village: row.village,
        Crop: row.crop,
        Variety: row.variety,
        "Class/Stage": row.classStage,
        "Expected Yield (QTL)": row.expectedYieldQtl,
        "Received (QTL)": row.registrationReceivedQtl,
        "Balance (QTL)": row.registrationBalanceQtl,
        Status: row.registrationBalanceQtl <= 0 ? "EXHAUSTED" : "ACTIVE"
      }));

      return {
        reportType,
        title: "Registration Pending vs Received",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Registrations": previewRows.length,
          "Total Expected (QTL)": roundQtl(
            previewRows.reduce((sum, row) => sum + Number(row["Expected Yield (QTL)"] ?? 0), 0)
          ),
          "Total Received (QTL)": roundQtl(
            previewRows.reduce((sum, row) => sum + Number(row["Received (QTL)"] ?? 0), 0)
          ),
          "Total Balance (QTL)": roundQtl(
            previewRows.reduce((sum, row) => sum + Number(row["Balance (QTL)"] ?? 0), 0)
          )
        },
        generatedAt,
        fileName: formatReportFileName(season, "registration-pending-vs-received")
      };
    }

    if (reportType === "LOT_WISE_STOCK_LEDGER") {
      const grouped = new Map<string, BaseReceiptRow[]>();
      for (const row of rows) {
        for (const lotCode of row.lotCodes) {
          grouped.set(lotCode, [...(grouped.get(lotCode) ?? []), row]);
        }
      }

      const previewRows = Array.from(grouped.entries()).map(([lotCode, lotRows], index) => ({
        "S.No.": index + 1,
        "Lot Code": lotCode,
        "Reg. Code": lotRows[0]?.cropRegistrationCode ?? "",
        Farmer: lotRows[0]?.farmerName ?? "",
        Godown: lotRows[0]?.godownName ?? "",
        Stack: lotRows[0]?.stackNo ?? "",
        Bags: lotRows.reduce((sum, row) => sum + row.noOfBags, 0),
        "Gross (QTL)": roundQtl(lotRows.reduce((sum, row) => sum + row.grossWeightQtl, 0)),
        "Net (QTL)": roundQtl(lotRows.reduce((sum, row) => sum + row.netWeightQtl, 0)),
        "Discrepancy Qty (QTL)": roundQtl(
          lotRows.reduce((sum, row) => sum + row.discrepancyNetQtl, 0)
        )
      }));

      return {
        reportType,
        title: "Lot-wise Stock Ledger",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Lots": previewRows.length,
          "Total Bags": previewRows.reduce((sum, row) => sum + Number(row.Bags ?? 0), 0),
          "Total Net (QTL)": roundQtl(
            previewRows.reduce((sum, row) => sum + Number(row["Net (QTL)"] ?? 0), 0)
          )
        },
        generatedAt,
        fileName: formatReportFileName(season, "lot-wise-stock-ledger")
      };
    }

    if (reportType === "STACK_WISE_STOCK_POSITION") {
      const grouped = new Map<string, Record<string, string | number>>();
      for (const row of rows) {
        const key = `${row.godownName}|${row.stackNo}`;
        const current = grouped.get(key) ?? {
          Godown: row.godownName,
          Stack: row.stackNo,
          "Reg. Count": 0,
          Bags: 0,
          "Gross (QTL)": 0,
          "Net (QTL)": 0,
          "Discrepancy Qty (QTL)": 0
        };
        current["Reg. Count"] = Number(current["Reg. Count"]) + 1;
        current.Bags = Number(current.Bags) + row.noOfBags;
        current["Gross (QTL)"] = roundQtl(Number(current["Gross (QTL)"]) + row.grossWeightQtl);
        current["Net (QTL)"] = roundQtl(Number(current["Net (QTL)"]) + row.netWeightQtl);
        current["Discrepancy Qty (QTL)"] = roundQtl(
          Number(current["Discrepancy Qty (QTL)"]) + row.discrepancyNetQtl
        );
        grouped.set(key, current);
      }

      const previewRows = Array.from(grouped.values()).map((item, index) => ({
        "S.No.": index + 1,
        ...item
      }));

      return {
        reportType,
        title: "Stack-wise Stock Position",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Stack Rows": previewRows.length,
          "Total Bags": previewRows.reduce(
            (sum, row) => sum + Number((row as Record<string, string | number>).Bags ?? 0),
            0
          ),
          "Total Net (QTL)": roundQtl(
            previewRows.reduce(
              (sum, row) =>
                sum + Number((row as Record<string, string | number>)["Net (QTL)"] ?? 0),
              0
            )
          )
        },
        generatedAt,
        fileName: formatReportFileName(season, "stack-wise-stock-position")
      };
    }

    if (reportType === "STACK_CARD_REGISTER") {
      const grouped = new Map<
        string,
        {
          godownName: string;
          stackNo: string;
          farmerName: string;
          regCode: string;
          bags: number;
          netQtl: number;
        }
      >();

      for (const row of rows) {
        const key = `${row.godownName}|${row.stackNo}|${row.cropRegistrationCode}`;
        const current = grouped.get(key) ?? {
          godownName: row.godownName,
          stackNo: row.stackNo,
          farmerName: row.farmerName,
          regCode: row.cropRegistrationCode,
          bags: 0,
          netQtl: 0
        };
        current.bags += row.noOfBags;
        current.netQtl = roundQtl(current.netQtl + row.netWeightQtl);
        grouped.set(key, current);
      }

      const previewRows = Array.from(grouped.values())
        .sort((left, right) => {
          const godownCompare = left.godownName.localeCompare(right.godownName, "en", {
            sensitivity: "base"
          });
          if (godownCompare !== 0) {
            return godownCompare;
          }
          const stackCompare = left.stackNo.localeCompare(right.stackNo, "en", {
            sensitivity: "base",
            numeric: true
          });
          if (stackCompare !== 0) {
            return stackCompare;
          }
          return left.regCode.localeCompare(right.regCode, "en", {
            sensitivity: "base",
            numeric: true
          });
        })
        .map((item, index) => ({
          "S.No.": index + 1,
          Godown: item.godownName,
          Stack: item.stackNo,
          "FARMER NAME": item.farmerName,
          "FARMER REG CODE": item.regCode,
          "NUMBER OF BAGS IN STACK": item.bags,
          "TOTAL NET WEIGHT IN STACK (QTL)": item.netQtl
        }));

      return {
        reportType,
        title: "Stack Card Register",
        columns: Object.keys(previewRows[0] ?? { "S.No.": 1 }),
        rows: previewRows,
        totals: {
          "Total Rows": previewRows.length,
          "Total Stacks": new Set(
            previewRows.map((row) => `${String(row.Godown)}|${String(row.Stack)}`)
          ).size,
          "Total Bags": previewRows.reduce(
            (sum, row) => sum + Number(row["NUMBER OF BAGS IN STACK"] ?? 0),
            0
          ),
          "Total Net (QTL)": roundQtl(
            previewRows.reduce(
              (sum, row) => sum + Number(row["TOTAL NET WEIGHT IN STACK (QTL)"] ?? 0),
              0
            )
          )
        },
        generatedAt,
        fileName: formatReportFileName(season, "stack-card-register")
      };
    }

    const discrepancyRows = rows
      .filter((row) => row.discrepancyNetQtl > 0 || row.discrepancyStatus !== "RESOLVED")
      .map((row, index) => ({
        "S.No.": index + 1,
        Date: row.receiptDate,
        "Receipt No.": row.receiptNo,
        "Reg. Code": row.cropRegistrationCode,
        Farmer: row.farmerName,
        Godown: row.godownName,
        Stack: row.stackNo,
        "Expected Yield (QTL)": row.expectedYieldQtl,
        "Receipt Net (QTL)": row.netWeightQtl,
        "Pending Excess (QTL)": row.discrepancyNetQtl,
        "Shifted Qty (QTL)": row.discrepancyShiftedQtyQtl,
        Status: row.discrepancyStatus,
        Remarks: row.remarks || "-"
      }));

    return {
      reportType,
      title: "Discrepancy Register",
      columns: Object.keys(discrepancyRows[0] ?? { "S.No.": 1 }),
      rows: discrepancyRows,
      totals: {
        "Total Rows": discrepancyRows.length,
        "Pending Excess (QTL)": roundQtl(
          discrepancyRows.reduce((sum, row) => sum + Number(row["Pending Excess (QTL)"] ?? 0), 0)
        ),
        "Shifted Qty (QTL)": roundQtl(
          discrepancyRows.reduce((sum, row) => sum + Number(row["Shifted Qty (QTL)"] ?? 0), 0)
        )
      },
      generatedAt,
      fileName: formatReportFileName(season, "discrepancy-register")
    };
  }

  private buildWorkbook(previews: ReportPreview[]) {
    const workbook = XLSX.utils.book_new();
    const usedSheetNames = new Set<string>();

    for (const preview of previews) {
      const rows: (string | number)[][] = [];
      rows.push([preview.title]);
      rows.push([`Generated At: ${preview.generatedAt}`]);
      rows.push([]);
      rows.push(preview.columns);
      for (const row of preview.rows) {
        rows.push(preview.columns.map((column) => row[column] ?? ""));
      }
      rows.push([]);
      rows.push(["Totals"]);
      for (const [key, value] of Object.entries(preview.totals)) {
        rows.push([key, value]);
      }

      const worksheet = XLSX.utils.aoa_to_sheet(rows);
      const baseSheetName = preview.title.slice(0, 31) || "Report";
      let sheetName = baseSheetName;
      let duplicateCounter = 2;
      while (usedSheetNames.has(sheetName)) {
        const suffix = `-${duplicateCounter}`;
        sheetName = `${baseSheetName.slice(0, Math.max(31 - suffix.length, 1))}${suffix}`;
        duplicateCounter += 1;
      }
      usedSheetNames.add(sheetName);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    return workbook;
  }

  async previewReport(filters: ReportFilterInput) {
    const rows = await this.buildBaseReceiptRows(filters);
    return this.buildPreviewFromRows(filters.reportType as ReportType, rows, filters);
  }

  async exportReports(filters: ReportFilterInput) {
    const rows = await this.buildBaseReceiptRows(filters);
    let previews: ReportPreview[];

    if ((filters.reportType as ReportType) === "DISTRICT_WISE_DETAIL") {
      const districtGroups = new Map<string, BaseReceiptRow[]>();

      for (const row of rows) {
        const districtKey = String(row.district || "UNSPECIFIED").trim() || "UNSPECIFIED";
        districtGroups.set(districtKey, [...(districtGroups.get(districtKey) ?? []), row]);
      }

      previews = Array.from(districtGroups.entries())
        .sort(([left], [right]) => left.localeCompare(right, "en", { sensitivity: "base" }))
        .map(([district, districtRows]) => {
          const preview = this.buildPreviewFromRows("DISTRICT_WISE_DETAIL", districtRows, {
            ...filters,
            district
          });
          return {
            ...preview,
            title: district || "UNSPECIFIED"
          };
        });

      if (previews.length === 0) {
        previews = [this.buildPreviewFromRows("DISTRICT_WISE_DETAIL", rows, filters)];
      }
    } else if ((filters.reportType as ReportType) === "STACK_CARD_REGISTER") {
      const stackGroups = new Map<string, BaseReceiptRow[]>();

      for (const row of rows) {
        const key = `${String(row.godownName || "UNSPECIFIED").trim() || "UNSPECIFIED"}|${String(
          row.stackNo || "UNSPECIFIED"
        ).trim() || "UNSPECIFIED"}`;
        stackGroups.set(key, [...(stackGroups.get(key) ?? []), row]);
      }

      previews = Array.from(stackGroups.entries())
        .sort(([left], [right]) => left.localeCompare(right, "en", { sensitivity: "base", numeric: true }))
        .map(([stackKey, stackRows]) => {
          const [godownName, stackNo] = stackKey.split("|");
          const preview = this.buildPreviewFromRows("STACK_CARD_REGISTER", stackRows, filters);
          return {
            ...preview,
            title: `${godownName} - Stack ${stackNo}`
          };
        });

      if (previews.length === 0) {
        previews = [this.buildPreviewFromRows("STACK_CARD_REGISTER", rows, filters)];
      }
    } else {
      previews = [this.buildPreviewFromRows(filters.reportType as ReportType, rows, filters)];
    }

    const workbook = this.buildWorkbook(previews);
    const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });

    return {
      fileName: previews[0]?.fileName ?? formatReportFileName(filters.season?.trim() || "ALL_SEASONS", "report"),
      content: buffer
    };
  }

  async validateAllDiscrepancies() {
    const registrationsWithReceipts = await ReceiptModel.distinct("cropRegistrationId");

    let reconciledRegistrations = 0;
    for (const registrationId of registrationsWithReceipts) {
      await this.reconcileDiscrepanciesForRegistration(String(registrationId));
      reconciledRegistrations += 1;
    }

    const discrepancies = await DiscrepancyModel.find().lean();
    const summary = {
      reconciledRegistrations,
      openDiscrepancies: discrepancies.filter((item) => item.status !== "RESOLVED").length,
      resolvedDiscrepancies: discrepancies.filter((item) => item.status === "RESOLVED").length,
      totalPendingExcessQtyQtl: roundQtl(
        discrepancies.reduce((sum, item) => sum + Number(item.excessQtyQtl ?? 0), 0)
      )
    };

    await AuditLogModel.create({
      entityType: "VALIDATION",
      entityId: "DISCREPANCY",
      action: "AUTO_RECONCILED_ALL",
      payload: summary
    });

    return {
      ...(await this.bootstrap()),
      validationSummary: {
        type: "DISCREPANCY_AUTO_VALIDATION",
        ...summary
      }
    };
  }

  async validateLots() {
    const receipts = await ReceiptModel.find().lean();
    const lots = await LotModel.find().lean();
    const allocationQtyByLotId = new Map<string, number>();

    for (const receipt of receipts) {
      for (const line of receipt.lines) {
        for (const allocation of line.allocations) {
          allocationQtyByLotId.set(
            allocation.lotId,
            roundQtl(
              (allocationQtyByLotId.get(allocation.lotId) ?? 0) + Number(allocation.qtyQtl ?? 0)
            )
          );
        }
      }
    }

    let updatedLots = 0;
    let overCapLots = 0;
    let orphanLotsVoided = 0;

    for (const lot of lots) {
      const allocatedQtyQtl = roundQtl(allocationQtyByLotId.get(lot.id) ?? 0);

      if (allocatedQtyQtl <= 0) {
        await LotModel.updateOne(
          { id: lot.id },
          {
            $set: {
              currentQtyQtl: 0,
              status: "VOID",
              voidReason: "Auto-validation: no active allocations found."
            }
          }
        );
        orphanLotsVoided += 1;
        continue;
      }

      const nextStatus = allocatedQtyQtl >= lot.maxAllowedQtyQtl ? "FULL" : "OPEN";
      if (
        roundQtl(Number(lot.currentQtyQtl ?? 0)) !== allocatedQtyQtl ||
        String(lot.status) !== nextStatus
      ) {
        await LotModel.updateOne(
          { id: lot.id },
          {
            $set: {
              currentQtyQtl: allocatedQtyQtl,
              status: nextStatus,
              voidReason: ""
            }
          }
        );
        updatedLots += 1;
      }

      if (allocatedQtyQtl - Number(lot.maxAllowedQtyQtl ?? 200) > 0.0001) {
        overCapLots += 1;
      }
    }

    const summary = {
      updatedLots,
      orphanLotsVoided,
      overCapLots,
      totalLotsChecked: lots.length
    };

    await AuditLogModel.create({
      entityType: "VALIDATION",
      entityId: "LOT",
      action: "AUTO_VALIDATED_ALL",
      payload: summary
    });

    return {
      ...(await this.bootstrap()),
      validationSummary: {
        type: "LOT_AUTO_VALIDATION",
        ...summary
      }
    };
  }

  async reindexLots() {
    const [lots, registrations] = await Promise.all([
      LotModel.find().sort({ cropRegistrationId: 1, lotNo: 1, createdAt: 1, lotCode: 1 }).lean(),
      RegistrationModel.find({}, { id: 1, year: 1, cropRegistrationCode: 1, _id: 0 }).lean()
    ]);

    const registrationById = new Map(
      registrations.map((registration) => [registration.id, registration])
    );
    const lotsByRegistrationId = new Map<string, typeof lots>();

    for (const lot of lots) {
      lotsByRegistrationId.set(lot.cropRegistrationId, [
        ...(lotsByRegistrationId.get(lot.cropRegistrationId) ?? []),
        lot
      ]);
    }

    let reindexedLots = 0;
    let updatedReceiptAllocations = 0;
    let affectedRegistrations = 0;

    for (const [registrationId, registrationLots] of lotsByRegistrationId.entries()) {
      const sortedLots = [...registrationLots].sort((left, right) => {
        if (Number(left.lotNo ?? 0) !== Number(right.lotNo ?? 0)) {
          return Number(left.lotNo ?? 0) - Number(right.lotNo ?? 0);
        }

        return String(left.createdAt ?? "").localeCompare(String(right.createdAt ?? ""));
      });

      const registration = registrationById.get(registrationId);
      const changedLotCodes = new Map<string, string>();
      let registrationChanged = false;

      for (let index = 0; index < sortedLots.length; index += 1) {
        const lot = sortedLots[index];
        const nextLotNo = index + 1;
        const lotPrefix = registration
          ? `${registration.year}/${registration.cropRegistrationCode}`
          : String(lot.lotCode).replace(/\/L\d+$/i, "");
        const nextLotCode = `${lotPrefix}/L${nextLotNo}`;

        if (Number(lot.lotNo ?? 0) === nextLotNo && String(lot.lotCode) === nextLotCode) {
          continue;
        }

        await LotModel.updateOne(
          { id: lot.id },
          {
            $set: {
              lotNo: nextLotNo,
              lotCode: nextLotCode
            }
          }
        );

        changedLotCodes.set(lot.id, nextLotCode);
        reindexedLots += 1;
        registrationChanged = true;
      }

      if (!registrationChanged) {
        continue;
      }

      affectedRegistrations += 1;

      const receipts = await ReceiptModel.find({
        cropRegistrationId: registrationId
      }).lean();

      for (const receipt of receipts) {
        let receiptChanged = false;
        const nextLines = receipt.lines.map((line: (typeof receipt.lines)[number]) => {
          let lineChanged = false;
          const nextAllocations = line.allocations.map(
            (allocation: (typeof line.allocations)[number]) => {
            const nextLotCode = changedLotCodes.get(allocation.lotId);
            if (!nextLotCode || allocation.lotCode === nextLotCode) {
              return allocation;
            }

            lineChanged = true;
            updatedReceiptAllocations += 1;
            return {
              ...allocation,
              lotCode: nextLotCode
            };
            }
          );

          if (!lineChanged) {
            return line;
          }

          receiptChanged = true;
          return {
            ...line,
            allocations: nextAllocations
          };
        });

        if (receiptChanged) {
          await ReceiptModel.updateOne(
            { id: receipt.id },
            {
              $set: {
                lines: nextLines
              }
            }
          );
        }
      }
    }

    const summary = {
      reindexedLots,
      updatedReceiptAllocations,
      affectedRegistrations,
      totalLotsChecked: lots.length
    };

    await AuditLogModel.create({
      entityType: "VALIDATION",
      entityId: "LOT_REINDEX",
      action: "AUTO_REINDEXED_ALL",
      payload: summary
    });

    return {
      ...(await this.bootstrap()),
      validationSummary: {
        type: "LOT_REINDEX",
        ...summary
      }
    };
  }
}
