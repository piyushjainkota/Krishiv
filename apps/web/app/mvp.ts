export type RegistrationStatus = "ACTIVE" | "BLOCKED" | "CLOSED" | "EXHAUSTED";
export type QualityStatus = "ACCEPTED" | "HOLD";
export type DiscrepancyStatus = "OPEN" | "SHIFT_PENDING" | "RESOLVED";
export type AppRole = "ADMIN" | "MANAGER" | "USER";

export interface AppUser {
  name: string;
  email: string;
  role: AppRole;
}

export interface RolePermissions {
  canEntry: boolean;
  canEdit: boolean;
  canDelete: boolean;
  canImport: boolean;
  canValidate: boolean;
  canVoucher: boolean;
  canShift: boolean;
  canMaintenance: boolean;
}

export interface RegistrationRecord {
  id: string;
  season: string;
  year: string;
  ro: string;
  cropRegistrationCode: string;
  farmerName: string;
  fatherName: string;
  village: string;
  block: string;
  district: string;
  crop: string;
  variety: string;
  classStage: string;
  registeredAreaHa: number;
  inspectedAreaHa: number;
  rejectedAreaHa: number;
  certifiedAreaHa: number;
  expectedYieldQtl: number;
  allowedIntakeQtl: number;
  totalReceivedQtl: number;
  balanceQtl: number;
  organizerId?: string;
  organizerName?: string;
  organizerCommissionRatePerQtl?: number;
  status: RegistrationStatus;
  sourceRowNumber: number;
}

export interface Organizer {
  id: string;
  name: string;
  mobile: string;
  village: string;
  district: string;
  commissionRatePerQtl: number;
  deductionAmount: number;
  isActive: boolean;
}

export interface Godown {
  id: string;
  name: string;
}

export interface Stack {
  id: string;
  godownId: string;
  stackNo: string;
}

export interface CertificationLot {
  id: string;
  cropRegistrationId: string;
  cropRegistrationCode: string;
  lotNo: number;
  lotCode: string;
  godownId: string;
  godownName: string;
  stackId: string;
  stackNo: string;
  qualityStatus: QualityStatus;
  currentQtyQtl: number;
  maxAllowedQtyQtl: number;
  status: "OPEN" | "FULL" | "VOID" | "CANCELLED";
  voidReason?: string;
  createdAt: string;
}

export interface IntakeReceiptLine {
  id: string;
  godownId: string;
  stackId?: string;
  stackNo: string;
  grossWeightQtl: number;
  qtyQtl: number;
  noOfBags: number;
  weightPerBagKg: number;
  netWeightQtl: number;
  moisturePercent: number;
  vehicleNo: string;
  remarks: string;
  allocations: {
    lotId: string;
    lotCode: string;
    qtyQtl: number;
  }[];
}

export interface IntakeReceipt {
  id: string;
  receiptNo: string;
  receiptDate: string;
  cropRegistrationId: string;
  cropRegistrationCode: string;
  farmerName: string;
  lines: IntakeReceiptLine[];
}

export interface IntakeDiscrepancy {
  id: string;
  discrepancyNo: string;
  season: string;
  cropRegistrationId: string;
  cropRegistrationCode: string;
  farmerName: string;
  receiptNo: string;
  receiptDate: string;
  godownId: string;
  godownName: string;
  stackId: string;
  stackNo: string;
  expectedQtyQtl: number;
  receiptNetQtyQtl: number;
  totalReceivedAfterReceiptQtl: number;
  excessQtyQtl: number;
  estimatedExcessBags: number;
  handlingMode: "MARK_STACK_FOR_SHIFT";
  status: DiscrepancyStatus;
  remarks: string;
  createdAt: string;
}

export interface DiscrepancyShift {
  id: string;
  discrepancyId: string;
  discrepancyNo: string;
  cropRegistrationCode: string;
  farmerName: string;
  fromGodownName: string;
  fromStackNo: string;
  toGodownId: string;
  toGodownName: string;
  toStackId: string;
  toStackNo: string;
  shiftedQtyQtl: number;
  shiftedBags: number;
  shiftDate: string;
  approvedBy: string;
  remarks: string;
  createdAt: string;
}

export interface FinancialVoucherLine {
  receiptId: string;
  receiptNo: string;
  receiptDate: string;
  vehicleNo: string;
  stackNo: string;
  bags: number;
  grossQtyQtl: number;
  netQtyQtl: number;
}

export interface FinancialVoucherPayment {
  id: string;
  paymentDate: string;
  amount: number;
  transactionNo: string;
  mode: string;
  remarks: string;
  createdBy?: string;
  updatedBy?: string;
}

export interface OrganizerPayment {
  id: string;
  organizerId: string;
  organizerName: string;
  paymentDate: string;
  amount: number;
  transactionNo: string;
  remarks: string;
  createdBy?: string;
  updatedBy?: string;
}

export interface FinancialVoucher {
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
  lastPaymentDate?: string;
  status: string;
  remarks: string;
  lines: FinancialVoucherLine[];
  payments: FinancialVoucherPayment[];
  createdAt?: string;
}

export function calculateNetWeightQtl(grossWeightQtl: number, noOfBags: number): number {
  return roundQtl(grossWeightQtl - grossWeightQtl * 0.003 - noOfBags * 0.007);
}

export function calculateWeightPerBagKg(netWeightQtl: number, noOfBags: number): number {
  if (noOfBags <= 0) {
    return 0;
  }

  return roundQtl((netWeightQtl * 100) / noOfBags);
}

function textBetween(source: string, prefix: string): string {
  const line = source
    .split("\n")
    .map((item) => item.trim())
    .find((item) => item.startsWith(prefix));

  return line ? line.slice(prefix.length).trim().replace(/,$/, "") : "";
}

function extractSeason(value: string): { season: string; year: string } {
  const match = value.match(/^([A-Z]+)[\s-]*(\d{4}-\d{2})$/i);
  if (!match) {
    return { season: "RABI", year: "2025-26" };
  }

  return {
    season: match[1].toUpperCase(),
    year: match[2]
  };
}

function roundQtl(value: number): number {
  return Number(value.toFixed(2));
}

export function parseRegistrationWorkbook(rows: unknown[][]): RegistrationRecord[] {
  const result: RegistrationRecord[] = [];

  for (let index = 1; index < rows.length; index += 1) {
    const row = rows[index];
    if (!row || row.length === 0 || !row[1] || !row[2]) {
      continue;
    }

    const seedGrowerDetails = String(row[1] ?? "");
    const cropDetails = String(row[2] ?? "");
    const landDetails = String(row[3] ?? "");

    const cropRegistrationCode = textBetween(cropDetails, "Crop Registration Code:")
      .replace(/\(MrNo:.*?\)/g, "")
      .trim();
    const crop = textBetween(cropDetails, "Crop:");
    const variety = textBetween(cropDetails, "Variety:");
    const classStage = textBetween(cropDetails, "Class:");
    const ro = textBetween(landDetails, "RO:");
    const farmerName = textBetween(seedGrowerDetails, "Name:");
    const fatherName = textBetween(seedGrowerDetails, "F/H Name:");
    const village = textBetween(seedGrowerDetails, "Village:");
    const block = textBetween(seedGrowerDetails, "Block:");
    const district = textBetween(seedGrowerDetails, "District:");
    const certifiedAreaHa = Number(row[8] ?? 0);
    const expectedYieldQtl = Number(row[9] ?? 0);
    const initialReceivedQtl = Number(row[10] ?? 0);
    const allowedIntakeQtl = roundQtl(Math.min(expectedYieldQtl, expectedYieldQtl));
    const totalReceivedQtl = roundQtl(initialReceivedQtl);
    const balanceQtl = roundQtl(Math.max(allowedIntakeQtl - totalReceivedQtl, 0));

    let status: RegistrationStatus = "ACTIVE";
    if (certifiedAreaHa <= 0 || expectedYieldQtl <= 0) {
      status = "BLOCKED";
    } else if (balanceQtl <= 0) {
      status = "EXHAUSTED";
    }

    if (!cropRegistrationCode) {
      continue;
    }

    result.push({
      id: `reg-${cropRegistrationCode}`,
      season: "RABI",
      year: "2025-26",
      ro,
      cropRegistrationCode,
      farmerName,
      fatherName,
      village,
      block,
      district,
      crop,
      variety,
      classStage,
      registeredAreaHa: Number(row[4] ?? 0),
      inspectedAreaHa: Number(row[5] ?? 0),
      rejectedAreaHa: Number(row[7] ?? 0) || 0,
      certifiedAreaHa,
      expectedYieldQtl,
      allowedIntakeQtl,
      totalReceivedQtl,
      balanceQtl,
      status,
      sourceRowNumber: index + 1
    });
  }

  return result;
}

export function parseReportWorkbookMeta(rows: unknown[][]): { season: string; year: string } {
  const seasonRow = rows
    .flat()
    .find((cell) => String(cell ?? "").toUpperCase().includes("SEASON"));

  return extractSeason(String(seasonRow ?? "RABI-2025-26").replace("Season:", "").trim());
}

export function planReceiptAllocation(params: {
  registration: RegistrationRecord;
  openLots: CertificationLot[];
  godown: Godown;
  stack: Stack;
  qtyQtl: number;
  qualityStatus: QualityStatus;
}): {
  createLots: Omit<CertificationLot, "id">[];
  updateLots: { lotId: string; nextQtyQtl: number; nextStatus: "OPEN" | "FULL" }[];
  allocations: { tempLotKey: string; lotCode: string; qtyQtl: number }[];
} {
  const { registration, openLots, godown, stack, qtyQtl, qualityStatus } = params;

  if (registration.status !== "ACTIVE") {
    throw new Error("Selected registration is not active for intake.");
  }

  if (qtyQtl <= 0) {
    throw new Error("Receipt quantity must be greater than zero.");
  }

  if (roundQtl(registration.totalReceivedQtl + qtyQtl) > registration.allowedIntakeQtl) {
    throw new Error("Receipt quantity exceeds allowed intake for this registration.");
  }

  const matchingLots = openLots
    .filter(
      (lot) =>
        lot.cropRegistrationId === registration.id &&
        lot.godownId === godown.id &&
        lot.stackId === stack.id &&
        lot.qualityStatus === qualityStatus &&
        lot.status === "OPEN"
    )
    .sort((a, b) => a.lotNo - b.lotNo);

  let remainingQty = roundQtl(qtyQtl);
  const updateLots: { lotId: string; nextQtyQtl: number; nextStatus: "OPEN" | "FULL" }[] = [];
  const createLots: Omit<CertificationLot, "id">[] = [];
  const allocations: { tempLotKey: string; lotCode: string; qtyQtl: number }[] = [];
  let nextLotNo =
    openLots.length > 0
      ? Math.max(
          ...openLots
            .filter((lot) => lot.cropRegistrationId === registration.id)
            .map((lot) => lot.lotNo)
        ) + 1
      : 1;

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

    updateLots.push({
      lotId: lot.id,
      nextQtyQtl: nextQty,
      nextStatus: nextQty >= lot.maxAllowedQtyQtl ? "FULL" : "OPEN"
    });

    allocations.push({
      tempLotKey: lot.id,
      lotCode: lot.lotCode,
      qtyQtl: allocatedQty
    });
  }

  while (remainingQty > 0) {
    const allocatedQty = roundQtl(Math.min(remainingQty, 200));
    const lotCode = `${registration.year}/${registration.cropRegistrationCode}/L${nextLotNo}`;
    const tempKey = `new-${registration.id}-${nextLotNo}`;

    createLots.push({
      cropRegistrationId: registration.id,
      cropRegistrationCode: registration.cropRegistrationCode,
      lotNo: nextLotNo,
      lotCode,
      godownId: godown.id,
      godownName: godown.name,
      stackId: stack.id,
      stackNo: stack.stackNo,
      qualityStatus,
      currentQtyQtl: allocatedQty,
      maxAllowedQtyQtl: 200,
      status: allocatedQty >= 200 ? "FULL" : "OPEN",
      createdAt: new Date().toISOString()
    });

    allocations.push({
      tempLotKey: tempKey,
      lotCode,
      qtyQtl: allocatedQty
    });

    nextLotNo += 1;
    remainingQty = roundQtl(remainingQty - allocatedQty);
  }

  return {
    createLots,
    updateLots,
    allocations
  };
}

export function nextRegistrationState(
  registration: RegistrationRecord,
  addedQtyQtl: number
): RegistrationRecord {
  const totalReceivedQtl = roundQtl(registration.totalReceivedQtl + addedQtyQtl);
  const balanceQtl = roundQtl(Math.max(registration.allowedIntakeQtl - totalReceivedQtl, 0));

  return {
    ...registration,
    totalReceivedQtl,
    balanceQtl,
    status: balanceQtl <= 0 ? "EXHAUSTED" : registration.certifiedAreaHa <= 0 || registration.expectedYieldQtl <= 0 ? "BLOCKED" : "ACTIVE"
  };
}
