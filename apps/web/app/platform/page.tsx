"use client";

import { useEffect, useRef, useState } from "react";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import type {
  AppUser,
  AppRole,
  CertificationLot,
  DiscrepancyShift,
  FinancialVoucher,
  FinancialVoucherPayment,
  Godown,
  IntakeDiscrepancy,
  IntakeReceipt,
  IntakeReceiptLine,
  Organizer,
  OrganizerPayment,
  RegistrationRecord,
  RolePermissions,
  StackAccommodation,
  Stack
} from "../mvp";
import {
  calculateNetWeightQtl,
  calculateWeightPerBagKg,
  parseRegistrationWorkbook,
  parseReportWorkbookMeta
} from "../mvp";

type ViewKey =
  | "dashboard"
  | "masters"
  | "import"
  | "registrations"
  | "intake"
  | "intakeEdit"
  | "reports"
  | "finance"
  | "commission"
  | "slips"
  | "lots"
  | "discrepancies"
  | "validations"
  | "backup"
  | "restore";

type SidebarSectionKey =
  | "overview"
  | "masterData"
  | "intakeOps"
  | "stockLots"
  | "financeOps"
  | "reporting"
  | "administration";
const API_BASE =
  process.env.NEXT_PUBLIC_API_BASE?.trim() ||
  (typeof window !== "undefined"
    ? `${window.location.protocol}//${window.location.hostname}:4000`
    : "http://127.0.0.1:4000");
const BRAND_LOGO_SRC = "/krishiv-logo.svg";
const COMPANY_NAME = "KRISHIV AGRI GENETICS LLP";
const IMPORT_PAGE_SIZE = 20;
const LOT_CAPACITY_QTL = 200;
type ImportSortKey = "farmerName" | "village" | "classStage";
type RegistrationSortKey =
  | "cropRegistrationCode"
  | "farmerName"
  | "village"
  | "expectedYieldQtl"
  | "totalReceivedQtl"
  | "balanceQtl"
  | "status";

type OrganizerCommissionRow = {
  organizer: Organizer;
  linkedRegistrations: RegistrationRecord[];
  farmerCount: number;
  totalIntakeQtl: number;
  ratePerQtl: number;
  grossCommissionAmount: number;
  deductionAmount: number;
  netPayableAmount: number;
  paidAmount: number;
  balanceAmount: number;
  paymentCount: number;
};

type OrganizerDashboardRow = OrganizerCommissionRow & {
  coveragePct: number;
};

type OrganizerPerformanceRow = {
  organizerId: string;
  organizerName: string;
  district: string;
  farmerCount: number;
  expectedYieldQtl: number;
  depositedQtl: number;
  pendingQtl: number;
  coveragePct: number;
};

type OrganizerFarmerPaymentPendingRow = {
  organizerId: string;
  organizerName: string;
  district: string;
  farmerCount: number;
  voucherCount: number;
  netPayableAmount: number;
  paidAmount: number;
  pendingAmount: number;
  paymentCompletionPct: number;
  overpaidCount: number;
};

  const navSections: {
    key: SidebarSectionKey;
    label: string;
    items: { key: ViewKey; label: string }[];
  }[] = [
    {
      key: "overview",
      label: "Overview",
      items: [{ key: "dashboard", label: "Dashboard" }]
    },
    {
      key: "masterData",
      label: "Masters",
      items: [
        { key: "masters", label: "Godown & Stack Masters" },
        { key: "import", label: "Farmer Master Import" },
        { key: "registrations", label: "Registration Master" }
      ]
    },
    {
      key: "intakeOps",
      label: "Intake",
      items: [
        { key: "intake", label: "Intake Entry" },
        { key: "intakeEdit", label: "Intake Entry Edit" },
        { key: "slips", label: "Slip Print Center" }
      ]
    },
    {
      key: "stockLots",
      label: "Stock & Lots",
      items: [
        { key: "lots", label: "Lot Ledger" },
        { key: "discrepancies", label: "Discrepancy Register" }
      ]
    },
    {
      key: "financeOps",
      label: "Finance",
      items: [
        { key: "finance", label: "Financial Voucher" },
        { key: "commission", label: "Organizer Commission" }
      ]
    },
    {
      key: "reporting",
      label: "Reports",
      items: [{ key: "reports", label: "Reports" }]
    },
    {
      key: "administration",
      label: "Administration",
      items: [
        { key: "validations", label: "Validation Center" },
        { key: "backup", label: "Database Backup" },
        { key: "restore", label: "Database Restore" }
      ]
    }
  ];

const rolePermissions: Record<AppRole, RolePermissions> = {
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

type ReportType =
  | "GODOWN_WISE_DETAIL"
  | "DISTRICT_WISE_DETAIL"
  | "FARMER_WISE_DETAIL"
  | "OVERALL_INTAKE"
  | "SUMMARY"
  | "DAILY_INTAKE_REGISTER"
  | "CUSTOM_DATE_PAYMENT_REGISTER"
  | "ORGANIZER_FARMER_PAYMENT_REGISTER"
  | "ORGANIZER_PAYMENT_TRANSACTION_REPORT"
  | "OVERPAID_FARMER_REPORT"
  | "RECEIPT_VOUCHER_TRACEABILITY_REPORT"
  | "ORGANIZER_INTAKE_PAYMENT_COMMISSION_REPORT"
  | "REGISTRATION_PENDING_RECEIVED"
  | "LOT_WISE_STOCK_LEDGER"
  | "ADJUSTED_LOT_FORMATION_REGISTER"
  | "ADJUSTED_LOT_LEDGER_FARMER_WISE"
  | "STACK_WISE_STOCK_POSITION"
  | "STACK_CARD_REGISTER"
  | "DISCREPANCY_REGISTER";

type ReportMode = "ALL" | "ACCEPTED_ONLY" | "DISCREPANCY_ONLY";
type SlipType = "FARMER_SINGLE_RECEIPT" | "FARMER_OVERALL" | "DAILY_CONSOLIDATED";

type ReportPreview = {
  reportType: ReportType;
  title: string;
  columns: string[];
  rows: Record<string, string | number>[];
  totals: Record<string, string | number>;
  generatedAt: string;
  fileName: string;
};

type OrganizerPaymentTransactionBlock = {
  organizerName: string;
  regCode: string;
  name: string;
  village: string;
  district: string;
  seasonLabel: string;
  totalBags: number;
  totalNetQtyQtl: number;
  ratePerQtl: number;
  grossAmount: number;
  deduction: number;
  finalAmount: number;
  payments: {
    transactionDate: string;
    payment: number;
    transactionNumber: string;
    transactionRemark: string;
  }[];
};

type StackCardRegisterSection = {
  key: string;
  godownId: string;
  godownName: string;
  stackNo: string;
  displayMode: "ORIGINAL" | "FINAL_ADJUSTED";
  rows: {
    regCode: string;
    farmerName: string;
    village: string;
    district: string;
    qtyQtl: number;
    bags: number;
    mark: string;
  }[];
  totalQtyQtl: number;
  totalBags: number;
  changedFarmerCount: number;
};

type AdjustedLotFormationRow = {
  srNo: number;
  regNo: string;
  farmerName: string;
  paymentVoucherNo: string;
  expectedYieldQtl: number;
  totalNetIntakeQtl: number;
  warehouseName: string;
  stackNumber: string;
  bags: number;
  netWeightQtl: number;
  lotId: string;
  moisturePercent: number;
  mark: string;
};

type DashboardAssistantResult = {
  title: string;
  summary: string;
  columns: string[];
  rows: Record<string, string | number>[];
};

type AdjustedStackCardPreview = {
  title: string;
  godownName: string;
  stackNo: string;
  generatedAt: string;
  fileName: string;
  totalAccommodatedQtyQtl: number;
  totalAccommodatedBags: number;
  changedFarmerCount: number;
  originalRows: {
    regCode: string;
    farmerName: string;
    village: string;
    district: string;
    qtyQtl: number;
    bags: number;
  }[];
    adjustedRows: {
      regCode: string;
      farmerName: string;
      village: string;
      district: string;
    originalQtyQtl: number;
    adjustedInQtyQtl: number;
      adjustedOutQtyQtl: number;
      finalQtyQtl: number;
      originalBags: number;
      adjustedInBags: number;
      adjustedOutBags: number;
      finalBags: number;
      changed: boolean;
    }[];
};

type PaymentRegisterRow = {
  date: string;
  voucherNo: string;
  farmerName: string;
  regCode: string;
  village: string;
  district: string;
  status: string;
  transactionNo: string;
  mode: string;
  amount: number;
};

type SlipPreview = {
  slipType: SlipType;
  title: string;
  pageSize: "A5" | "A4";
  orientation: "portrait";
  template?: "FARMER_SINGLE_CLASSIC" | "FARMER_OVERALL_CLASSIC" | "DAILY_CONSOLIDATED_CLASSIC";
  slipNo?: string;
  summary: { label: string; value: string }[];
  tableColumns: string[];
  tableRows: Record<string, string | number>[];
  totals: { label: string; value: string }[];
  extraSections?: { title: string; lines: { label: string; value: string }[] }[];
  footerNote: string;
};

type VoucherPreview = {
  voucher: FinancialVoucher;
  hasDiscrepancy: boolean;
};

type CoreBootstrapPayload = {
  registrations: RegistrationRecord[];
  godowns: Godown[];
  stacks: Stack[];
  organizers: Organizer[];
  features?: { discrepancyWorkflow?: boolean };
};

type OperationalBootstrapPayload = {
  lots: CertificationLot[];
  discrepancies: IntakeDiscrepancy[];
  discrepancyShifts: DiscrepancyShift[];
  stackAccommodations: StackAccommodation[];
  financialVouchers: FinancialVoucher[];
  organizerPayments: OrganizerPayment[];
  receipts: IntakeReceipt[];
  features?: { discrepancyWorkflow?: boolean };
};

type FullBootstrapPayload = CoreBootstrapPayload & OperationalBootstrapPayload;
type AppBootstrapPayload = FullBootstrapPayload & {
  validationSummary?: unknown;
};

const reportTypeOptions: { value: ReportType; label: string }[] = [
  { value: "GODOWN_WISE_DETAIL", label: "Godown Wise Detail" },
  { value: "DISTRICT_WISE_DETAIL", label: "District Wise Detail" },
  { value: "FARMER_WISE_DETAIL", label: "Farmer Wise Detail" },
  { value: "OVERALL_INTAKE", label: "Overall Intake" },
  { value: "SUMMARY", label: "Summary" },
  { value: "DAILY_INTAKE_REGISTER", label: "Daily Intake Register" },
  { value: "CUSTOM_DATE_PAYMENT_REGISTER", label: "Custom Date Payment Register" },
  { value: "ORGANIZER_FARMER_PAYMENT_REGISTER", label: "Organizer Wise Farmer Payment" },
  { value: "ORGANIZER_PAYMENT_TRANSACTION_REPORT", label: "Organizer Payment Transaction Report" },
  { value: "OVERPAID_FARMER_REPORT", label: "Overpaid Farmer Report" },
  { value: "RECEIPT_VOUCHER_TRACEABILITY_REPORT", label: "Receipt to Voucher Traceability" },
  { value: "ORGANIZER_INTAKE_PAYMENT_COMMISSION_REPORT", label: "Organizer Intake vs Payment vs Commission" },
  { value: "REGISTRATION_PENDING_RECEIVED", label: "Registration Pending vs Received" },
  { value: "LOT_WISE_STOCK_LEDGER", label: "Lot-wise Stock Ledger" },
  { value: "ADJUSTED_LOT_FORMATION_REGISTER", label: "Adjusted Lot Formation" },
  { value: "ADJUSTED_LOT_LEDGER_FARMER_WISE", label: "Adjusted Lot Ledger Farmer Wise" },
  { value: "STACK_WISE_STOCK_POSITION", label: "Stack-wise Stock Position" },
  { value: "STACK_CARD_REGISTER", label: "Stack Card Register" },
  { value: "DISCREPANCY_REGISTER", label: "Discrepancy Register" }
];

const defaultGodowns: Godown[] = [
  { id: "godown-1", name: "Main Godown" },
  { id: "godown-2", name: "North Godown" }
];

const defaultStacks: Stack[] = [
  { id: "stack-1", godownId: "godown-1", stackNo: "A-01" },
  { id: "stack-2", godownId: "godown-1", stackNo: "A-02" },
  { id: "stack-3", godownId: "godown-2", stackNo: "B-01" }
];

function createClientId(prefix: string) {
  const webCrypto = typeof globalThis !== "undefined" ? globalThis.crypto : undefined;
  if (webCrypto && typeof webCrypto.randomUUID === "function") {
    return webCrypto.randomUUID();
  }

  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function createReceiptLine(
  godowns: Godown[],
  stacks: Stack[]
): Omit<IntakeReceiptLine, "allocations"> {
  const firstGodown = godowns[0];

  return {
    id: createClientId("line"),
    godownId: firstGodown?.id ?? "",
    stackNo: "",
    grossWeightQtl: 0,
    qtyQtl: 0,
    noOfBags: 0,
    weightPerBagKg: 0,
    netWeightQtl: 0,
    moisturePercent: 0,
    vehicleNo: "",
    remarks: ""
  };
}

function nextReceiptNo(receipts: IntakeReceipt[]): string {
  const max = receipts.reduce((current, item) => {
    const parsed = Number(item.receiptNo);
    return Number.isFinite(parsed) ? Math.max(current, parsed) : current;
  }, 0);
  return String(max + 1).padStart(3, "0");
}

function formatNumber(value: number): string {
  return Number(value.toFixed(2)).toLocaleString("en-IN");
}

function roundQtl(value: number): number {
  return Number(value.toFixed(2));
}

function getVoucherFinalPayable(voucher: FinancialVoucher) {
  return Number(voucher.finalPayableAmount ?? voucher.netPayableAmount ?? 0);
}

function getVoucherTotalPaid(voucher: FinancialVoucher) {
  return Number(
    voucher.totalPaidAmount ??
      voucher.payments?.reduce((sum, item) => sum + Number(item.amount ?? 0), 0) ??
      0
  );
}

function getVoucherBalance(voucher: FinancialVoucher) {
  return Number(
    voucher.balanceAmount ?? Number((getVoucherFinalPayable(voucher) - getVoucherTotalPaid(voucher)).toFixed(2))
  );
}

function isVoucherLockedStatus(status: string) {
  return status === "PAID" || status === "OVERPAID";
}

function formatDateDisplay(value: string) {
  if (!value) {
    return "";
  }

  const [year, month, day] = value.split("-");
  if (!year || !month || !day) {
    return value;
  }

  return `${day}-${month}-${year}`;
}

function drawKrishivPdfLogo(doc: jsPDF, x: number, y: number, scale = 1) {
  doc.setFont("times", "bold");
  doc.setFontSize(19 * scale);
  doc.setTextColor(242, 140, 34);
  doc.text("K", x, y + 11 * scale);
  doc.setTextColor(0, 138, 100);
  doc.rect(x + 5.8 * scale, y + 1.4 * scale, 2.2 * scale, 10.6 * scale, "F");
  doc.setFillColor(139, 197, 64);
  doc.ellipse(x + 11.5 * scale, y + 2.6 * scale, 5.5 * scale, 2.5 * scale, "F");
  doc.setTextColor(242, 140, 34);
  doc.text("ri", x + 13 * scale, y + 11 * scale);
  doc.setTextColor(7, 150, 184);
  doc.text("shiv", x + 27 * scale, y + 11 * scale);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(5.8 * scale);
  doc.setTextColor(45, 45, 45);
  doc.text("SEEDS", x + 44 * scale, y + 16 * scale);
  doc.setTextColor(0, 0, 0);
  doc.setFont("helvetica", "normal");
}

function renderPdfBrandHeader(
  doc: jsPDF,
  title: string,
  options: { left: number; right: number; y: number; compact?: boolean; logoDataUrl?: string }
) {
  const { left, right, compact = false, logoDataUrl = "" } = options;
  const pageWidth = doc.internal.pageSize.getWidth();
  let y = options.y;
  const scale = compact ? 0.86 : 1;
  const logoWidth = (compact ? 43 : 52) * scale;
  const logoHeight = logoWidth / 2.385;

  doc.setDrawColor(91, 61, 38);
  doc.setLineWidth(0.3);
  doc.line(left, y, right, y);
  y += compact ? 2.6 : 3.6;
  if (logoDataUrl) {
    doc.addImage(logoDataUrl, "PNG", pageWidth / 2 - logoWidth / 2, y, logoWidth, logoHeight);
  } else {
    drawKrishivPdfLogo(doc, pageWidth / 2 - 66 * scale / 2, y, scale);
  }
  y += logoHeight + (compact ? 2.8 : 3.6);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(compact ? 7.6 : 8.6);
  doc.setTextColor(91, 61, 38);
  doc.text(COMPANY_NAME, pageWidth / 2, y, { align: "center" });
  y += compact ? 4 : 4.6;
  doc.setFontSize(compact ? 8.6 : 10.2);
  doc.setTextColor(0, 0, 0);
  doc.text(title, pageWidth / 2, y, { align: "center" });
  y += compact ? 3.2 : 4;
  doc.line(left, y, right, y);
  y += compact ? 4.2 : 5.2;
  return y;
}

function truncateToTwoDecimals(value: number): number {
  return Math.floor(value * 100) / 100;
}

function getReceiptRemarks(receipt: IntakeReceipt): string {
  const remarks = receipt.lines
    .map((line) => line.remarks?.trim() ?? "")
    .filter((value) => value.length > 0);

  return remarks.length > 0 ? remarks.join(" | ") : "-";
}

function buildLotLedgerRows(lots: CertificationLot[], receipts: IntakeReceipt[]) {
  const lotLookup = new Map(
    lots.map((lot) => [
      lot.id,
      {
        ...lot,
        displayQtyQtl: lot.currentQtyQtl,
        bags: 0,
        exactBags: 0
      }
    ])
  );

  let totalReceiptBags = 0;

  for (const receipt of receipts) {
    for (const line of receipt.lines) {
      const lineBags = Math.max(0, Math.floor(Number(line.noOfBags ?? 0)));
      const lineNetQty = Number(line.netWeightQtl ?? 0);
      totalReceiptBags += lineBags;

      for (const allocation of line.allocations) {
        const lotRow = lotLookup.get(allocation.lotId);

        if (!lotRow || lineBags <= 0 || lineNetQty <= 0) {
          continue;
        }

        lotRow.exactBags += (lineBags * Number(allocation.qtyQtl ?? 0)) / lineNetQty;
      }
    }
  }

  const computedRows = lots.map((lot) => {
    const computed = lotLookup.get(lot.id);
    const exactBags = computed?.exactBags ?? 0;

    return {
      ...lot,
      exactBags,
      bags: Math.floor(exactBags),
      displayQtyQtl: lot.currentQtyQtl
    };
  });

  let assignedBags = computedRows.reduce((sum, lot) => sum + lot.bags, 0);
  let remainingBagsToAssign = Math.max(totalReceiptBags - assignedBags, 0);

  computedRows
    .slice()
    .sort((left, right) => {
      const leftRemainder = left.exactBags - Math.floor(left.exactBags);
      const rightRemainder = right.exactBags - Math.floor(right.exactBags);
      return rightRemainder - leftRemainder;
    })
    .forEach((lot) => {
      if (remainingBagsToAssign <= 0) {
        return;
      }

      lot.bags += 1;
      remainingBagsToAssign -= 1;
    });

  return computedRows
    .sort((left, right) => left.lotCode.localeCompare(right.lotCode, "en", { numeric: true, sensitivity: "base" }));
}

function groupDiscrepancies(discrepancies: IntakeDiscrepancy[]) {
  const tree = new Map<
    string,
    Map<string, Map<string, Map<string, IntakeDiscrepancy[]>>>
  >();

  for (const discrepancy of discrepancies) {
    const seasonNode = tree.get(discrepancy.season) ?? new Map();
    tree.set(discrepancy.season, seasonNode);
    const godownNode = seasonNode.get(discrepancy.godownName) ?? new Map();
    seasonNode.set(discrepancy.godownName, godownNode);
    const stackNode = godownNode.get(discrepancy.stackNo) ?? new Map();
    godownNode.set(discrepancy.stackNo, stackNode);
    const regNode = stackNode.get(discrepancy.cropRegistrationCode) ?? [];
    stackNode.set(discrepancy.cropRegistrationCode, [...regNode, discrepancy]);
  }

  return tree;
}

function sumReceiptNetQty(receipt: IntakeReceipt) {
  return roundQtl(
    receipt.lines.reduce((sum, line) => sum + Number(line.netWeightQtl ?? line.qtyQtl ?? 0), 0)
  );
}

export default function HomePage() {
  type ValidationSummary = {
    type: string;
    reconciledRegistrations?: number;
    openDiscrepancies?: number;
    resolvedDiscrepancies?: number;
    totalPendingExcessQtyQtl?: number;
    updatedLots?: number;
    orphanLotsRemoved?: number;
    overCapLots?: number;
    totalLotsChecked?: number;
    reindexedLots?: number;
    updatedReceiptAllocations?: number;
    affectedRegistrations?: number;
  } | null;

  const [activeView, setActiveView] = useState<ViewKey>("dashboard");
  const [expandedNavSections, setExpandedNavSections] = useState<Record<SidebarSectionKey, boolean>>({
    overview: true,
    masterData: true,
    intakeOps: true,
    stockLots: true,
    financeOps: true,
    reporting: true,
    administration: false
  });
  const [currentUser, setCurrentUser] = useState<AppUser | null>(null);
  const [currentPermissions, setCurrentPermissions] = useState<RolePermissions | null>(null);
  const [authToken, setAuthToken] = useState("");
  const [pdfLogoDataUrl, setPdfLogoDataUrl] = useState("");
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [registrations, setRegistrations] = useState<RegistrationRecord[]>([]);
  const [godowns, setGodowns] = useState<Godown[]>(defaultGodowns);
  const [stacks, setStacks] = useState<Stack[]>(defaultStacks);
  const [organizers, setOrganizers] = useState<Organizer[]>([]);
  const [lots, setLots] = useState<CertificationLot[]>([]);
  const [discrepancies, setDiscrepancies] = useState<IntakeDiscrepancy[]>([]);
  const [discrepancyShifts, setDiscrepancyShifts] = useState<DiscrepancyShift[]>([]);
  const [stackAccommodations, setStackAccommodations] = useState<StackAccommodation[]>([]);
  const [financialVouchers, setFinancialVouchers] = useState<FinancialVoucher[]>([]);
  const [organizerPayments, setOrganizerPayments] = useState<OrganizerPayment[]>([]);
  const [receipts, setReceipts] = useState<IntakeReceipt[]>([]);
  const [features, setFeatures] = useState({ discrepancyWorkflow: false });
  const [selectedRegistrationId, setSelectedRegistrationId] = useState("");
  const [registrationSearch, setRegistrationSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("ALL");
  const [importSortBy, setImportSortBy] = useState<ImportSortKey>("farmerName");
  const [importPage, setImportPage] = useState(1);
  const [registrationSortBy, setRegistrationSortBy] =
    useState<RegistrationSortKey>("farmerName");
  const [registrationSortDirection, setRegistrationSortDirection] = useState<"asc" | "desc">("asc");
  const [reportSeasonLabel, setReportSeasonLabel] = useState("RABI 2025-26");
  const [reportType, setReportType] = useState<ReportType>("GODOWN_WISE_DETAIL");
  const [reportFromDate, setReportFromDate] = useState("");
  const [reportToDate, setReportToDate] = useState("");
  const [reportCrop, setReportCrop] = useState("");
  const [reportVariety, setReportVariety] = useState("");
  const [reportClassStage, setReportClassStage] = useState("");
  const [reportDistrict, setReportDistrict] = useState("");
  const [reportGodownId, setReportGodownId] = useState("");
  const [reportStackNo, setReportStackNo] = useState("");
  const [reportRegistrationCode, setReportRegistrationCode] = useState("");
  const [reportFarmerName, setReportFarmerName] = useState("");
  const [reportOrganizerName, setReportOrganizerName] = useState("");
  const [reportVillage, setReportVillage] = useState("");
  const [reportPaymentStatus, setReportPaymentStatus] = useState("");
  const [reportMode, setReportMode] = useState<ReportMode>("ALL");
  const [reportPreview, setReportPreview] = useState<ReportPreview | null>(null);
  const [slipSearch, setSlipSearch] = useState("");
  const [slipDistrictFilter, setSlipDistrictFilter] = useState("");
  const [slipVillageFilter, setSlipVillageFilter] = useState("");
  const [slipClassFilter, setSlipClassFilter] = useState("");
  const [slipCropFilter, setSlipCropFilter] = useState("");
  const [slipOnlyWithIntake, setSlipOnlyWithIntake] = useState(true);
  const [slipType, setSlipType] = useState<SlipType>("FARMER_SINGLE_RECEIPT");
  const [slipRegistrationId, setSlipRegistrationId] = useState("");
  const [slipReceiptNo, setSlipReceiptNo] = useState("");
  const [slipDate, setSlipDate] = useState("");
  const [slipPreview, setSlipPreview] = useState<SlipPreview | null>(null);
  const [slipModalOpen, setSlipModalOpen] = useState(false);
  const [voucherGenerationSearch, setVoucherGenerationSearch] = useState("");
  const [voucherGenerationDistrictFilter, setVoucherGenerationDistrictFilter] = useState("");
  const [voucherRegisterSearch, setVoucherRegisterSearch] = useState("");
  const [voucherRegisterDistrictFilter, setVoucherRegisterDistrictFilter] = useState("");
  const [voucherRegisterStatusFilter, setVoucherRegisterStatusFilter] = useState("");
  const [voucherRegisterOrganizerFilter, setVoucherRegisterOrganizerFilter] = useState("");
  const [voucherRegistrationId, setVoucherRegistrationId] = useState("");
  const [voucherModalOpen, setVoucherModalOpen] = useState(false);
  const [voucherDate, setVoucherDate] = useState(new Date().toISOString().slice(0, 10));
  const [certifiedRate, setCertifiedRate] = useState("");
  const [discrepancyRate, setDiscrepancyRate] = useState("");
  const [voucherDeduction, setVoucherDeduction] = useState("0");
  const [voucherAdminPassword, setVoucherAdminPassword] = useState("");
  const [voucherRemarks, setVoucherRemarks] = useState("");
  const [voucherPreview, setVoucherPreview] = useState<VoucherPreview | null>(null);
  const [paymentLedgerVoucherId, setPaymentLedgerVoucherId] = useState("");
  const [paymentDate, setPaymentDate] = useState(new Date().toISOString().slice(0, 10));
  const [paymentAmount, setPaymentAmount] = useState("");
  const [paymentTransactionNo, setPaymentTransactionNo] = useState("");
  const [paymentRemarks, setPaymentRemarks] = useState("");
  const [editingVoucherPaymentId, setEditingVoucherPaymentId] = useState("");
  const [paymentAdminPassword, setPaymentAdminPassword] = useState("");
  const [organizerSearch, setOrganizerSearch] = useState("");
  const [organizerDistrictFilter, setOrganizerDistrictFilter] = useState("");
  const [newOrganizerName, setNewOrganizerName] = useState("");
  const [newOrganizerMobile, setNewOrganizerMobile] = useState("");
  const [newOrganizerVillage, setNewOrganizerVillage] = useState("");
  const [newOrganizerDistrict, setNewOrganizerDistrict] = useState("");
  const [newOrganizerRate, setNewOrganizerRate] = useState("");
  const [newOrganizerDeduction, setNewOrganizerDeduction] = useState("0");
  const [newOrganizerActive, setNewOrganizerActive] = useState(true);
  const [editingOrganizerId, setEditingOrganizerId] = useState("");
  const [organizerAssignmentRegistrationId, setOrganizerAssignmentRegistrationId] = useState("");
  const [organizerAssignmentOrganizerId, setOrganizerAssignmentOrganizerId] = useState("");
  const [organizerLedgerOrganizerId, setOrganizerLedgerOrganizerId] = useState("");
  const [organizerPaymentDate, setOrganizerPaymentDate] = useState(new Date().toISOString().slice(0, 10));
  const [organizerPaymentAmount, setOrganizerPaymentAmount] = useState("");
  const [organizerPaymentTransactionNo, setOrganizerPaymentTransactionNo] = useState("");
  const [organizerPaymentRemarks, setOrganizerPaymentRemarks] = useState("");
  const [editingOrganizerPaymentId, setEditingOrganizerPaymentId] = useState("");
  const [dashboardQuestion, setDashboardQuestion] = useState("");
  const [dashboardAssistantResult, setDashboardAssistantResult] = useState<DashboardAssistantResult | null>(null);
  const [isSavingReceipt, setIsSavingReceipt] = useState(false);
  const [isAddingVoucherPayment, setIsAddingVoucherPayment] = useState(false);
  const [dashboardExpandedSections, setDashboardExpandedSections] = useState<
    Record<string, boolean>
  >({
    districts: false,
    godowns: false,
    organizers: false,
    organizerFarmerPending: false,
    organizerPerformance: false,
    organizerNoIntake: false,
    pendingRegistrations: false,
    stackHotspots: false,
    recentReceipts: false,
    recentVouchers: false
  });
  const [hasLoadedCoreData, setHasLoadedCoreData] = useState(false);
  const [hasLoadedOperationalData, setHasLoadedOperationalData] = useState(false);
  const [isLoadingOperationalData, setIsLoadingOperationalData] = useState(false);
  const operationalLoadInFlightRef = useRef(false);
  const [importMessage, setImportMessage] = useState("Import the farmer master Excel to begin.");
  const [importAdminPassword, setImportAdminPassword] = useState("");
  const [backupDirectory, setBackupDirectory] = useState("mongo-backups");
  const [restoreDirectory, setRestoreDirectory] = useState("");
  const [maintenanceMessage, setMaintenanceMessage] = useState("");
  const [receiptNo, setReceiptNo] = useState("001");
  const [receiptDate, setReceiptDate] = useState(new Date().toISOString().slice(0, 10));
  const [draftLines, setDraftLines] = useState<Omit<IntakeReceiptLine, "allocations">[]>(() =>
    [createReceiptLine(defaultGodowns, defaultStacks)]
  );
  const [newGodownName, setNewGodownName] = useState("");
  const [newStackGodownId, setNewStackGodownId] = useState(defaultGodowns[0].id);
  const [newStackNo, setNewStackNo] = useState("");
  const [toast, setToast] = useState("");
  const [editingReceiptNo, setEditingReceiptNo] = useState("");
  const [depositViewRegistrationId, setDepositViewRegistrationId] = useState("");
  const [selectedDiscrepancyId, setSelectedDiscrepancyId] = useState("");
  const [shiftTargetGodownId, setShiftTargetGodownId] = useState(defaultGodowns[0].id);
  const [shiftTargetStackNo, setShiftTargetStackNo] = useState("");
  const [accommodationTargetRegistrationId, setAccommodationTargetRegistrationId] = useState("");
  const [accommodationQtyQtl, setAccommodationQtyQtl] = useState("");
  const [accommodationBags, setAccommodationBags] = useState("");
  const [accommodationDate, setAccommodationDate] = useState(new Date().toISOString().slice(0, 10));
  const [accommodationRemarks, setAccommodationRemarks] = useState("");
  const [editingAccommodationId, setEditingAccommodationId] = useState("");
  const [discrepancyWorkflowMode, setDiscrepancyWorkflowMode] = useState<"accommodation" | "shift">(
    "accommodation"
  );
  const [shiftQtyQtl, setShiftQtyQtl] = useState("");
  const [shiftBags, setShiftBags] = useState("");
  const [shiftDate, setShiftDate] = useState(new Date().toISOString().slice(0, 10));
  const [shiftApprovedBy, setShiftApprovedBy] = useState("");
  const [shiftRemarks, setShiftRemarks] = useState("");
  const [validationSummary, setValidationSummary] = useState<ValidationSummary>(null);

  function authHeaders(extraHeaders?: Record<string, string>) {
    return {
      ...(authToken
        ? {
            Authorization: `Bearer ${authToken}`
          }
        : {}),
      ...(extraHeaders ?? {})
    };
  }

  async function fetchWithAuth(input: RequestInfo | URL, init?: RequestInit) {
    const headers = authHeaders((init?.headers as Record<string, string> | undefined) ?? undefined);
    return fetch(input, {
      ...init,
      headers
    });
  }

  function applyCoreBootstrap(data: CoreBootstrapPayload) {
    setRegistrations(data.registrations ?? []);
    setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
    setStacks(data.stacks?.length ? data.stacks : defaultStacks);
    setOrganizers(data.organizers ?? []);
    setFeatures({
      discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
    });
    setSelectedRegistrationId((current) => current || data.registrations?.[0]?.id || "");
    setSlipRegistrationId((current) => current || data.registrations?.[0]?.id || "");
    setHasLoadedCoreData(true);
  }

  function applyOperationalBootstrap(data: OperationalBootstrapPayload) {
    setLots(data.lots ?? []);
    setDiscrepancies(data.discrepancies ?? []);
    setDiscrepancyShifts(data.discrepancyShifts ?? []);
    setStackAccommodations(data.stackAccommodations ?? []);
    setFinancialVouchers(data.financialVouchers ?? []);
    setOrganizerPayments(data.organizerPayments ?? []);
    setReceipts(data.receipts ?? []);
    setFeatures((current) => ({
      ...current,
      discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow ?? current.discrepancyWorkflow)
    }));
    setSlipDate((current) => current || data.receipts?.[0]?.receiptDate || "");
    setReceiptNo(nextReceiptNo(data.receipts ?? []));
    setHasLoadedOperationalData(true);
  }

  function applyFullBootstrap(data: FullBootstrapPayload) {
    applyCoreBootstrap(data);
    applyOperationalBootstrap(data);
  }

  function notifyUser(message: string, useDialog = true) {
    setToast(message);
    if (useDialog && typeof window !== "undefined") {
      window.alert(message);
    }
  }

  function confirmDestructiveAction({
    itemLabel
  }: {
    itemLabel: string;
  }) {
    if (typeof window === "undefined") {
      return false;
    }

    return window.confirm(`Do you want to process the deletion?\n\n${itemLabel}`);
  }

  function requirePermission(permission: keyof RolePermissions, message: string) {
    if (!effectivePermissions?.[permission]) {
      notifyUser(message);
      return false;
    }
    return true;
  }

  async function handleLogin() {
    setLoginError("");
    const response = await fetch(`${API_BASE}/api/auth/login`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        email: loginEmail.trim(),
        password: loginPassword
      })
    });

    if (!response.ok) {
      const errorBody = await response.json().catch(() => null);
      throw new Error(errorBody?.message || "Invalid login.");
    }

    const data = (await response.json()) as {
      user: AppUser;
      permissions: RolePermissions;
      token: string;
    };
    setCurrentUser(data.user);
    setCurrentPermissions(data.permissions);
    setAuthToken(data.token);
    if (typeof window !== "undefined") {
      window.localStorage.setItem(
        "krishiv-auth",
        JSON.stringify({
          user: data.user,
          permissions: data.permissions,
          token: data.token
        })
      );
    }
    setToast(`Logged in as ${data.user.role}.`);
  }

  function handleLogout() {
    setCurrentUser(null);
    setCurrentPermissions(null);
    setAuthToken("");
    setLoginError("");
    setVoucherModalOpen(false);
    setSlipModalOpen(false);
    setHasLoadedCoreData(false);
    setHasLoadedOperationalData(false);
    setOrganizers([]);
    setLots([]);
    setDiscrepancies([]);
    setDiscrepancyShifts([]);
    setStackAccommodations([]);
    setFinancialVouchers([]);
    setOrganizerPayments([]);
    setReceipts([]);
    if (typeof window !== "undefined") {
      window.localStorage.removeItem("krishiv-auth");
    }
  }

  async function loadCoreBootstrap() {
    const response = await fetchWithAuth(`${API_BASE}/api/seed/bootstrap/core`);
    if (!response.ok) {
      throw new Error("Unable to load seed data from API.");
    }

    const data = (await response.json()) as CoreBootstrapPayload;
    applyCoreBootstrap(data);
  }

  async function loadOperationalBootstrap() {
    if (operationalLoadInFlightRef.current) {
      return;
    }

    operationalLoadInFlightRef.current = true;
    setIsLoadingOperationalData(true);
    try {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/bootstrap/operations`);
      if (!response.ok) {
        throw new Error("Unable to load operational data from API.");
      }

      const data = (await response.json()) as OperationalBootstrapPayload;
      applyOperationalBootstrap(data);
    } finally {
      operationalLoadInFlightRef.current = false;
      setIsLoadingOperationalData(false);
    }
  }

  async function loadBootstrap() {
    const response = await fetchWithAuth(`${API_BASE}/api/seed/bootstrap`);
    if (!response.ok) {
      throw new Error("Unable to load seed data from API.");
    }

    const data = (await response.json()) as FullBootstrapPayload;
    applyFullBootstrap(data);
  }

  useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }
    void fetch(BRAND_LOGO_SRC)
      .then((response) => (response.ok ? response.text() : ""))
      .then((svgText) => {
        const matchedLogo = svgText.match(/xlink:href="(data:image\/png;base64,[^"]+)"/i);
        if (matchedLogo?.[1]) {
          setPdfLogoDataUrl(matchedLogo[1]);
        }
      })
      .catch(() => {
        setPdfLogoDataUrl("");
      });

    const raw = window.localStorage.getItem("krishiv-auth");
    if (!raw) {
      return;
    }
    try {
      const parsed = JSON.parse(raw) as { user: AppUser; permissions: RolePermissions; token?: string };
      if (!parsed.token) {
        window.localStorage.removeItem("krishiv-auth");
        return;
      }
      setCurrentUser(parsed.user);
      setCurrentPermissions(parsed.permissions);
      setAuthToken(parsed.token);
    } catch {
      window.localStorage.removeItem("krishiv-auth");
    }
  }, []);

  useEffect(() => {
    if (!currentUser) {
      return;
    }
    void loadCoreBootstrap()
      .then(() => {
        void loadOperationalBootstrap().catch((error) => {
          setToast(error instanceof Error ? error.message : "Unable to load operational data.");
        });
      })
      .catch((error) => {
      setToast(error instanceof Error ? error.message : "Unable to connect to API.");
      setGodowns(defaultGodowns);
      setStacks(defaultStacks);
      });
  }, [currentUser]);

  useEffect(() => {
    if (!currentUser || hasLoadedOperationalData || isLoadingOperationalData) {
      return;
    }

    const viewsNeedingOperationalData: ViewKey[] = [
      "dashboard",
      "registrations",
      "intake",
      "intakeEdit",
      "reports",
      "finance",
      "commission",
      "slips",
      "lots",
      "discrepancies",
      "validations",
      "backup",
      "restore"
    ];

    if (!viewsNeedingOperationalData.includes(activeView)) {
      return;
    }

    void loadOperationalBootstrap().catch((error) => {
      setToast(error instanceof Error ? error.message : "Unable to load operational data.");
    });
  }, [activeView, currentUser, hasLoadedOperationalData, isLoadingOperationalData]);

  const selectedRegistration = registrations.find((item) => item.id === selectedRegistrationId);
  const editingReceipt = receipts.find((item) => item.receiptNo === editingReceiptNo);
  const editingReceiptQtyQtl = editingReceipt
    ? editingReceipt.lines.reduce((sum, line) => sum + Number(line.qtyQtl ?? 0), 0)
    : 0;
  const availableBalanceForSave = selectedRegistration
    ? Number(selectedRegistration.balanceQtl) + editingReceiptQtyQtl
    : 0;

  const filteredRegistrations = registrations.filter((item) => {
    if (item.status === "BLOCKED") {
      return false;
    }
    const matchesSearch =
      item.cropRegistrationCode.toLowerCase().includes(registrationSearch.toLowerCase()) ||
      item.farmerName.toLowerCase().includes(registrationSearch.toLowerCase()) ||
      item.village.toLowerCase().includes(registrationSearch.toLowerCase());
    const matchesStatus = statusFilter === "ALL" || item.status === statusFilter;
    return matchesSearch && matchesStatus;
  });
  const sortedRegistrationRows = filteredRegistrations.slice().sort((left, right) => {
    const leftValue = left[registrationSortBy];
    const rightValue = right[registrationSortBy];
    const direction = registrationSortDirection === "asc" ? 1 : -1;

    if (typeof leftValue === "number" && typeof rightValue === "number") {
      return (leftValue - rightValue) * direction;
    }

    return String(leftValue).localeCompare(String(rightValue), "en", {
      sensitivity: "base"
    }) * direction;
  });

  const sortedImportRegistrations = registrations
    .slice()
    .sort((left, right) => left[importSortBy].localeCompare(right[importSortBy], "en", { sensitivity: "base" }));
  const importPageCount = Math.max(1, Math.ceil(sortedImportRegistrations.length / IMPORT_PAGE_SIZE));
  const safeImportPage = Math.min(importPage, importPageCount);
  const paginatedImportRegistrations = sortedImportRegistrations.slice(
    (safeImportPage - 1) * IMPORT_PAGE_SIZE,
    safeImportPage * IMPORT_PAGE_SIZE
  );
  const lotLedgerRows = buildLotLedgerRows(lots, receipts);
  const discrepancyTree = groupDiscrepancies(discrepancies);
  const openDiscrepancies = discrepancies.filter((item) => item.status !== "RESOLVED");
  const dashboardMetrics = {
    expectedYield: registrations.reduce((sum, item) => sum + item.expectedYieldQtl, 0),
    grossReceived: receipts.reduce(
      (sum, receipt) =>
        sum +
        receipt.lines.reduce(
          (lineSum, line) => lineSum + Number(line.grossWeightQtl ?? 0),
          0
        ),
      0
    ),
    netReceived: receipts.reduce(
      (sum, receipt) =>
        sum +
        receipt.lines.reduce(
          (lineSum, line) => lineSum + Number(line.netWeightQtl ?? line.qtyQtl ?? 0),
          0
        ),
      0
    ),
    pending: registrations.reduce((sum, item) => sum + item.balanceQtl, 0),
    activeRegistrations: registrations.filter((item) => item.status === "ACTIVE").length,
    openLots: lots.filter((item) => item.status === "OPEN").length,
    fullLots: lots.filter((item) => item.status === "FULL").length,
    totalLots: lots.length,
    intakeBags: receipts.reduce(
      (sum, receipt) =>
        sum + receipt.lines.reduce((lineSum, line) => lineSum + Number(line.noOfBags ?? 0), 0),
      0
    ),
    discrepancyQty: openDiscrepancies.reduce((sum, item) => sum + item.excessQtyQtl, 0),
    discrepancyBags: openDiscrepancies.reduce((sum, item) => sum + item.estimatedExcessBags, 0),
    discrepancyCount: openDiscrepancies.length,
    shiftedQty: discrepancyShifts.reduce((sum, item) => sum + item.shiftedQtyQtl, 0),
    shiftedBags: discrepancyShifts.reduce((sum, item) => sum + item.shiftedBags, 0),
    shiftedCases: discrepancyShifts.length
  };
  const todayIso = new Date().toISOString().slice(0, 10);
  const intakeCoveragePct =
    dashboardMetrics.expectedYield > 0
      ? Math.min((dashboardMetrics.netReceived / dashboardMetrics.expectedYield) * 100, 100)
      : 0;
  const discrepancyExposurePct =
    dashboardMetrics.netReceived > 0
      ? (dashboardMetrics.discrepancyQty / dashboardMetrics.netReceived) * 100
      : 0;
  const totalLotCapacityQtl = lots.length * 200;
  const lotUtilizationPct =
    totalLotCapacityQtl > 0
      ? Math.min(
          (lots.reduce((sum, item) => sum + Number(item.currentQtyQtl ?? 0), 0) / totalLotCapacityQtl) * 100,
          100
        )
      : 0;
  const averageNetPerReceipt =
    receipts.length > 0 ? dashboardMetrics.netReceived / receipts.length : 0;
  const averageWeightPerBagKg =
    dashboardMetrics.intakeBags > 0
      ? (dashboardMetrics.netReceived * 100) / dashboardMetrics.intakeBags
      : 0;
  const draftVoucherCount = financialVouchers.filter((item) => item.status === "DRAFT").length;
  const paidVoucherCount = financialVouchers.filter((item) =>
    ["PAID", "OVERPAID"].includes(item.status)
  ).length;
  const unpaidVoucherCount = financialVouchers.filter(
    (item) => !["PAID", "OVERPAID"].includes(item.status)
  ).length;
  const voucherBalanceOutstanding = financialVouchers.reduce(
    (sum, item) => sum + Number(item.balanceAmount ?? 0),
    0
  );
  const todayReceipts = receipts.filter((item) => item.receiptDate === todayIso);
  const todayReceiptBags = todayReceipts.reduce(
    (sum, receipt) =>
      sum + receipt.lines.reduce((lineSum, line) => lineSum + Number(line.noOfBags ?? 0), 0),
    0
  );
  const todayReceiptNet = todayReceipts.reduce((sum, receipt) => sum + sumReceiptNetQty(receipt), 0);
  const isPaymentRegisterReport = reportType === "CUSTOM_DATE_PAYMENT_REGISTER";
  const isOrganizerFarmerPaymentReport = reportType === "ORGANIZER_FARMER_PAYMENT_REGISTER";
  const isOrganizerPaymentTransactionReport =
    reportType === "ORGANIZER_PAYMENT_TRANSACTION_REPORT";
  const isOverpaidFarmerReport = reportType === "OVERPAID_FARMER_REPORT";
  const isReceiptVoucherTraceabilityReport = reportType === "RECEIPT_VOUCHER_TRACEABILITY_REPORT";
  const isOrganizerIntakePaymentCommissionReport =
    reportType === "ORGANIZER_INTAKE_PAYMENT_COMMISSION_REPORT";
  const isAdjustedLotFormationReport = reportType === "ADJUSTED_LOT_FORMATION_REGISTER";
  const isAdjustedLotLedgerFarmerWiseReport = reportType === "ADJUSTED_LOT_LEDGER_FARMER_WISE";
  const isStackCardRegisterReport = reportType === "STACK_CARD_REGISTER";
  const paymentRegisterDistrictOptions = Array.from(
    new Set(financialVouchers.map((item) => item.district?.trim()).filter(Boolean))
  ).sort((left, right) => String(left).localeCompare(String(right), "en", { sensitivity: "base" }));
  const paymentRegisterVillageOptions = Array.from(
    new Set(
      financialVouchers
        .filter((item) => !reportDistrict || item.district === reportDistrict)
        .map((item) => item.village?.trim())
        .filter(Boolean)
    )
  ).sort((left, right) => String(left).localeCompare(String(right), "en", { sensitivity: "base" }));
  const paymentRegisterFarmerOptions = Array.from(
    new Set(
      financialVouchers
        .filter((item) => !reportDistrict || item.district === reportDistrict)
        .filter((item) => !reportVillage || item.village === reportVillage)
        .map((item) => item.farmerName?.trim())
        .filter(Boolean)
    )
  ).sort((left, right) => String(left).localeCompare(String(right), "en", { sensitivity: "base" }));
  const paymentRegisterStatusOptions = Array.from(
    new Set(financialVouchers.map((item) => item.status?.trim()).filter(Boolean))
  ).sort((left, right) => String(left).localeCompare(String(right), "en", { sensitivity: "base" }));
  const organizerById = new Map(organizers.map((item) => [item.id, item]));
  const organizerDistrictOptions = Array.from(
    new Set(organizers.map((item) => item.district?.trim()).filter(Boolean))
  ).sort((left, right) => String(left).localeCompare(String(right), "en", { sensitivity: "base" }));
  const filteredOrganizerMasterRows = organizers
    .filter((organizer) => {
      const query = organizerSearch.trim().toLowerCase();
      const matchesSearch =
        !query ||
        organizer.name.toLowerCase().includes(query) ||
        organizer.mobile.toLowerCase().includes(query) ||
        organizer.village.toLowerCase().includes(query) ||
        organizer.district.toLowerCase().includes(query);
      const matchesDistrict = !organizerDistrictFilter || organizer.district === organizerDistrictFilter;
      return matchesSearch && matchesDistrict;
    })
    .slice()
    .sort((left, right) => left.name.localeCompare(right.name, "en", { sensitivity: "base" }));
  const organizerCommissionRows = organizers
    .map((organizer) => {
      const linkedRegistrations = registrations.filter((item) => item.organizerId === organizer.id);
      const totalIntakeQtl = roundQtl(
        linkedRegistrations.reduce((sum, item) => sum + Number(item.totalReceivedQtl ?? 0), 0)
      );
      const ratePerQtl = roundQtl(Number(organizer.commissionRatePerQtl ?? 0));
      const grossCommissionAmount = roundQtl(
        linkedRegistrations.reduce(
          (sum, item) =>
            sum + Number(item.totalReceivedQtl ?? 0) * Number(item.organizerCommissionRatePerQtl ?? ratePerQtl),
          0
        )
      );
      const deductionAmount = roundQtl(Number(organizer.deductionAmount ?? 0));
      const netPayableAmount = roundQtl(grossCommissionAmount - deductionAmount);
      const payments = organizerPayments.filter((item) => item.organizerId === organizer.id);
      const paidAmount = roundQtl(payments.reduce((sum, item) => sum + Number(item.amount ?? 0), 0));
      return {
        organizer,
        linkedRegistrations,
        farmerCount: linkedRegistrations.length,
        totalIntakeQtl,
        ratePerQtl,
        grossCommissionAmount,
        deductionAmount,
        netPayableAmount,
        paidAmount,
        balanceAmount: roundQtl(netPayableAmount - paidAmount),
        paymentCount: payments.length
      };
    })
    .filter((row) => {
      const query = organizerSearch.trim().toLowerCase();
      const matchesSearch =
        !query ||
        row.organizer.name.toLowerCase().includes(query) ||
        row.organizer.mobile.toLowerCase().includes(query) ||
        row.organizer.village.toLowerCase().includes(query) ||
        row.organizer.district.toLowerCase().includes(query);
      const matchesDistrict = !organizerDistrictFilter || row.organizer.district === organizerDistrictFilter;
      return matchesSearch && matchesDistrict;
    })
    .sort((left, right) => left.organizer.name.localeCompare(right.organizer.name, "en", { sensitivity: "base" }));
  const organizerDashboardRows: OrganizerDashboardRow[] = organizerCommissionRows
    .map((row) => ({
      ...row,
      coveragePct:
        row.netPayableAmount > 0
          ? roundQtl((row.paidAmount / row.netPayableAmount) * 100)
          : 0
    }))
    .sort((left, right) => right.balanceAmount - left.balanceAmount);
  const visibleOrganizerDashboardRows = dashboardExpandedSections.organizers
    ? organizerDashboardRows
    : organizerDashboardRows.slice(0, 5);
  const organizerLinkedFarmerCount = registrations.filter((item) => Boolean(item.organizerId)).length;
  const organizerGrossCommission = roundQtl(
    organizerCommissionRows.reduce((sum, row) => sum + row.grossCommissionAmount, 0)
  );
  const farmerGrossPayableTotal = roundQtl(
    financialVouchers.reduce((sum, voucher) => sum + Number(voucher.grossPayableAmount ?? 0), 0)
  );
  const averagePurchaseRate =
    dashboardMetrics.netReceived > 0
      ? roundQtl((farmerGrossPayableTotal + organizerGrossCommission) / dashboardMetrics.netReceived)
      : 0;
  const organizerDeductionTotal = roundQtl(
    organizerCommissionRows.reduce((sum, row) => sum + row.deductionAmount, 0)
  );
  const organizerNetPayableTotal = roundQtl(
    organizerCommissionRows.reduce((sum, row) => sum + row.netPayableAmount, 0)
  );
  const organizerPaidTotal = roundQtl(
    organizerCommissionRows.reduce((sum, row) => sum + row.paidAmount, 0)
  );
  const organizerBalanceTotal = roundQtl(
    organizerCommissionRows.reduce((sum, row) => sum + row.balanceAmount, 0)
  );
  const organizerPendingCount = organizerCommissionRows.filter((row) => row.balanceAmount > 0).length;
  const organizerNoIntakeRows = organizerCommissionRows
    .map((row) => {
      const zeroIntakeRegistrations = row.linkedRegistrations
        .filter((registration) => Number(registration.totalReceivedQtl ?? 0) <= 0)
        .slice()
        .sort((left, right) => Number(right.balanceQtl ?? 0) - Number(left.balanceQtl ?? 0));
      return {
        ...row,
        zeroIntakeRegistrations,
        zeroIntakeFarmerCount: zeroIntakeRegistrations.length,
        zeroIntakePendingQty: roundQtl(
          zeroIntakeRegistrations.reduce((sum, registration) => sum + Number(registration.balanceQtl ?? 0), 0)
        ),
        farmerPreview: zeroIntakeRegistrations
          .slice(0, 3)
          .map((registration) => registration.farmerName)
          .join(", ")
      };
    })
    .filter((row) => row.zeroIntakeFarmerCount > 0)
    .sort((left, right) => right.zeroIntakePendingQty - left.zeroIntakePendingQty);
  const visibleOrganizerNoIntakeRows = dashboardExpandedSections.organizerNoIntake
    ? organizerNoIntakeRows
    : organizerNoIntakeRows.slice(0, 5);
  const organizerNoIntakeFarmerCount = organizerNoIntakeRows.reduce(
    (sum, row) => sum + row.zeroIntakeFarmerCount,
    0
  );
  const organizerNoIntakePendingQty = roundQtl(
    organizerNoIntakeRows.reduce((sum, row) => sum + row.zeroIntakePendingQty, 0)
  );
  const organizerFarmerPaymentPendingRows: OrganizerFarmerPaymentPendingRow[] = Array.from(
    financialVouchers.reduce((map, voucher) => {
      const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
      const organizerId = String(registration?.organizerId ?? "").trim() || "direct-farmer";
      const organizerName = String(registration?.organizerName ?? "").trim() || "Direct Farmer";
      const organizerRecord = organizerById.get(organizerId);
      const current = map.get(organizerId) ?? {
        organizerId,
        organizerName: organizerRecord?.name || organizerName,
        district: organizerRecord?.district || registration?.district || voucher.district || "-",
        farmerCodes: new Set<string>(),
        voucherCount: 0,
        netPayableAmount: 0,
        paidAmount: 0,
        pendingAmount: 0,
        overpaidCount: 0
      };

      current.farmerCodes.add(voucher.cropRegistrationCode);
      current.voucherCount += 1;
      current.netPayableAmount = roundQtl(current.netPayableAmount + getVoucherFinalPayable(voucher));
      current.paidAmount = roundQtl(current.paidAmount + getVoucherTotalPaid(voucher));
      current.pendingAmount = roundQtl(current.pendingAmount + getVoucherBalance(voucher));
      if (getVoucherBalance(voucher) < 0 || voucher.status === "OVERPAID") {
        current.overpaidCount += 1;
      }

      map.set(organizerId, current);
      return map;
    }, new Map<
      string,
      {
        organizerId: string;
        organizerName: string;
        district: string;
        farmerCodes: Set<string>;
        voucherCount: number;
        netPayableAmount: number;
        paidAmount: number;
        pendingAmount: number;
        overpaidCount: number;
      }
    >())
      .values()
  )
    .map((row) => {
      const netPayableAmount = roundQtl(row.netPayableAmount);
      const paidAmount = roundQtl(row.paidAmount);
      return {
        organizerId: row.organizerId,
        organizerName: row.organizerName,
        district: row.district,
        farmerCount: row.farmerCodes.size,
        voucherCount: row.voucherCount,
        netPayableAmount,
        paidAmount,
        pendingAmount: roundQtl(netPayableAmount - paidAmount),
        paymentCompletionPct: netPayableAmount > 0 ? roundQtl((paidAmount / netPayableAmount) * 100) : 0,
        overpaidCount: row.overpaidCount
      };
    })
    .sort((left, right) => {
      if (right.pendingAmount !== left.pendingAmount) {
        return right.pendingAmount - left.pendingAmount;
      }
      return left.organizerName.localeCompare(right.organizerName, "en", { sensitivity: "base" });
    });
  const visibleOrganizerFarmerPaymentPendingRows = dashboardExpandedSections.organizerFarmerPending
    ? organizerFarmerPaymentPendingRows
    : organizerFarmerPaymentPendingRows.slice(0, 5);
  const organizerFarmerPaymentPendingTotal = roundQtl(
    organizerFarmerPaymentPendingRows.reduce((sum, row) => sum + row.pendingAmount, 0)
  );
  const organizerFarmerPaymentPaidTotal = roundQtl(
    organizerFarmerPaymentPendingRows.reduce((sum, row) => sum + row.paidAmount, 0)
  );
  const organizerFarmerPaymentNetTotal = roundQtl(
    organizerFarmerPaymentPendingRows.reduce((sum, row) => sum + row.netPayableAmount, 0)
  );
  const organizerPerformanceRows: OrganizerPerformanceRow[] = Array.from(
    registrations
      .reduce((map, registration) => {
        const organizerId = String(registration.organizerId ?? "").trim() || "direct-farmer";
        const organizerName = String(registration.organizerName ?? "").trim() || "Direct Farmer";
        const organizerRecord = organizerById.get(organizerId);
        const current = map.get(organizerId) ?? {
          organizerId,
          organizerName: organizerRecord?.name || organizerName,
          district: organizerRecord?.district || registration.district || "-",
          farmerCount: 0,
          expectedYieldQtl: 0,
          depositedQtl: 0
        };
        current.farmerCount += 1;
        current.expectedYieldQtl += Number(registration.expectedYieldQtl ?? 0);
        current.depositedQtl += Number(registration.totalReceivedQtl ?? 0);
        if (!current.district || current.district === "-") {
          current.district = organizerRecord?.district || registration.district || "-";
        }
        map.set(organizerId, current);
        return map;
      }, new Map<string, { organizerId: string; organizerName: string; district: string; farmerCount: number; expectedYieldQtl: number; depositedQtl: number }>())
      .values()
  )
    .map((row) => {
      const expectedYieldQtl = roundQtl(row.expectedYieldQtl);
      const depositedQtl = roundQtl(row.depositedQtl);
      const pendingQtl = roundQtl(Math.max(expectedYieldQtl - depositedQtl, 0));
      return {
        organizerId: row.organizerId,
        organizerName: row.organizerName,
        district: row.district,
        farmerCount: row.farmerCount,
        expectedYieldQtl,
        depositedQtl,
        pendingQtl,
        coveragePct: expectedYieldQtl > 0 ? roundQtl((depositedQtl / expectedYieldQtl) * 100) : 0
      };
    })
    .sort((left, right) => {
      if (right.depositedQtl !== left.depositedQtl) {
        return right.depositedQtl - left.depositedQtl;
      }
      return left.organizerName.localeCompare(right.organizerName, "en", { sensitivity: "base" });
    });
  const visibleOrganizerPerformanceRows = dashboardExpandedSections.organizerPerformance
    ? organizerPerformanceRows
    : organizerPerformanceRows.slice(0, 5);
  const organizerExpectedYieldTotal = roundQtl(
    organizerPerformanceRows.reduce((sum, row) => sum + row.expectedYieldQtl, 0)
  );
  const organizerDepositedTotal = roundQtl(
    organizerPerformanceRows.reduce((sum, row) => sum + row.depositedQtl, 0)
  );
  const organizerPendingYieldTotal = roundQtl(
    organizerPerformanceRows.reduce((sum, row) => sum + row.pendingQtl, 0)
  );
  const organizerReportOptions = Array.from(
    new Set(
      [
        ...organizers
          .filter((organizer) => !reportDistrict || organizer.district === reportDistrict)
          .map((organizer) => organizer.name.trim()),
        ...financialVouchers
          .filter((voucher) => !reportDistrict || voucher.district === reportDistrict)
          .map((voucher) => {
            const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
            return registration?.organizerName?.trim() || "Direct Farmer";
          })
      ].filter(Boolean)
    )
  ).sort((left, right) => left.localeCompare(right, "en", { sensitivity: "base" }));
  const organizerAssignmentRegistration =
    registrations.find((item) => item.id === organizerAssignmentRegistrationId) ?? null;
  const organizerLedgerSummary =
    organizerCommissionRows.find((item) => item.organizer.id === organizerLedgerOrganizerId) ?? null;
  const organizerLedgerPayments = organizerPayments
    .filter((item) => item.organizerId === organizerLedgerOrganizerId)
    .slice()
    .sort((left, right) => {
      const dateCompare = left.paymentDate.localeCompare(right.paymentDate);
      if (dateCompare !== 0) {
        return dateCompare;
      }
      return left.transactionNo.localeCompare(right.transactionNo, "en", { sensitivity: "base" });
    });
  const visibleNavSections = navSections
    .map((section) => ({
      ...section,
      items: section.items.filter(
        (item) => item.key !== "discrepancies" || features.discrepancyWorkflow
      )
    }))
    .filter((section) => section.items.length > 0);
  const visibleNavItems = visibleNavSections.flatMap((section) => section.items);
  const isNavItemDisabled = (key: ViewKey) =>
    ((key === "backup" || key === "restore" || key === "masters") && !isAdminUser) ||
    ((key === "import") && !effectivePermissions?.canImport);
  const activeNavSectionKey =
    visibleNavSections.find((section) => section.items.some((item) => item.key === activeView))?.key ?? "overview";
  const activeNavItem = visibleNavItems.find((item) => item.key === activeView) ?? null;
  function toggleNavSection(sectionKey: SidebarSectionKey) {
    setExpandedNavSections((current) => ({
      ...current,
      [sectionKey]: !current[sectionKey]
    }));
  }
  function openViewFromNav(view: ViewKey, sectionKey: SidebarSectionKey) {
    setExpandedNavSections((current) => ({
      ...current,
      [sectionKey]: true
    }));
    setActiveView(view);
  }
  function toggleDashboardSection(sectionKey: keyof typeof dashboardExpandedSections) {
    setDashboardExpandedSections((current) => ({
      ...current,
      [sectionKey]: !current[sectionKey]
    }));
  }
  function openRegistrationFromDashboard(registrationId: string) {
    setDepositViewRegistrationId(registrationId);
  }
  function openReceiptFromDashboard(receiptId: string) {
    const receipt = receipts.find((item) => item.id === receiptId);
    if (!receipt) {
      return;
    }
    setDepositViewRegistrationId(receipt.cropRegistrationId);
  }
  function openVoucherFromDashboard(voucherId: string) {
    const voucher = financialVouchers.find((item) => item.id === voucherId);
    if (!voucher) {
      return;
    }
    openPaymentLedger(voucher);
  }
  function openStackHotspotFromDashboard(godownName: string, stackNo: string) {
    const target = openDiscrepancies.find(
      (item) => item.godownName === godownName && item.stackNo === stackNo
    );
    if (!target) {
      return;
    }
    setSelectedDiscrepancyId(target.id);
    setActiveView("discrepancies");
  }
  const openDiscrepancyKeySet = new Set(
    openDiscrepancies.map((item) => `${item.cropRegistrationCode}::${item.stackNo}`)
  );
  const selectedDiscrepancy =
    discrepancies.find((item) => item.id === selectedDiscrepancyId) ?? openDiscrepancies[0] ?? null;
  const selectedDiscrepancyAccommodations = selectedDiscrepancy
    ? stackAccommodations
        .filter((item) => item.discrepancyId === selectedDiscrepancy.id)
        .slice()
        .sort((left, right) =>
          `${left.adjustmentDate}-${left.createdAt}`.localeCompare(`${right.adjustmentDate}-${right.createdAt}`)
        )
    : [];
  const allStackAccommodationRows = stackAccommodations
    .slice()
    .sort((left, right) =>
      `${right.adjustmentDate}-${right.createdAt}`.localeCompare(`${left.adjustmentDate}-${left.createdAt}`)
    );
  const selectedStackAccommodationRows = selectedDiscrepancy
    ? allStackAccommodationRows
        .filter(
          (item) =>
            item.godownId === selectedDiscrepancy.godownId &&
            String(item.stackNo ?? "").trim() === String(selectedDiscrepancy.stackNo ?? "").trim()
        )
        .slice()
        .sort((left, right) =>
          `${left.adjustmentDate}-${left.createdAt}`.localeCompare(`${right.adjustmentDate}-${right.createdAt}`)
        )
    : [];
  const adjustedStackCardPreview = selectedDiscrepancy
    ? (() => {
        const originalRows = Array.from(
          receipts.reduce(
            (map, receipt) => {
              const sameStackLines = receipt.lines.filter(
                (line) =>
                  line.godownId === selectedDiscrepancy.godownId &&
                  String(line.stackNo ?? "").trim() === String(selectedDiscrepancy.stackNo ?? "").trim()
              );
              if (!sameStackLines.length) {
                return map;
              }
              const registration = registrations.find((item) => item.id === receipt.cropRegistrationId);
              const regCode = receipt.cropRegistrationCode;
              const current = map.get(regCode) ?? {
                regCode,
                farmerName: receipt.farmerName,
                village: registration?.village || "-",
                district: registration?.district || "-",
                qtyQtl: 0,
                bags: 0
              };
              current.qtyQtl = roundQtl(
                current.qtyQtl +
                  sameStackLines.reduce((sum, line) => sum + Number(line.netWeightQtl ?? line.qtyQtl ?? 0), 0)
              );
              current.bags += sameStackLines.reduce((sum, line) => sum + Number(line.noOfBags ?? 0), 0);
              map.set(regCode, current);
              return map;
            },
            new Map<
              string,
              {
                regCode: string;
                farmerName: string;
                village: string;
                district: string;
                qtyQtl: number;
                bags: number;
              }
            >()
          ).values()
        ).sort((left, right) => left.regCode.localeCompare(right.regCode, "en", { numeric: true }));

        const adjustedRows = originalRows.map((row) => {
          const adjustedInRows = selectedStackAccommodationRows.filter(
            (item) => item.targetRegistrationCode === row.regCode
          );
          const adjustedOutRows = selectedStackAccommodationRows.filter(
            (item) => item.sourceRegistrationCode === row.regCode
          );
          const adjustedInQtyQtl = roundQtl(
            adjustedInRows.reduce((sum, item) => sum + Number(item.adjustedQtyQtl ?? 0), 0)
          );
          const adjustedOutQtyQtl = roundQtl(
            adjustedOutRows.reduce((sum, item) => sum + Number(item.adjustedQtyQtl ?? 0), 0)
          );
          const adjustedInBags = adjustedInRows.reduce((sum, item) => sum + Number(item.adjustedBags ?? 0), 0);
          const adjustedOutBags = adjustedOutRows.reduce((sum, item) => sum + Number(item.adjustedBags ?? 0), 0);
          return {
            regCode: row.regCode,
            farmerName: row.farmerName,
            village: row.village,
            district: row.district,
            originalQtyQtl: roundQtl(row.qtyQtl),
            adjustedInQtyQtl,
            adjustedOutQtyQtl,
            finalQtyQtl: roundQtl(row.qtyQtl + adjustedInQtyQtl - adjustedOutQtyQtl),
            originalBags: row.bags,
            adjustedInBags,
            adjustedOutBags,
            finalBags: row.bags + adjustedInBags - adjustedOutBags,
            changed:
              adjustedInQtyQtl > 0 ||
              adjustedOutQtyQtl > 0 ||
              adjustedInBags > 0 ||
              adjustedOutBags > 0
          };
        });

        return {
          title: "Adjusted Stack Card Preview",
          godownName: selectedDiscrepancy.godownName,
          stackNo: selectedDiscrepancy.stackNo,
          generatedAt: new Date().toISOString(),
          fileName: `adjusted-stack-card-${selectedDiscrepancy.godownName
            .replace(/[^a-z0-9]+/gi, "-")
            .replace(/^-+|-+$/g, "")
            .toLowerCase()}-${String(selectedDiscrepancy.stackNo)
            .replace(/[^a-z0-9]+/gi, "-")
            .replace(/^-+|-+$/g, "")
            .toLowerCase()}.xlsx`,
          totalAccommodatedQtyQtl: roundQtl(
            selectedStackAccommodationRows.reduce((sum, item) => sum + Number(item.adjustedQtyQtl ?? 0), 0)
          ),
          totalAccommodatedBags: selectedStackAccommodationRows.reduce(
            (sum, item) => sum + Number(item.adjustedBags ?? 0),
            0
          ),
          changedFarmerCount: adjustedRows.filter((row) => row.changed).length,
          originalRows,
          adjustedRows
        } satisfies AdjustedStackCardPreview;
      })()
    : null;
  const selectedDiscrepancyMappedQty = roundQtl(
    selectedDiscrepancyAccommodations.reduce((sum, item) => sum + Number(item.adjustedQtyQtl ?? 0), 0)
  );
  const selectedDiscrepancyMappedBags = selectedDiscrepancyAccommodations.reduce(
    (sum, item) => sum + Number(item.adjustedBags ?? 0),
    0
  );
  const selectedDiscrepancyRemainingForAccommodation = selectedDiscrepancy
    ? roundQtl(Math.max(Number(selectedDiscrepancy.excessQtyQtl ?? 0) - selectedDiscrepancyMappedQty, 0))
    : 0;
  const selectedDiscrepancyRemainingBagsForAccommodation = selectedDiscrepancy
    ? Math.max(Number(selectedDiscrepancy.estimatedExcessBags ?? 0) - selectedDiscrepancyMappedBags, 0)
    : 0;
  const pendingAccommodationRows = discrepancies
    .map((item) => {
      const mappedQtyQtl = roundQtl(
        stackAccommodations
          .filter((entry) => entry.discrepancyId === item.id)
          .reduce((sum, entry) => sum + Number(entry.adjustedQtyQtl ?? 0), 0)
      );
      const mappedBags = stackAccommodations
        .filter((entry) => entry.discrepancyId === item.id)
        .reduce((sum, entry) => sum + Number(entry.adjustedBags ?? 0), 0);
      const remainingQtyQtl = roundQtl(Math.max(Number(item.excessQtyQtl ?? 0) - mappedQtyQtl, 0));
      const remainingBags = Math.max(Number(item.estimatedExcessBags ?? 0) - mappedBags, 0);
      return {
        discrepancy: item,
        mappedQtyQtl,
        mappedBags,
        remainingQtyQtl,
        remainingBags
      };
    })
    .filter((item) => item.discrepancy.status !== "RESOLVED" && item.remainingQtyQtl > 0)
    .sort((left, right) => {
      if (right.remainingQtyQtl !== left.remainingQtyQtl) {
        return right.remainingQtyQtl - left.remainingQtyQtl;
      }
      return left.discrepancy.receiptDate.localeCompare(right.discrepancy.receiptDate);
    });
  const eligibleAccommodationTargets = selectedDiscrepancy
    ? Array.from(
        receipts.reduce(
          (map, receipt) => {
            if (receipt.cropRegistrationId === selectedDiscrepancy.cropRegistrationId) {
              return map;
            }
            const sameStackLines = receipt.lines.filter(
              (line) =>
                line.godownId === selectedDiscrepancy.godownId &&
                String(line.stackNo ?? "").trim() === String(selectedDiscrepancy.stackNo ?? "").trim()
            );
            if (!sameStackLines.length) {
              return map;
            }
            const registration = registrations.find((item) => item.id === receipt.cropRegistrationId);
            if (!registration) {
              return map;
            }
            const current = map.get(registration.id) ?? {
              registration,
              stackQtyQtl: 0,
              stackBags: 0
            };
            current.stackQtyQtl += sameStackLines.reduce((sum, line) => sum + Number(line.netWeightQtl ?? 0), 0);
            current.stackBags += sameStackLines.reduce((sum, line) => sum + Number(line.noOfBags ?? 0), 0);
            map.set(registration.id, current);
            return map;
          },
          new Map<
            string,
            {
              registration: RegistrationRecord;
              stackQtyQtl: number;
              stackBags: number;
            }
          >()
        ).values()
      ).sort((left, right) => right.registration.balanceQtl - left.registration.balanceQtl)
    : [];
  useEffect(() => {
    if (!selectedDiscrepancy) {
      setAccommodationTargetRegistrationId("");
      setAccommodationQtyQtl("");
      setAccommodationBags("");
      setAccommodationRemarks("");
      setEditingAccommodationId("");
      setDiscrepancyWorkflowMode("accommodation");
      return;
    }
    setAccommodationTargetRegistrationId((current) =>
      current && eligibleAccommodationTargets.some((item) => item.registration.id === current)
        ? current
        : eligibleAccommodationTargets[0]?.registration.id ?? ""
    );
    setAccommodationQtyQtl((current) =>
      current && Number(current) > 0 ? current : String(selectedDiscrepancyRemainingForAccommodation || "")
    );
    setAccommodationBags((current) =>
      current && Number(current) >= 0 ? current : String(selectedDiscrepancyRemainingBagsForAccommodation || 0)
    );
    setAccommodationDate(new Date().toISOString().slice(0, 10));
    setAccommodationRemarks("");
    setEditingAccommodationId("");
  }, [
    selectedDiscrepancy?.id,
    eligibleAccommodationTargets,
    selectedDiscrepancyRemainingForAccommodation,
    selectedDiscrepancyRemainingBagsForAccommodation
  ]);
  const depositViewRegistration =
    registrations.find((item) => item.id === depositViewRegistrationId) ?? null;
  const registrationReceipts = depositViewRegistration
    ? receipts
        .filter((item) => item.cropRegistrationId === depositViewRegistration.id)
        .slice()
        .sort((left, right) => left.receiptDate.localeCompare(right.receiptDate))
    : [];
  const registrationLots = depositViewRegistration
    ? lotLedgerRows.filter((item) => item.cropRegistrationId === depositViewRegistration.id)
    : [];
  const registrationDiscrepancies = depositViewRegistration
    ? discrepancies.filter((item) => item.cropRegistrationId === depositViewRegistration.id)
    : [];
  const registrationShifts = depositViewRegistration
    ? discrepancyShifts.filter(
        (item) => item.cropRegistrationCode === depositViewRegistration.cropRegistrationCode
      )
    : [];
  const slipRegistration = registrations.find((item) => item.id === slipRegistrationId) ?? null;
  const slipReceipts = slipRegistration
    ? receipts
        .filter((item) => item.cropRegistrationId === slipRegistration.id)
        .slice()
        .sort((left, right) =>
          `${left.receiptDate}-${left.receiptNo}`.localeCompare(`${right.receiptDate}-${right.receiptNo}`)
        )
    : [];
  const slipReceipt = slipReceipts.find((item) => item.receiptNo === slipReceiptNo) ?? null;
  const slipDailyReceipts =
    slipRegistration && slipDate
      ? slipReceipts.filter((item) => item.receiptDate === slipDate)
      : [];
  const slipDiscrepancies = slipRegistration
    ? discrepancies.filter((item) => item.cropRegistrationId === slipRegistration.id)
    : [];
  const slipShifts = slipRegistration
    ? discrepancyShifts.filter((item) => item.cropRegistrationCode === slipRegistration.cropRegistrationCode)
    : [];
  const slipLots = slipRegistration
    ? lotLedgerRows.filter((item) => item.cropRegistrationId === slipRegistration.id)
    : [];
  const slipDistrictOptions = Array.from(
    new Set(registrations.map((item) => item.district).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, "en", { sensitivity: "base" }));
  const slipVillageOptions = Array.from(
    new Set(registrations.map((item) => item.village).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, "en", { sensitivity: "base" }));
  const slipClassOptions = Array.from(
    new Set(registrations.map((item) => item.classStage).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, "en", { sensitivity: "base" }));
  const slipCropOptions = Array.from(
    new Set(registrations.map((item) => item.crop).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, "en", { sensitivity: "base" }));
  const slipRegistrationRows = registrations
    .filter((item) => item.status !== "BLOCKED")
    .filter((item) => {
      const query = slipSearch.trim().toLowerCase();
      if (!query) {
        return true;
      }

      return (
        item.cropRegistrationCode.toLowerCase().includes(query) ||
        item.farmerName.toLowerCase().includes(query) ||
        item.village.toLowerCase().includes(query) ||
        item.district.toLowerCase().includes(query) ||
        item.classStage.toLowerCase().includes(query)
      );
    })
    .filter((item) => !slipDistrictFilter || item.district === slipDistrictFilter)
    .filter((item) => !slipVillageFilter || item.village === slipVillageFilter)
    .filter((item) => !slipClassFilter || item.classStage === slipClassFilter)
    .filter((item) => !slipCropFilter || item.crop === slipCropFilter)
    .filter((item) => !slipOnlyWithIntake || item.totalReceivedQtl > 0)
    .slice()
    .sort((left, right) =>
      left.farmerName.localeCompare(right.farmerName, "en", { sensitivity: "base" })
    );
  const voucherRegistration = registrations.find((item) => item.id === voucherRegistrationId) ?? null;
  const isAdminUser = currentUser?.role === "ADMIN";
  const effectivePermissions =
    currentUser
      ? {
          ...rolePermissions[currentUser.role],
          ...(currentPermissions ?? {})
        }
      : null;
  const voucherByRegistrationId = new Map(
    financialVouchers.map((item) => [item.cropRegistrationId, item] as const)
  );
  const paymentLedgerVoucher =
    financialVouchers.find((item) => item.id === paymentLedgerVoucherId) ?? null;
  const activeVoucher =
    (voucherRegistrationId ? voucherByRegistrationId.get(voucherRegistrationId) : null) ??
    voucherPreview?.voucher ??
    null;
  const voucherRows = registrations
    .filter((item) => item.status !== "BLOCKED")
    .filter((item) => {
      const query = voucherGenerationSearch.trim().toLowerCase();
      const matchesSearch =
        !query ||
        item.cropRegistrationCode.toLowerCase().includes(query) ||
        item.farmerName.toLowerCase().includes(query) ||
        item.village.toLowerCase().includes(query) ||
        item.district.toLowerCase().includes(query);
      const matchesDistrict =
        !voucherGenerationDistrictFilter || item.district === voucherGenerationDistrictFilter;
      return matchesSearch && matchesDistrict;
    })
    .slice()
    .sort((left, right) =>
      left.farmerName.localeCompare(right.farmerName, "en", { sensitivity: "base" })
    );
  const voucherDistrictOptions = Array.from(
    new Set(registrations.map((item) => item.district).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, "en", { sensitivity: "base" }));
  const voucherRegisterStatusOptions = Array.from(
    new Set(financialVouchers.map((item) => String(item.status ?? "").trim()).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, "en", { sensitivity: "base" }));
  const getVoucherOrganizerName = (voucher: FinancialVoucher) => {
    const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
    return String(registration?.organizerName ?? "").trim() || "Direct Farmer";
  };
  const voucherRegisterOrganizerOptions = Array.from(
    new Set(financialVouchers.map((item) => getVoucherOrganizerName(item)).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, "en", { sensitivity: "base" }));
  const filteredVoucherRegisterRows = financialVouchers
    .filter((voucher) => {
      const query = voucherRegisterSearch.trim().toLowerCase();
      const matchesSearch =
        !query ||
        voucher.voucherNo.toLowerCase().includes(query) ||
        voucher.cropRegistrationCode.toLowerCase().includes(query) ||
        voucher.farmerName.toLowerCase().includes(query) ||
        voucher.village.toLowerCase().includes(query) ||
        voucher.district.toLowerCase().includes(query);
      const matchesDistrict =
        !voucherRegisterDistrictFilter || voucher.district === voucherRegisterDistrictFilter;
      const matchesStatus = !voucherRegisterStatusFilter || voucher.status === voucherRegisterStatusFilter;
      const matchesOrganizer =
        !voucherRegisterOrganizerFilter || getVoucherOrganizerName(voucher) === voucherRegisterOrganizerFilter;
      return matchesSearch && matchesDistrict && matchesStatus && matchesOrganizer;
    })
    .slice()
    .sort((left, right) =>
      left.farmerName.localeCompare(right.farmerName, "en", { sensitivity: "base" })
    );
  const registrationDiscrepancyMap = new Map<string, { qtyQtl: number; bags: number }>();
  discrepancies.forEach((item) => {
    const current = registrationDiscrepancyMap.get(item.cropRegistrationId) ?? { qtyQtl: 0, bags: 0 };
    registrationDiscrepancyMap.set(item.cropRegistrationId, {
      qtyQtl: roundQtl(current.qtyQtl + Number(item.excessQtyQtl ?? 0)),
      bags: current.bags + Number(item.estimatedExcessBags ?? 0)
    });
  });
  const registrationBagMap = new Map<string, number>();
  receipts.forEach((receipt) => {
    const totalBags = receipt.lines.reduce(
      (sum, line) => sum + Number(line.noOfBags ?? 0),
      0
    );
    registrationBagMap.set(
      receipt.cropRegistrationId,
      (registrationBagMap.get(receipt.cropRegistrationId) ?? 0) + totalBags
    );
  });
  const districtDashboardRows = Array.from(
    registrations.reduce(
      (map, registration) => {
        const key = registration.district || "UNSPECIFIED";
        const current = map.get(key) ?? {
          district: key,
          registrations: 0,
          expectedYield: 0,
          receivedNet: 0,
          pending: 0,
          bags: 0,
          discrepancyCases: 0
        };
        current.registrations += 1;
        current.expectedYield += Number(registration.expectedYieldQtl ?? 0);
        current.receivedNet += Number(registration.totalReceivedQtl ?? 0);
        current.pending += Number(registration.balanceQtl ?? 0);
        current.bags += registrationBagMap.get(registration.id) ?? 0;
        current.discrepancyCases += openDiscrepancies.filter(
          (item) => item.cropRegistrationId === registration.id
        ).length;
        map.set(key, current);
        return map;
      },
      new Map<
        string,
        {
          district: string;
          registrations: number;
          expectedYield: number;
          receivedNet: number;
          pending: number;
          bags: number;
          discrepancyCases: number;
        }
      >()
    ).values()
  )
    .map((item) => ({
      ...item,
      coveragePct: item.expectedYield > 0 ? (item.receivedNet / item.expectedYield) * 100 : 0
    }))
    .sort((left, right) => right.receivedNet - left.receivedNet);
  const pendingRegistrationRows = registrations
    .filter((item) => item.balanceQtl > 0)
    .slice()
    .sort((left, right) => right.balanceQtl - left.balanceQtl);
  const stackHotspotRows = Array.from(
    openDiscrepancies.reduce(
      (map, item) => {
        const key = `${item.godownName}::${item.stackNo}`;
        const current = map.get(key) ?? {
          key,
          godownName: item.godownName,
          stackNo: item.stackNo,
          cases: 0,
          bags: 0,
          qtyQtl: 0
        };
        current.cases += 1;
        current.bags += Number(item.estimatedExcessBags ?? 0);
        current.qtyQtl += Number(item.excessQtyQtl ?? 0);
        map.set(key, current);
        return map;
      },
      new Map<
        string,
        {
          key: string;
          godownName: string;
          stackNo: string;
          cases: number;
          bags: number;
          qtyQtl: number;
        }
      >()
    ).values()
  )
    .sort((left, right) => right.qtyQtl - left.qtyQtl);
  const registrationReceiptLines = registrationReceipts.flatMap((receipt) =>
    receipt.lines.map((line) => ({
      receiptNo: receipt.receiptNo,
      receiptDate: receipt.receiptDate,
      vehicleNo: line.vehicleNo,
      stackNo: line.stackNo,
      godownName: godowns.find((item) => item.id === line.godownId)?.name ?? "-",
      bags: Number(line.noOfBags ?? 0),
      grossWeightQtl: Number(line.grossWeightQtl ?? 0),
      netWeightQtl: Number(line.netWeightQtl ?? 0),
      moisturePercent: Number(line.moisturePercent ?? 0),
      remarks: line.remarks
    }))
  );
  const depositSummary = depositViewRegistration
    ? {
        totalGrossQtl: registrationReceiptLines.reduce((sum, item) => sum + item.grossWeightQtl, 0),
        totalNetQtl: registrationReceiptLines.reduce((sum, item) => sum + item.netWeightQtl, 0),
        totalBags: registrationReceiptLines.reduce((sum, item) => sum + item.bags, 0),
        discrepancyQtyQtl: registrationDiscrepancies.reduce(
          (sum, item) => sum + item.excessQtyQtl,
          0
        ),
        shiftedQtyQtl: registrationShifts.reduce((sum, item) => sum + item.shiftedQtyQtl, 0),
        shiftedBags: registrationShifts.reduce((sum, item) => sum + item.shiftedBags, 0)
      }
    : null;
  const stackWiseDepositRows = depositViewRegistration
    ? Object.values(
        registrationReceiptLines.reduce<
          Record<
            string,
            {
              godownName: string;
              stackNo: string;
              totalGrossQtl: number;
              totalNetQtl: number;
              totalBags: number;
            }
          >
        >((accumulator, item) => {
          const key = `${item.godownName}::${item.stackNo}`;
          if (!accumulator[key]) {
            accumulator[key] = {
              godownName: item.godownName,
              stackNo: item.stackNo,
              totalGrossQtl: 0,
              totalNetQtl: 0,
              totalBags: 0
            };
          }
          accumulator[key].totalGrossQtl += item.grossWeightQtl;
          accumulator[key].totalNetQtl += item.netWeightQtl;
          accumulator[key].totalBags += item.bags;
          return accumulator;
        }, {})
      )
    : [];
  const recentReceipts = receipts
    .slice()
    .sort((left, right) =>
      `${right.receiptDate}-${right.receiptNo}`.localeCompare(`${left.receiptDate}-${left.receiptNo}`)
    );
  const recentVoucherRows = financialVouchers
    .slice()
    .sort((left, right) =>
      `${right.voucherDate}-${right.voucherNo}`.localeCompare(`${left.voucherDate}-${left.voucherNo}`)
    );
  const visibleRecentReceiptRows = dashboardExpandedSections.recentReceipts
    ? recentReceipts
    : recentReceipts.slice(0, 5);
  const visibleRecentVoucherRows = dashboardExpandedSections.recentVouchers
    ? recentVoucherRows
    : recentVoucherRows.slice(0, 5);
  const topGodownStock = godowns
    .map((godown) => {
      const godownLots = lots.filter((lot) => lot.godownId === godown.id);
      const qtyQtl = godownLots.reduce((sum, lot) => sum + Number(lot.currentQtyQtl ?? 0), 0);
      const capacityQtl = godownLots.length * 200;
      return {
        name: godown.name,
        qtyQtl,
        lots: godownLots.length,
        fullLots: godownLots.filter((lot) => lot.status === "FULL").length,
        utilizationPct: capacityQtl > 0 ? (qtyQtl / capacityQtl) * 100 : 0
      };
    })
    .sort((left, right) => right.qtyQtl - left.qtyQtl);
  const topDiscrepancies = openDiscrepancies
    .slice()
    .sort((left, right) => right.excessQtyQtl - left.excessQtyQtl)
    .slice(0, 5);
  const visibleDistrictDashboardRows = dashboardExpandedSections.districts
    ? districtDashboardRows
    : districtDashboardRows.slice(0, 5);
  const visibleGodownStockRows = dashboardExpandedSections.godowns
    ? topGodownStock
    : topGodownStock.slice(0, 5);
  const visiblePendingRegistrationRows = dashboardExpandedSections.pendingRegistrations
    ? pendingRegistrationRows
    : pendingRegistrationRows.slice(0, 5);
  const visibleStackHotspotRows = dashboardExpandedSections.stackHotspots
    ? stackHotspotRows
    : stackHotspotRows.slice(0, 5);
  const shortSnapshot = [
    `${receipts.length} receipts`,
    `${dashboardMetrics.intakeBags} bags`,
    `${formatNumber(dashboardMetrics.grossReceived)} gross QTL`,
    `${formatNumber(dashboardMetrics.netReceived)} net QTL`,
    `${dashboardMetrics.totalLots} lots`,
    `${dashboardMetrics.discrepancyCount} open discrepancies`
  ].join(" | ");

  const dashboardAssistantSuggestions = [
    "Which organizer farmers have not started intake?",
    "Show organizer-wise pending commission balance",
    "Show pending registrations for intake",
    "Which vouchers are still draft or unpaid?",
    "Show open discrepancy stacks",
    "Show district-wise intake performance"
  ];

  function normalizeDashboardQuestion(value: string) {
    return value.trim().toLowerCase();
  }

  function findQuestionMatch(question: string, values: string[]) {
    const normalizedValues = values
      .map((value) => value?.trim())
      .filter(Boolean)
      .sort((left, right) => String(right).length - String(left).length);
    return normalizedValues.find((value) => question.includes(String(value).toLowerCase())) ?? "";
  }

  function runDashboardAssistant(questionOverride?: string) {
    const rawQuestion = questionOverride ?? dashboardQuestion;
    const normalizedQuestion = normalizeDashboardQuestion(rawQuestion);

    if (!normalizedQuestion) {
      setDashboardAssistantResult({
        title: "Ask Your System",
        summary:
          "Type a question about organizers, pending intake, vouchers, discrepancies, or districts to get a live dashboard answer.",
        columns: ["Suggested Questions"],
        rows: dashboardAssistantSuggestions.map((item) => ({ "Suggested Questions": item }))
      });
      return;
    }

    const asksOrganizer = normalizedQuestion.includes("organizer") || normalizedQuestion.includes("organiser");
    const asksNoIntake =
      normalizedQuestion.includes("no intake") ||
      normalizedQuestion.includes("not yet intake") ||
      normalizedQuestion.includes("zero intake") ||
      normalizedQuestion.includes("not started intake");
    const asksPending =
      normalizedQuestion.includes("pending") ||
      normalizedQuestion.includes("balance") ||
      normalizedQuestion.includes("outstanding") ||
      normalizedQuestion.includes("unpaid");
    const asksVoucher =
      normalizedQuestion.includes("voucher") ||
      normalizedQuestion.includes("payment") ||
      normalizedQuestion.includes("paid") ||
      normalizedQuestion.includes("draft");
    const asksDiscrepancy =
      normalizedQuestion.includes("discrepancy") ||
      normalizedQuestion.includes("stack") ||
      normalizedQuestion.includes("hotspot");
    const asksDistrict =
      normalizedQuestion.includes("district") ||
      normalizedQuestion.includes("performance") ||
      normalizedQuestion.includes("coverage");
    const asksRegistration =
      normalizedQuestion.includes("registration") ||
      normalizedQuestion.includes("farmer") ||
      normalizedQuestion.includes("intake");
    const asksDirectFarmer =
      normalizedQuestion.includes("direct farmer") || normalizedQuestion.includes("direct farmers");
    const asksTop =
      normalizedQuestion.includes("top") ||
      normalizedQuestion.includes("highest") ||
      normalizedQuestion.includes("maximum") ||
      normalizedQuestion.includes("most");
    const asksSummary =
      normalizedQuestion.includes("summary") ||
      normalizedQuestion.includes("total") ||
      normalizedQuestion.includes("how much");
    const matchedDistrict = findQuestionMatch(
      normalizedQuestion,
      districtDashboardRows.map((item) => item.district)
    );
    const matchedOrganizer = findQuestionMatch(
      normalizedQuestion,
      organizers.map((item) => item.name).concat(["direct farmer"])
    );
    const matchedVillage = findQuestionMatch(
      normalizedQuestion,
      registrations.map((item) => item.village).filter(Boolean)
    );
    const matchedFarmer = findQuestionMatch(
      normalizedQuestion,
      registrations.map((item) => item.farmerName)
    );
    const matchedStatus = ["draft", "paid", "overpaid", "unpaid", "pending"].find((item) =>
      normalizedQuestion.includes(item)
    );
    const asksSeedDeductionNotMade =
      normalizedQuestion.includes("seed deduction is not made") ||
      normalizedQuestion.includes("seed deduction not made") ||
      normalizedQuestion.includes("deduction is not made") ||
      normalizedQuestion.includes("deduction not made") ||
      normalizedQuestion.includes("deduction is not done") ||
      normalizedQuestion.includes("deduction not done") ||
      normalizedQuestion.includes("without deduction") ||
      normalizedQuestion.includes("no deduction");
    const asksUniqueFarmerList =
      normalizedQuestion.includes("unique farmer") ||
      normalizedQuestion.includes("unique farmer name") ||
      normalizedQuestion.includes("unique farmers");
    const asksSortByOrganizer =
      normalizedQuestion.includes("sort by organiser") ||
      normalizedQuestion.includes("sort by organizer") ||
      normalizedQuestion.includes("organizer wise") ||
      normalizedQuestion.includes("organiser wise");
    const questionFilters: string[] = [];
    if (matchedDistrict) {
      questionFilters.push(`District: ${matchedDistrict}`);
    }
    if (matchedOrganizer) {
      questionFilters.push(`Organizer: ${matchedOrganizer}`);
    }
    if (matchedVillage) {
      questionFilters.push(`Village: ${matchedVillage}`);
    }
    if (matchedFarmer) {
      questionFilters.push(`Farmer: ${matchedFarmer}`);
    }
    if (matchedStatus) {
      questionFilters.push(`Status: ${matchedStatus.toUpperCase()}`);
    }
    const filterSuffix = questionFilters.length ? ` Filtered by ${questionFilters.join(" | ")}.` : "";

    const matchesOrganizerFilter = (organizerName: string) =>
      !matchedOrganizer ||
      organizerName.toLowerCase() === matchedOrganizer ||
      (matchedOrganizer === "direct farmer" && organizerName.toLowerCase() === "direct farmer");
    const matchesDistrictFilter = (district: string) =>
      !matchedDistrict || String(district || "").toLowerCase() === matchedDistrict;
    const matchesVillageFilter = (village: string) =>
      !matchedVillage || String(village || "").toLowerCase() === matchedVillage;
    const matchesFarmerFilter = (farmerName: string) =>
      !matchedFarmer || String(farmerName || "").toLowerCase() === matchedFarmer;

    if (asksOrganizer && asksNoIntake) {
      const filteredRows = organizerNoIntakeRows.filter(
        (row) =>
          matchesOrganizerFilter(row.organizer.name) &&
          matchesDistrictFilter(row.organizer.district) &&
          (!matchedFarmer ||
            row.zeroIntakeRegistrations.some((registration) => matchesFarmerFilter(registration.farmerName))) &&
          (!matchedVillage ||
            row.zeroIntakeRegistrations.some((registration) => matchesVillageFilter(registration.village)))
      );
      const resultRows = asksTop ? filteredRows.slice(0, 1) : filteredRows;
      setDashboardAssistantResult({
        title: "Organizer Farmers Not Yet Intake",
        summary: `${resultRows.reduce(
          (sum, row) => sum + row.zeroIntakeFarmerCount,
          0
        )} organizer-linked farmer(s) still have zero intake against ${formatNumber(
          resultRows.reduce((sum, row) => sum + row.zeroIntakePendingQty, 0)
        )} QTL pending quantity.${filterSuffix}`,
        columns: ["Organizer", "District", "No-Intake Farmers", "Pending Qty", "Farmer Preview"],
        rows: resultRows.map((row) => ({
          Organizer: row.organizer.name,
          District: row.organizer.district || "-",
          "No-Intake Farmers": row.zeroIntakeFarmerCount,
          "Pending Qty": `${formatNumber(row.zeroIntakePendingQty)} QTL`,
          "Farmer Preview": row.farmerPreview || "-"
        }))
      });
      return;
    }

    if (asksOrganizer && (asksPending || normalizedQuestion.includes("commission"))) {
      const filteredRows = organizerDashboardRows.filter(
        (row) =>
          matchesOrganizerFilter(row.organizer.name) &&
          matchesDistrictFilter(row.organizer.district) &&
          (!matchedFarmer ||
            row.linkedRegistrations.some((registration) => matchesFarmerFilter(registration.farmerName))) &&
          (!matchedVillage ||
            row.linkedRegistrations.some((registration) => matchesVillageFilter(registration.village)))
      );
      const resultRows = asksTop ? filteredRows.slice(0, 1) : filteredRows;
      setDashboardAssistantResult({
        title: "Organizer Commission Pending Balance",
        summary: `${resultRows.filter((row) => row.balanceAmount > 0).length} organizer(s) still have ${formatNumber(
          resultRows.reduce((sum, row) => sum + row.balanceAmount, 0)
        )} INR pending after ${formatNumber(
          resultRows.reduce((sum, row) => sum + row.deductionAmount, 0)
        )} INR deduction.${filterSuffix}`,
        columns: [
          "Organizer",
          "District",
          "Farmers",
          "Gross Commission",
          "Deduction",
          "Net Payable",
          "Paid",
          "Balance"
        ],
        rows: resultRows.map((row) => ({
          Organizer: row.organizer.name,
          District: row.organizer.district || "-",
          Farmers: row.farmerCount,
          "Gross Commission": formatNumber(row.grossCommissionAmount),
          Deduction: formatNumber(row.deductionAmount),
          "Net Payable": formatNumber(row.netPayableAmount),
          Paid: formatNumber(row.paidAmount),
          Balance: formatNumber(row.balanceAmount)
        }))
      });
      return;
    }

    if (asksDirectFarmer && asksVoucher) {
      const directFarmerRows = financialVouchers
        .filter((voucher) => {
          const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
          return (
            !registration?.organizerId &&
            matchesDistrictFilter(voucher.district) &&
            matchesVillageFilter(voucher.village) &&
            matchesFarmerFilter(voucher.farmerName)
          );
        })
        .sort((left, right) =>
          `${left.voucherDate}-${left.voucherNo}`.localeCompare(`${right.voucherDate}-${right.voucherNo}`)
        );
      const resultRows = asksTop
        ? directFarmerRows
            .slice()
            .sort((left, right) => Number(right.balanceAmount ?? 0) - Number(left.balanceAmount ?? 0))
            .slice(0, 1)
        : directFarmerRows;
      setDashboardAssistantResult({
        title: "Direct Farmer Payment Position",
        summary: `${resultRows.length} voucher(s) belong to farmers without organizer linkage.${filterSuffix}`,
        columns: ["Date", "Voucher No.", "Farmer", "District", "Final Payable", "Paid", "Balance", "Status"],
        rows: resultRows.map((row) => ({
          Date: row.voucherDate,
          "Voucher No.": row.voucherNo,
          Farmer: row.farmerName,
          District: row.district || "-",
          "Final Payable": formatNumber(row.finalPayableAmount),
          Paid: formatNumber(row.totalPaidAmount),
          Balance: formatNumber(row.balanceAmount),
          Status: row.status
        }))
      });
      return;
    }

    if (asksSeedDeductionNotMade) {
      const deductionRows = financialVouchers
        .filter((voucher) => {
          const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
          const organizerName = registration?.organizerId
            ? organizers.find((item) => item.id === registration.organizerId)?.name || "Direct Farmer"
            : "Direct Farmer";
          return (
            Number(voucher.deductionAmount ?? 0) <= 0 &&
            matchesDistrictFilter(voucher.district) &&
            matchesVillageFilter(voucher.village) &&
            matchesFarmerFilter(voucher.farmerName) &&
            matchesOrganizerFilter(organizerName)
          );
        })
        .map((voucher) => {
          const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
          const organizerName = registration?.organizerId
            ? organizers.find((item) => item.id === registration.organizerId)?.name || "Direct Farmer"
            : "Direct Farmer";
          return {
            organizerName,
            voucher
          };
        });

      const uniqueRows = asksUniqueFarmerList
        ? Array.from(
            deductionRows.reduce(
              (map, row) => {
                const key = row.voucher.cropRegistrationId || row.voucher.cropRegistrationCode;
                if (!map.has(key)) {
                  map.set(key, row);
                }
                return map;
              },
              new Map<
                string,
                {
                  organizerName: string;
                  voucher: FinancialVoucher;
                }
              >()
            ).values()
          )
        : deductionRows;

      const sortedRows = uniqueRows.slice().sort((left, right) => {
        if (asksSortByOrganizer || asksOrganizer) {
          const organizerCompare = left.organizerName.localeCompare(right.organizerName, "en", {
            sensitivity: "base"
          });
          if (organizerCompare !== 0) {
            return organizerCompare;
          }
        }
        return left.voucher.farmerName.localeCompare(right.voucher.farmerName, "en", {
          sensitivity: "base"
        });
      });

      setDashboardAssistantResult({
        title: "Farmers With Seed Deduction Not Made",
        summary: `${sortedRows.length} farmer voucher row(s) match the condition that seed deduction amount is zero.${filterSuffix}`,
        columns: [
          "Organizer",
          "Farmer",
          "Reg. Code",
          "Village",
          "District",
          "Voucher No.",
          "Final Payable",
          "Deduction Amount",
          "Status"
        ],
        rows: sortedRows.map((row) => ({
          Organizer: row.organizerName,
          Farmer: row.voucher.farmerName,
          "Reg. Code": row.voucher.cropRegistrationCode,
          Village: row.voucher.village || "-",
          District: row.voucher.district || "-",
          "Voucher No.": row.voucher.voucherNo,
          "Final Payable": formatNumber(row.voucher.finalPayableAmount),
          "Deduction Amount": formatNumber(row.voucher.deductionAmount),
          Status: row.voucher.status
        }))
      });
      return;
    }

    if (asksVoucher && normalizedQuestion.includes("overpaid")) {
      const overpaidRows = financialVouchers
        .filter(
          (row) =>
            Number(row.balanceAmount ?? 0) < 0 &&
            matchesDistrictFilter(row.district) &&
            matchesVillageFilter(row.village) &&
            matchesFarmerFilter(row.farmerName) &&
            (!matchedOrganizer ||
              (registrations.find((item) => item.id === row.cropRegistrationId)?.organizerId
                ? organizers
                    .find(
                      (organizer) =>
                        organizer.id ===
                        registrations.find((item) => item.id === row.cropRegistrationId)?.organizerId
                    )
                    ?.name.toLowerCase() === matchedOrganizer
                : matchedOrganizer === "direct farmer"))
        )
        .sort((left, right) => Number(left.balanceAmount ?? 0) - Number(right.balanceAmount ?? 0));
      const resultRows = asksTop ? overpaidRows.slice(0, 1) : overpaidRows;
      setDashboardAssistantResult({
        title: "Overpaid Farmer Vouchers",
        summary: `${resultRows.length} voucher(s) currently show negative balance, meaning overpayment has been recorded.${filterSuffix}`,
        columns: ["Date", "Voucher No.", "Farmer", "District", "Final Payable", "Paid", "Balance", "Status"],
        rows: resultRows.map((row) => ({
          Date: row.voucherDate,
          "Voucher No.": row.voucherNo,
          Farmer: row.farmerName,
          District: row.district || "-",
          "Final Payable": formatNumber(row.finalPayableAmount),
          Paid: formatNumber(row.totalPaidAmount),
          Balance: formatNumber(row.balanceAmount),
          Status: row.status
        }))
      });
      return;
    }

    if (asksVoucher && (normalizedQuestion.includes("draft") || asksPending)) {
      const voucherRows = financialVouchers
        .filter((row) => {
          const registration = registrations.find((item) => item.id === row.cropRegistrationId);
          const organizerName = registration?.organizerId
            ? organizers.find((item) => item.id === registration.organizerId)?.name || ""
            : "Direct Farmer";
          const statusMatches =
            matchedStatus === "draft"
              ? row.status === "DRAFT"
              : matchedStatus === "paid"
                ? row.status === "PAID"
                : matchedStatus === "overpaid"
                  ? Number(row.balanceAmount ?? 0) < 0
                  : matchedStatus === "unpaid" || matchedStatus === "pending"
                    ? Number(row.balanceAmount ?? 0) > 0
                    : normalizedQuestion.includes("draft")
                      ? row.status === "DRAFT"
                      : Number(row.balanceAmount ?? 0) > 0;
          return (
            statusMatches &&
            matchesDistrictFilter(row.district) &&
            matchesVillageFilter(row.village) &&
            matchesFarmerFilter(row.farmerName) &&
            matchesOrganizerFilter(organizerName)
          );
        })
        .slice()
        .sort((left, right) => {
          if (asksTop) {
            return Number(right.balanceAmount ?? 0) - Number(left.balanceAmount ?? 0);
          }
          return `${left.voucherDate}-${left.voucherNo}`.localeCompare(`${right.voucherDate}-${right.voucherNo}`);
        });
      const resultRows = asksTop ? voucherRows.slice(0, 1) : voucherRows;
      setDashboardAssistantResult({
        title: normalizedQuestion.includes("draft") ? "Draft Vouchers" : "Pending Farmer Payment Vouchers",
        summary: normalizedQuestion.includes("draft")
          ? `${resultRows.length} voucher(s) are still in draft mode.${filterSuffix}`
          : `${resultRows.length} voucher(s) are not fully settled and show pending farmer payment balance.${filterSuffix}`,
        columns: ["Date", "Voucher No.", "Farmer", "District", "Final Payable", "Paid", "Balance", "Status"],
        rows: resultRows.map((row) => ({
          Date: row.voucherDate,
          "Voucher No.": row.voucherNo,
          Farmer: row.farmerName,
          District: row.district || "-",
          "Final Payable": formatNumber(row.finalPayableAmount),
          Paid: formatNumber(row.totalPaidAmount),
          Balance: formatNumber(row.balanceAmount),
          Status: row.status
        }))
      });
      return;
    }

    if (asksDiscrepancy) {
      const discrepancyRows = openDiscrepancies
        .filter((row) => {
          const registration = registrations.find((item) => item.id === row.cropRegistrationId);
          return (
            matchesDistrictFilter(registration?.district || "") &&
            matchesVillageFilter(registration?.village || "") &&
            matchesFarmerFilter(row.farmerName)
          );
        })
        .slice()
        .sort((left, right) => Number(right.excessQtyQtl ?? 0) - Number(left.excessQtyQtl ?? 0));
      const resultRows = asksTop ? discrepancyRows.slice(0, 1) : discrepancyRows;
      setDashboardAssistantResult({
        title: "Open Discrepancy Stack Position",
        summary: `${resultRows.length} open discrepancy case(s) match this question across ${stackHotspotRows.length} hotspot group(s).${filterSuffix}`,
        columns: ["Reg. Code", "Farmer", "Godown", "Stack", "Excess Bags", "Excess Qty", "Status"],
        rows: resultRows.map((row) => ({
            "Reg. Code": row.cropRegistrationCode,
            Farmer: row.farmerName,
            Godown: row.godownName,
            Stack: row.stackNo,
            "Excess Bags": row.estimatedExcessBags,
            "Excess Qty": `${formatNumber(row.excessQtyQtl)} QTL`,
            Status: row.status
          }))
      });
      return;
    }

    if (asksDistrict) {
      const districtRows = districtDashboardRows
        .filter((row) => matchesDistrictFilter(row.district))
        .slice()
        .sort((left, right) => {
          if (asksTop) {
            return Number(right.receivedNet ?? 0) - Number(left.receivedNet ?? 0);
          }
          return Number(right.receivedNet ?? 0) - Number(left.receivedNet ?? 0);
        });
      const resultRows = asksTop ? districtRows.slice(0, 1) : districtRows;
      setDashboardAssistantResult({
        title: "District Intake Performance",
        summary: `${resultRows.length} district row(s) match this question, sorted by received quantity.${filterSuffix}`,
        columns: ["District", "Registrations", "Received", "Coverage", "Pending", "Discrepancies"],
        rows: resultRows.map((row) => ({
          District: row.district,
          Registrations: row.registrations,
          Received: `${formatNumber(row.receivedNet)} QTL`,
          Coverage: `${formatNumber(row.coveragePct)}%`,
          Pending: `${formatNumber(row.pending)} QTL`,
          Discrepancies: row.discrepancyCases
        }))
      });
      return;
    }

    if (asksRegistration || asksNoIntake) {
      const registrationRows = pendingRegistrationRows
        .filter((row) => {
          const organizerName =
            organizers.find((item) => item.id === row.organizerId)?.name || "Direct Farmer";
          return (
            matchesDistrictFilter(row.district) &&
            matchesVillageFilter(row.village) &&
            matchesFarmerFilter(row.farmerName) &&
            matchesOrganizerFilter(organizerName)
          );
        })
        .slice()
        .sort((left, right) => Number(right.balanceQtl ?? 0) - Number(left.balanceQtl ?? 0));
      const resultRows = asksTop ? registrationRows.slice(0, 1) : registrationRows;
      setDashboardAssistantResult({
        title: "Pending Farmer Intake",
        summary: `${resultRows.length} registration(s) still have pending intake balance.${filterSuffix}`,
        columns: ["Reg. Code", "Farmer", "District", "Received", "Pending", "Organizer"],
        rows: resultRows.map((row) => ({
          "Reg. Code": row.cropRegistrationCode,
          Farmer: row.farmerName,
          District: row.district || "-",
          Received: `${formatNumber(row.totalReceivedQtl)} QTL`,
          Pending: `${formatNumber(row.balanceQtl)} QTL`,
          Organizer: organizers.find((item) => item.id === row.organizerId)?.name || "Direct Farmer"
        }))
      });
      return;
    }

    setDashboardAssistantResult({
      title: "Dashboard Command Answer",
      summary:
        "I could not match that question well enough yet. Try asking with names like district, organizer, village, farmer, or status such as draft, unpaid, paid, or overpaid.",
      columns: ["Suggested Questions"],
      rows: dashboardAssistantSuggestions.map((item) => ({ "Suggested Questions": item }))
    });
  }

  async function readWorkbook(file: File) {
    const data = await file.arrayBuffer();
    return XLSX.read(data, { type: "array" });
  }

  async function handleRegistrationImport(file: File) {
    if (!requirePermission("canImport", "Only Admin can import farmer master data.")) {
      return;
    }
    if (!importAdminPassword.trim()) {
      notifyUser("Enter admin import password before importing farmer master.");
      return;
    }
    const workbook = await readWorkbook(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true
    }) as unknown[][];
    const imported = parseRegistrationWorkbook(rows);
    const response = await fetchWithAuth(`${API_BASE}/api/seed/import/registrations`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        fileName: file.name,
        seasonLabel: reportSeasonLabel,
        importPassword: importAdminPassword,
        registrations: imported
      })
    });

    if (!response.ok) {
      throw new Error("Unable to import registrations into MongoDB.");
    }

    const data = (await response.json()) as AppBootstrapPayload;

    setRegistrations(data.registrations ?? []);
    setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
    setStacks(data.stacks?.length ? data.stacks : defaultStacks);
    setLots(data.lots ?? []);
    setDiscrepancies(data.discrepancies ?? []);
    setDiscrepancyShifts(data.discrepancyShifts ?? []);
    setStackAccommodations(data.stackAccommodations ?? []);
    setFinancialVouchers(data.financialVouchers ?? []);
    setReceipts(data.receipts ?? []);
    setFeatures({
      discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
    });
    setSelectedRegistrationId(data.registrations?.[0]?.id ?? "");
    setReceiptNo(nextReceiptNo(data.receipts ?? []));
    setImportPage(1);
    setImportMessage(`${imported.length} registration records imported successfully into MongoDB.`);
    setImportAdminPassword("");
    notifyUser("Farmer master imported.");
  }

  async function handleReportTemplateImport(file: File) {
    const workbook = await readWorkbook(file);
    const summarySheet = workbook.Sheets[workbook.SheetNames[2]];
    const rows = XLSX.utils.sheet_to_json(summarySheet, {
      header: 1,
      raw: true
    }) as unknown[][];
    const meta = parseReportWorkbookMeta(rows);
    setReportSeasonLabel(`${meta.season} ${meta.year}`);
    notifyUser("Report template linked.");
  }

  function addGodown() {
    if (!requirePermission("canImport", "Only Admin can manage godown master.")) {
      return;
    }
    if (!newGodownName.trim()) {
      return;
    }
    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/masters/godowns`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          name: newGodownName.trim()
        })
      });

      if (!response.ok) {
        throw new Error("Unable to add godown.");
      }

      const created = (await response.json()) as Godown;
      setGodowns((current) => [...current, created]);
      setNewGodownName("");
      notifyUser("Godown added.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to add godown.");
    });
  }

  function addStack() {
    if (!requirePermission("canImport", "Only Admin can manage stack master.")) {
      return;
    }
    if (!newStackNo.trim()) {
      return;
    }
    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/masters/stacks`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          godownId: newStackGodownId,
          stackNo: newStackNo.trim()
        })
      });

      if (!response.ok) {
        throw new Error("Unable to add stack.");
      }

      const created = (await response.json()) as Stack;
      setStacks((current) => [...current, created]);
      setNewStackNo("");
      notifyUser("Stack added.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to add stack.");
    });
  }

  function resetOrganizerForm() {
    setEditingOrganizerId("");
    setNewOrganizerName("");
    setNewOrganizerMobile("");
    setNewOrganizerVillage("");
    setNewOrganizerDistrict("");
    setNewOrganizerRate("");
    setNewOrganizerDeduction("0");
    setNewOrganizerActive(true);
  }

  function openOrganizerForEdit(organizer: Organizer) {
    setEditingOrganizerId(organizer.id);
    setNewOrganizerName(organizer.name);
    setNewOrganizerMobile(organizer.mobile ?? "");
    setNewOrganizerVillage(organizer.village ?? "");
    setNewOrganizerDistrict(organizer.district ?? "");
    setNewOrganizerRate(String(organizer.commissionRatePerQtl ?? 0));
    setNewOrganizerDeduction(String(organizer.deductionAmount ?? 0));
    setNewOrganizerActive(Boolean(organizer.isActive));
    setActiveView("commission");
  }

  function saveOrganizer() {
    if (!requirePermission("canVoucher", "Your role cannot manage organizer commission master.")) {
      return;
    }
    if (!newOrganizerName.trim()) {
      notifyUser("Enter organizer name.");
      return;
    }
    if (!newOrganizerRate.trim() || Number(newOrganizerRate) < 0) {
      notifyUser("Enter a valid commission rate per QTL.");
      return;
    }
    if (!newOrganizerDeduction.trim() || Number(newOrganizerDeduction) < 0) {
      notifyUser("Enter a valid deduction amount.");
      return;
    }

    void (async () => {
      const isEditing = Boolean(editingOrganizerId);
      const response = await fetchWithAuth(
        isEditing
          ? `${API_BASE}/api/seed/masters/organizers/${editingOrganizerId}`
          : `${API_BASE}/api/seed/masters/organizers`,
        {
          method: isEditing ? "PUT" : "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            name: newOrganizerName.trim(),
            mobile: newOrganizerMobile.trim(),
            village: newOrganizerVillage.trim(),
            district: newOrganizerDistrict.trim(),
            commissionRatePerQtl: Number(newOrganizerRate),
            deductionAmount: Number(newOrganizerDeduction),
            isActive: newOrganizerActive
          })
        }
      );
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to save organizer.");
      }
      await loadBootstrap();
      resetOrganizerForm();
      notifyUser(isEditing ? "Organizer updated successfully." : "Organizer created successfully.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to save organizer.");
    });
  }

  function deleteOrganizer(organizer: Organizer) {
    if (!requirePermission("canDelete", "Only Admin can delete organizer master.")) {
      return;
    }
    const confirmed = confirmDestructiveAction({
      itemLabel: `Organizer: ${organizer.name}`
    });
    if (!confirmed) {
      notifyUser("Delete cancelled. No organizer was deleted.", false);
      return;
    }

    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/masters/organizers/${organizer.id}`, {
        method: "DELETE"
      });
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to delete organizer.");
      }
      await loadBootstrap();
      if (editingOrganizerId === organizer.id) {
        resetOrganizerForm();
      }
      notifyUser("Organizer deleted successfully.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to delete organizer.");
    });
  }

  function openOrganizerAssignment(registrationId: string) {
    setOrganizerAssignmentRegistrationId(registrationId);
    const registration = registrations.find((item) => item.id === registrationId);
    setOrganizerAssignmentOrganizerId(registration?.organizerId ?? "");
  }

  function closeOrganizerAssignment() {
    setOrganizerAssignmentRegistrationId("");
    setOrganizerAssignmentOrganizerId("");
  }

  function saveOrganizerAssignment() {
    if (!requirePermission("canEdit", "Your role cannot assign organizer to farmer registration.")) {
      return;
    }
    if (!organizerAssignmentRegistrationId) {
      notifyUser("Select a farmer registration first.");
      return;
    }

    void (async () => {
      const response = await fetchWithAuth(
        `${API_BASE}/api/seed/registrations/${organizerAssignmentRegistrationId}/organizer`,
        {
          method: "PUT",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            organizerId: organizerAssignmentOrganizerId
          })
        }
      );
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to assign organizer.");
      }
      await loadBootstrap();
      closeOrganizerAssignment();
      notifyUser("Organizer mapping saved successfully.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to assign organizer.");
    });
  }

  function openOrganizerLedger(organizerId: string) {
    setOrganizerLedgerOrganizerId(organizerId);
    setOrganizerPaymentDate(new Date().toISOString().slice(0, 10));
    setOrganizerPaymentAmount("");
    setOrganizerPaymentTransactionNo("");
    setOrganizerPaymentRemarks("");
    setEditingOrganizerPaymentId("");
  }

  function closeOrganizerLedger() {
    setOrganizerLedgerOrganizerId("");
    setOrganizerPaymentDate(new Date().toISOString().slice(0, 10));
    setOrganizerPaymentAmount("");
    setOrganizerPaymentTransactionNo("");
    setOrganizerPaymentRemarks("");
    setEditingOrganizerPaymentId("");
  }

  function beginEditOrganizerPayment(payment: OrganizerPayment) {
    setEditingOrganizerPaymentId(payment.id);
    setOrganizerPaymentDate(payment.paymentDate);
    setOrganizerPaymentAmount(String(payment.amount));
    setOrganizerPaymentTransactionNo(payment.transactionNo);
    setOrganizerPaymentRemarks(payment.remarks ?? "");
  }

  function saveOrganizerPayment() {
    if (!requirePermission("canVoucher", "Your role cannot record organizer commission payments.")) {
      return;
    }
    if (!organizerLedgerOrganizerId) {
      notifyUser("Open organizer ledger first.");
      return;
    }
    if (!organizerPaymentDate) {
      notifyUser("Select payment date.");
      return;
    }
    if (!organizerPaymentAmount.trim() || Number(organizerPaymentAmount) <= 0) {
      notifyUser("Enter a valid payment amount.");
      return;
    }
    if (!organizerPaymentTransactionNo.trim()) {
      notifyUser("Transaction number is required.");
      return;
    }

    void (async () => {
      const isEditing = Boolean(editingOrganizerPaymentId);
      const response = await fetchWithAuth(
        isEditing
          ? `${API_BASE}/api/seed/organizer-payments/${editingOrganizerPaymentId}`
          : `${API_BASE}/api/seed/organizer-payments`,
        {
          method: isEditing ? "PUT" : "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            organizerId: organizerLedgerOrganizerId,
            paymentDate: organizerPaymentDate,
            amount: Number(organizerPaymentAmount),
            transactionNo: organizerPaymentTransactionNo.trim(),
            remarks: organizerPaymentRemarks.trim()
          })
        }
      );
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to save organizer commission payment.");
      }
      await loadOperationalBootstrap();
      setOrganizerPaymentDate(new Date().toISOString().slice(0, 10));
      setOrganizerPaymentAmount("");
      setOrganizerPaymentTransactionNo("");
      setOrganizerPaymentRemarks("");
      setEditingOrganizerPaymentId("");
      notifyUser(isEditing ? "Organizer commission payment updated." : "Organizer commission payment recorded.");
    })().catch((error) => {
      notifyUser(
        error instanceof Error ? error.message : "Unable to save organizer commission payment."
      );
    });
  }

  function deleteOrganizerPayment(payment: OrganizerPayment) {
    if (!requirePermission("canVoucher", "Your role cannot delete organizer commission payments.")) {
      return;
    }
    const confirmed = confirmDestructiveAction({
      itemLabel: `Transaction No.: ${payment.transactionNo}`
    });
    if (!confirmed) {
      notifyUser("Delete cancelled. No organizer commission payment was deleted.", false);
      return;
    }

    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/organizer-payments/${payment.id}`, {
        method: "DELETE"
      });
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to delete organizer commission payment.");
      }
      await loadOperationalBootstrap();
      if (editingOrganizerPaymentId === payment.id) {
        setOrganizerPaymentDate(new Date().toISOString().slice(0, 10));
        setOrganizerPaymentAmount("");
        setOrganizerPaymentTransactionNo("");
        setOrganizerPaymentRemarks("");
        setEditingOrganizerPaymentId("");
      }
      notifyUser("Organizer commission payment deleted.");
    })().catch((error) => {
      notifyUser(
        error instanceof Error ? error.message : "Unable to delete organizer commission payment."
      );
    });
  }

  function updateDraftLine(
    lineId: string,
    field: keyof Omit<IntakeReceiptLine, "allocations">,
    value: string
  ) {
    setDraftLines((current) =>
      current.map((line) => {
        if (line.id !== lineId) {
          return line;
        }

        const nextLine = {
          ...line,
          [field]:
            field === "grossWeightQtl" || field === "noOfBags" || field === "moisturePercent"
              ? Number(value)
              : value
        } as Omit<IntakeReceiptLine, "allocations">;

        const netWeightQtl = calculateNetWeightQtl(
          Number(nextLine.grossWeightQtl ?? 0),
          Number(nextLine.noOfBags ?? 0)
        );
        const weightPerBagKg = calculateWeightPerBagKg(
          netWeightQtl,
          Number(nextLine.noOfBags ?? 0)
        );

        return {
          ...nextLine,
          netWeightQtl,
          weightPerBagKg,
          qtyQtl: netWeightQtl
        };
      })
    );
  }

  function addDraftLine() {
    setDraftLines((current) => [...current, createReceiptLine(godowns, stacks)]);
  }

  function removeDraftLine(lineId: string) {
    setDraftLines((current) =>
      current.length === 1 ? current : current.filter((line) => line.id !== lineId)
    );
  }

  function saveReceipt() {
    if (
      !requirePermission(
        editingReceiptNo ? "canEdit" : "canEntry",
        editingReceiptNo
          ? "Your role cannot edit intake entries."
          : "Your role cannot create intake entries."
      )
    ) {
      return;
    }
    if (!selectedRegistration) {
      notifyUser("Select a registration first.");
      return;
    }
    if (isSavingReceipt) {
      notifyUser("Receipt save is already in progress. Please wait.");
      return;
    }

    const isEditing = Boolean(editingReceiptNo);
    const currentReceiptNo = receiptNo;

    void (async () => {
      setIsSavingReceipt(true);
      const totalDraftNetWeight = draftLines.reduce(
        (sum, line) => sum + Number(line.netWeightQtl ?? 0),
        0
      );

      if (!features.discrepancyWorkflow && totalDraftNetWeight > availableBalanceForSave) {
        throw new Error("Total net weight cannot exceed expected available balance.");
      }

      for (const line of draftLines) {
        const godown = godowns.find((item) => item.id === line.godownId);
        if (!godown) {
          throw new Error("Select godown for each intake line.");
        }
        if (!line.stackNo.trim()) {
          throw new Error("Enter stack number for each intake line.");
        }
        if (!features.discrepancyWorkflow && line.netWeightQtl > availableBalanceForSave) {
          throw new Error("Net weight cannot exceed expected available balance.");
        }
      }

      const url = editingReceiptNo
        ? `${API_BASE}/api/seed/intake/receipts/${editingReceiptNo}`
        : `${API_BASE}/api/seed/intake/receipts`;
      const method = editingReceiptNo ? "PUT" : "POST";

      const finalResponse = await fetchWithAuth(url, {
        method,
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          receiptNo,
          receiptDate,
          cropRegistrationId: selectedRegistration.id,
          lines: draftLines
        })
      });

      if (!finalResponse.ok) {
        const errorBody = await finalResponse.json().catch(() => null);
        throw new Error(
          errorBody?.message ||
            "Unable to save intake receipt into MongoDB."
        );
      }

      await finalResponse.json();
      const data = (await fetchWithAuth(`${API_BASE}/api/seed/bootstrap`).then((response) => {
        if (!response.ok) {
          throw new Error("Unable to refresh saved receipt data.");
        }

        return response.json();
      })) as AppBootstrapPayload;

      setRegistrations(data.registrations ?? []);
      const nextGodowns = data.godowns?.length ? data.godowns : defaultGodowns;
      const nextStacks = data.stacks?.length ? data.stacks : defaultStacks;
      setGodowns(nextGodowns);
      setStacks(nextStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setStackAccommodations(data.stackAccommodations ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      const savedReceiptNo = currentReceiptNo;
      const savedReceiptDiscrepancies = (data.discrepancies ?? []).filter(
        (item) => item.receiptNo === savedReceiptNo
      );
      setReceiptNo(nextReceiptNo(data.receipts ?? []));
      setReceiptDate(new Date().toISOString().slice(0, 10));
      setDraftLines([createReceiptLine(nextGodowns, nextStacks)]);
      setEditingReceiptNo("");
      if (savedReceiptDiscrepancies.length > 0) {
        const totalExcessQty = savedReceiptDiscrepancies.reduce(
          (sum, item) => sum + Number(item.excessQtyQtl ?? 0),
          0
        );
        notifyUser(
          `Entry saved.\nDiscrepancy found for receipt ${savedReceiptNo}.\nExcess quantity: ${formatNumber(totalExcessQty)} QTL.`
        );
      } else {
        notifyUser("Entry saved.");
      }
      setActiveView("registrations");
      setToast(
        isEditing
          ? `Entry saved. Receipt ${currentReceiptNo} updated.`
          : `Entry saved. Receipt ${currentReceiptNo} saved.`
      );
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to save receipt.");
    }).finally(() => {
      setIsSavingReceipt(false);
    });
  }

  function startEditReceipt(receipt: IntakeReceipt) {
    if (!requirePermission("canEdit", "Your role cannot edit intake entries.")) {
      return;
    }
    setEditingReceiptNo(receipt.receiptNo);
    setReceiptNo(receipt.receiptNo);
    setReceiptDate(receipt.receiptDate);
    setSelectedRegistrationId(receipt.cropRegistrationId);
    setDraftLines(
        receipt.lines.map((line) => ({
          id: line.id,
          godownId: line.godownId,
          stackId: line.stackId,
          stackNo: line.stackNo,
          grossWeightQtl: line.grossWeightQtl,
          qtyQtl: line.qtyQtl,
          noOfBags: line.noOfBags,
          weightPerBagKg: line.weightPerBagKg,
          netWeightQtl: line.netWeightQtl,
          moisturePercent: line.moisturePercent,
          vehicleNo: line.vehicleNo,
          remarks: line.remarks
        }))
    );
    setActiveView("intake");
    setToast(`Editing receipt ${receipt.receiptNo}.`);
  }

  function startEditReceiptFromDeposit(receiptNoToEdit: string) {
    const receipt = receipts.find((item) => item.receiptNo === receiptNoToEdit);
    if (!receipt) {
      notifyUser("Receipt not found.");
      return;
    }
    setDepositViewRegistrationId("");
    startEditReceipt(receipt);
  }

  function deleteReceipt(receiptRefToDelete: string, receiptNoToDelete: string) {
    if (!requirePermission("canDelete", "Only Admin can delete intake entries.")) {
      return;
    }
    const confirmed = confirmDestructiveAction({
      itemLabel: `Receipt No.: ${receiptNoToDelete}`
    });
    if (!confirmed) {
      notifyUser("Delete cancelled. No receipt was deleted.", false);
      return;
    }
    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/intake/receipts/${receiptRefToDelete}`, {
        method: "DELETE"
      });
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to delete receipt.");
      }
      const data = (await response.json()) as AppBootstrapPayload;
      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setStackAccommodations(data.stackAccommodations ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setReceiptNo(nextReceiptNo(data.receipts ?? []));
      if (editingReceiptNo === receiptNoToDelete) {
        setEditingReceiptNo("");
        setDraftLines([createReceiptLine(godowns, stacks)]);
      }
      notifyUser(`Receipt ${receiptNoToDelete} deleted.`);
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to delete receipt.");
    });
  }

  function goToIntakeForRegistration(registrationId: string) {
    setSelectedRegistrationId(registrationId);
    setEditingReceiptNo("");
    setDraftLines([createReceiptLine(godowns, stacks)]);
    setReceiptNo(nextReceiptNo(receipts));
    setReceiptDate(new Date().toISOString().slice(0, 10));
    setActiveView("intake");
    setToast("Registration loaded into Intake Entry.");
  }

  function startShiftEntry(discrepancy: IntakeDiscrepancy) {
    if (!requirePermission("canShift", "Your role cannot enter discrepancy shift.")) {
      return;
    }
    setSelectedDiscrepancyId(discrepancy.id);
    setDiscrepancyWorkflowMode("shift");
    setShiftTargetGodownId(godowns[0]?.id ?? "");
    setShiftTargetStackNo("");
    setShiftQtyQtl(String(discrepancy.excessQtyQtl));
    setShiftBags(String(discrepancy.estimatedExcessBags));
    setShiftDate(new Date().toISOString().slice(0, 10));
    setShiftApprovedBy("");
    setShiftRemarks("");
    setToast(`Shift entry opened for ${discrepancy.discrepancyNo}.`);
  }

  function saveStackAccommodation() {
    if (!requirePermission("canShift", "Your role cannot save stack accommodation.")) {
      return;
    }
    if (!selectedDiscrepancy) {
      notifyUser("Select a discrepancy first.");
      return;
    }

    void (async () => {
      const adjustedQty = Number(accommodationQtyQtl);
      const adjustedBagCount = Number(accommodationBags || 0);

      if (!accommodationTargetRegistrationId) {
        throw new Error("Select target farmer in the same stack.");
      }
      if (!Number.isFinite(adjustedQty) || adjustedQty <= 0) {
        throw new Error("Enter valid accommodation quantity.");
      }
      if (!Number.isFinite(adjustedBagCount) || adjustedBagCount < 0) {
        throw new Error("Enter valid accommodation bags.");
      }

      const response = await fetchWithAuth(
        editingAccommodationId
          ? `${API_BASE}/api/seed/discrepancies/accommodations/${editingAccommodationId}`
          : `${API_BASE}/api/seed/discrepancies/accommodations`,
        {
        method: editingAccommodationId ? "PUT" : "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          discrepancyId: selectedDiscrepancy.id,
          targetRegistrationId: accommodationTargetRegistrationId,
          adjustedQtyQtl: adjustedQty,
          adjustedBags: adjustedBagCount,
          adjustmentDate: accommodationDate,
          remarks: accommodationRemarks,
          adminPassword: ""
        })
      });

      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to save stack accommodation.");
      }

      const data = (await response.json()) as AppBootstrapPayload;

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setStackAccommodations(data.stackAccommodations ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setEditingAccommodationId("");
      setAccommodationTargetRegistrationId("");
      setAccommodationQtyQtl("");
      setAccommodationBags("");
      setAccommodationDate(new Date().toISOString().slice(0, 10));
      setAccommodationRemarks("");
      notifyUser(
        editingAccommodationId
          ? "Stack accommodation updated. This mapping only affects accommodation register view."
          : "Stack accommodation saved. This mapping only affects accommodation register view."
      );
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to save stack accommodation.");
    });
  }

  function beginEditStackAccommodation(accommodation: StackAccommodation) {
    setSelectedDiscrepancyId(accommodation.discrepancyId);
    setDiscrepancyWorkflowMode("accommodation");
    setEditingAccommodationId(accommodation.id);
    setAccommodationTargetRegistrationId(accommodation.targetRegistrationId);
    setAccommodationQtyQtl(String(accommodation.adjustedQtyQtl));
    setAccommodationBags(String(accommodation.adjustedBags));
    setAccommodationDate(accommodation.adjustmentDate);
    setAccommodationRemarks(accommodation.remarks ?? "");
    setToast(`Accommodation edit opened for ${accommodation.sourceRegistrationCode} -> ${accommodation.targetRegistrationCode}.`);
  }

  function deleteStackAccommodation(accommodation: StackAccommodation) {
    if (!requirePermission("canShift", "Your role cannot delete stack accommodation.")) {
      return;
    }
    if (
      !confirmDestructiveAction({
        itemLabel: `Accommodation: ${accommodation.sourceRegistrationCode} -> ${accommodation.targetRegistrationCode}`
      })
    ) {
      return;
    }

    void (async () => {
      const response = await fetchWithAuth(
        `${API_BASE}/api/seed/discrepancies/accommodations/${accommodation.id}`,
        {
          method: "DELETE",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({ adminPassword: "" })
        }
      );
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to delete stack accommodation.");
      }
      const data = (await response.json()) as AppBootstrapPayload;
      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setStackAccommodations(data.stackAccommodations ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      if (editingAccommodationId === accommodation.id) {
        setEditingAccommodationId("");
        setAccommodationTargetRegistrationId("");
        setAccommodationQtyQtl("");
        setAccommodationBags("");
        setAccommodationDate(new Date().toISOString().slice(0, 10));
        setAccommodationRemarks("");
      }
      notifyUser("Stack accommodation deleted.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to delete stack accommodation.");
    });
  }

  function downloadAdjustedStackCardExcel() {
    if (!adjustedStackCardPreview) {
      notifyUser("Select a discrepancy with stack accommodation data first.");
      return;
    }

    const originalTotalQty = roundQtl(
      adjustedStackCardPreview.originalRows.reduce((sum, row) => sum + Number(row.qtyQtl ?? 0), 0)
    );
    const originalTotalBags = adjustedStackCardPreview.originalRows.reduce(
      (sum, row) => sum + Number(row.bags ?? 0),
      0
    );
    const finalTotalQty = roundQtl(
      adjustedStackCardPreview.adjustedRows.reduce((sum, row) => sum + Number(row.finalQtyQtl ?? 0), 0)
    );
    const finalTotalBags = adjustedStackCardPreview.adjustedRows.reduce(
      (sum, row) => sum + Number(row.finalBags ?? 0),
      0
    );

    const rows: (string | number)[][] = [
      [COMPANY_NAME],
      [adjustedStackCardPreview.title],
      [],
      ["Godown", adjustedStackCardPreview.godownName],
      ["Stack No.", adjustedStackCardPreview.stackNo],
      ["Generated On", new Date(adjustedStackCardPreview.generatedAt).toLocaleString("en-IN")],
      ["Note", "For reference only. Does not alter main stock records."],
      [],
      ["Original Stack Farmer Rows"],
      ["S. No.", "Reg. Code", "Farmer Name", "Village", "Qty (QTL)", "Bags"],
      ...adjustedStackCardPreview.originalRows.map((row, index) => [
        index + 1,
        row.regCode,
        row.farmerName,
        row.village,
        row.qtyQtl,
        row.bags
      ]),
      ["", "", "", "TOTAL", originalTotalQty, originalTotalBags],
      [],
      ["Final Adjusted Farmer-wise Stack Position"],
      ["S. No.", "Reg. Code", "Farmer Name", "Village", "Qty (QTL)", "Bags", "Mark"],
      ...adjustedStackCardPreview.adjustedRows.map((row, index) => [
        index + 1,
        row.regCode,
        row.farmerName,
        row.village,
        row.finalQtyQtl,
        row.finalBags,
        row.changed ? "*" : ""
      ]),
      ["", "", "", "TOTAL", finalTotalQty, finalTotalBags, ""],
      [],
      ["Change Summary"],
      ["Changed Farmers", adjustedStackCardPreview.changedFarmerCount],
      ["Accommodated Qty (QTL)", adjustedStackCardPreview.totalAccommodatedQtyQtl],
      ["Accommodated Bags", adjustedStackCardPreview.totalAccommodatedBags],
      [],
      ["* Marked rows indicate farmers affected in accommodation."]
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Adjusted Stack Card");
    XLSX.writeFile(workbook, adjustedStackCardPreview.fileName);
    notifyUser("Adjusted stack card Excel downloaded.");
  }

  function downloadAdjustedStackCardPdf() {
    if (!adjustedStackCardPreview) {
      notifyUser("Select a discrepancy with stack accommodation data first.");
      return;
    }

    const originalTotalQty = roundQtl(
      adjustedStackCardPreview.originalRows.reduce((sum, row) => sum + Number(row.qtyQtl ?? 0), 0)
    );
    const originalTotalBags = adjustedStackCardPreview.originalRows.reduce(
      (sum, row) => sum + Number(row.bags ?? 0),
      0
    );
    const finalTotalQty = roundQtl(
      adjustedStackCardPreview.adjustedRows.reduce((sum, row) => sum + Number(row.finalQtyQtl ?? 0), 0)
    );
    const finalTotalBags = adjustedStackCardPreview.adjustedRows.reduce(
      (sum, row) => sum + Number(row.finalBags ?? 0),
      0
    );

    const pdf = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4"
    });
    const pageWidth = pdf.internal.pageSize.getWidth();
    let y = renderPdfBrandHeader(pdf, adjustedStackCardPreview.title, {
      left: 14,
      right: pageWidth - 14,
      y: 10,
      compact: false,
      logoDataUrl: pdfLogoDataUrl
    });

    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(8.5);
    pdf.text(`Godown: ${adjustedStackCardPreview.godownName}`, 14, y);
    pdf.text(`Stack No.: ${adjustedStackCardPreview.stackNo}`, 95, y);
    pdf.text(
      `Generated: ${new Date(adjustedStackCardPreview.generatedAt).toLocaleString("en-IN")}`,
      170,
      y,
      { align: "right" }
    );
    y += 5;
    pdf.text("For reference only. Does not alter main stock records.", 14, y);

    autoTable(pdf, {
      startY: y + 4,
      head: [["Original Stack Position", "", "", "", "", ""]],
      body: [
        ["S. No.", "Reg. Code", "Farmer Name", "Village", "Qty (QTL)", "Bags"],
        ...adjustedStackCardPreview.originalRows.map((row, index) => [
          String(index + 1),
          row.regCode,
          row.farmerName,
          row.village,
          formatNumber(row.qtyQtl),
          String(row.bags)
        ]),
        ["", "", "", "TOTAL", formatNumber(originalTotalQty), String(originalTotalBags)]
      ],
      theme: "grid",
      styles: {
        font: "helvetica",
        fontSize: 7,
        cellPadding: 1.5,
        textColor: [47, 30, 18]
      },
      headStyles: {
        fillColor: [127, 75, 38],
        fontSize: 7.5,
        halign: "left"
      },
      bodyStyles: {
        valign: "middle"
      },
      margin: { left: 14, right: 14 }
    });

    const originalTable = pdf as jsPDF & { lastAutoTable?: { finalY?: number } };
    autoTable(pdf, {
      startY: (originalTable.lastAutoTable?.finalY ?? y + 4) + 5,
      head: [["S. No.", "Reg. Code", "Farmer Name", "Village", "Qty (QTL)", "Bags", "Mark"]],
      body: [
        ...adjustedStackCardPreview.adjustedRows.map((row, index) => [
          String(index + 1),
          row.regCode,
          row.farmerName,
          row.village,
          formatNumber(row.finalQtyQtl),
          String(row.finalBags),
          row.changed ? "*" : ""
        ]),
        ["", "", "", "TOTAL", formatNumber(finalTotalQty), String(finalTotalBags), ""]
      ],
      theme: "grid",
      styles: {
        font: "helvetica",
        fontSize: 6.5,
        cellPadding: 1.3,
        overflow: "linebreak",
        textColor: [47, 30, 18]
      },
      headStyles: {
        fillColor: [48, 86, 61],
        fontSize: 7,
        halign: "center"
      },
      bodyStyles: {
        valign: "middle"
      },
      margin: { left: 14, right: 14, bottom: 10 }
    });

    const finalTable = pdf as jsPDF & { lastAutoTable?: { finalY?: number } };
    autoTable(pdf, {
      startY: (finalTable.lastAutoTable?.finalY ?? y + 4) + 5,
      head: [["Change Summary", "Value"]],
      body: [
        ["Changed Farmers", String(adjustedStackCardPreview.changedFarmerCount)],
        ["Accommodated Qty", `${formatNumber(adjustedStackCardPreview.totalAccommodatedQtyQtl)} QTL`],
        ["Accommodated Bags", String(adjustedStackCardPreview.totalAccommodatedBags)],
        ["Note", "* Marked rows indicate farmers affected in accommodation."]
      ],
      theme: "grid",
      styles: {
        font: "helvetica",
        fontSize: 7,
        cellPadding: 1.5,
        textColor: [47, 30, 18]
      },
      headStyles: {
        fillColor: [94, 63, 45],
        fontSize: 7.2
      },
      margin: { left: 14, right: 120, bottom: 10 },
      tableWidth: 75
    });

    pdf.save(adjustedStackCardPreview.fileName.replace(/\.xlsx$/i, ".pdf"));
    notifyUser("Adjusted stack card PDF downloaded.");
  }

  function saveDiscrepancyShift() {
    if (!requirePermission("canShift", "Your role cannot save discrepancy shift.")) {
      return;
    }
    if (!selectedDiscrepancy) {
      notifyUser("Select a discrepancy first.");
      return;
    }

    void (async () => {
      const shiftedQty = Number(shiftQtyQtl);
      const shiftedBagCount = Number(shiftBags);

      if (!shiftTargetGodownId) {
        throw new Error("Select target godown.");
      }
      if (!shiftTargetStackNo.trim()) {
        throw new Error("Enter target non-certification stack.");
      }
      if (!Number.isFinite(shiftedQty) || shiftedQty <= 0) {
        throw new Error("Enter valid shifted quantity.");
      }
      if (!Number.isFinite(shiftedBagCount) || shiftedBagCount <= 0) {
        throw new Error("Enter valid shifted bags.");
      }

      const response = await fetchWithAuth(`${API_BASE}/api/seed/discrepancies/shifts`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          discrepancyId: selectedDiscrepancy.id,
          toGodownId: shiftTargetGodownId,
          toStackNo: shiftTargetStackNo.trim(),
          shiftedQtyQtl: shiftedQty,
          shiftedBags: shiftedBagCount,
          shiftDate,
          approvedBy: shiftApprovedBy,
          remarks: shiftRemarks
        })
      });

      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to save discrepancy shift.");
      }

      const data = (await response.json()) as AppBootstrapPayload;

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setStackAccommodations(data.stackAccommodations ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setSelectedDiscrepancyId("");
      setShiftTargetStackNo("");
      setShiftQtyQtl("");
      setShiftBags("");
      setShiftApprovedBy("");
      setShiftRemarks("");
      notifyUser("Discrepancy shift saved.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to save discrepancy shift.");
    });
  }

  function runDiscrepancyValidation() {
    if (!requirePermission("canValidate", "Only Admin can run discrepancy validation.")) {
      return;
    }
    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/validations/discrepancies`, {
        method: "POST"
      });
      if (!response.ok) {
        throw new Error("Unable to run discrepancy auto validation.");
      }

      const data = (await response.json()) as AppBootstrapPayload;

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setStackAccommodations(data.stackAccommodations ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setValidationSummary((data.validationSummary as ValidationSummary) ?? null);
      notifyUser("Discrepancy auto validation completed.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to run discrepancy validation.");
    });
  }

  function runLotValidation() {
    if (!requirePermission("canValidate", "Only Admin can run lot validation.")) {
      return;
    }
    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/validations/lots`, {
        method: "POST"
      });
      if (!response.ok) {
        throw new Error("Unable to run lot auto validation.");
      }

      const data = (await response.json()) as AppBootstrapPayload;

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setStackAccommodations(data.stackAccommodations ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setValidationSummary((data.validationSummary as ValidationSummary) ?? null);
      notifyUser("Lot auto validation completed.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to run lot validation.");
    });
  }

  function runLotReindex() {
    if (!requirePermission("canValidate", "Only Admin can reindex lot sequence.")) {
      return;
    }
    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/validations/lots/reindex`, {
        method: "POST"
      });
      if (!response.ok) {
        throw new Error("Unable to reindex lot sequence.");
      }

      const data = (await response.json()) as AppBootstrapPayload;

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setStackAccommodations(data.stackAccommodations ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setValidationSummary((data.validationSummary as ValidationSummary) ?? null);
      notifyUser("Lot sequence reindexed successfully.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to reindex lot sequence.");
    });
  }

  function buildReportRequest() {
    return {
      reportType,
      season: reportSeasonLabel,
      fromDate: reportFromDate,
      toDate: reportToDate,
      crop: reportCrop,
      variety: reportVariety,
      classStage: reportClassStage,
      district: reportDistrict,
      godownId: reportGodownId,
      stackNo: reportStackNo,
      cropRegistrationCode: reportRegistrationCode,
      farmerName: reportFarmerName,
      reportMode,
      includeDiscrepancy: reportMode !== "ACCEPTED_ONLY"
    };
  }

  function buildStackCardRegisterSections(): StackCardRegisterSection[] {
    const fromDate = reportFromDate.trim();
    const toDate = reportToDate.trim();
    const districtFilter = reportDistrict.trim().toLowerCase();
    const godownFilter = reportGodownId.trim();
    const stackFilter = reportStackNo.trim().toLowerCase();
    const registrationFilter = reportRegistrationCode.trim().toLowerCase();
    const farmerFilter = reportFarmerName.trim().toLowerCase();
    const cropFilter = reportCrop.trim().toLowerCase();
    const varietyFilter = reportVariety.trim().toLowerCase();
    const classFilter = reportClassStage.trim().toLowerCase();

    const registrationLookup = new Map(registrations.map((item) => [item.id, item] as const));
    const godownLookup = new Map(godowns.map((item) => [item.id, item.name] as const));
    const eligibleStackKeys = new Set<string>();
    const stackRows = new Map<
      string,
      {
        godownId: string;
        godownName: string;
        stackNo: string;
        rows: Map<
          string,
          {
            regCode: string;
            farmerName: string;
            village: string;
            district: string;
            qtyQtl: number;
            bags: number;
          }
        >;
      }
    >();

    for (const receipt of receipts) {
      if (fromDate && receipt.receiptDate < fromDate) {
        continue;
      }
      if (toDate && receipt.receiptDate > toDate) {
        continue;
      }
      if (registrationFilter && !receipt.cropRegistrationCode.toLowerCase().includes(registrationFilter)) {
        continue;
      }
      if (farmerFilter && !receipt.farmerName.toLowerCase().includes(farmerFilter)) {
        continue;
      }

      const registration = registrationLookup.get(receipt.cropRegistrationId);
      if (districtFilter && !(registration?.district || "").toLowerCase().includes(districtFilter)) {
        continue;
      }
      if (cropFilter && !(registration?.crop || "").toLowerCase().includes(cropFilter)) {
        continue;
      }
      if (varietyFilter && !(registration?.variety || "").toLowerCase().includes(varietyFilter)) {
        continue;
      }
      if (classFilter && !(registration?.classStage || "").toLowerCase().includes(classFilter)) {
        continue;
      }

      for (const line of receipt.lines) {
        const normalizedStackNo = String(line.stackNo ?? "").trim();
        if (!normalizedStackNo) {
          continue;
        }
        if (godownFilter && line.godownId !== godownFilter) {
          continue;
        }
        if (stackFilter && normalizedStackNo.toLowerCase() !== stackFilter) {
          continue;
        }

        const key = `${line.godownId}::${normalizedStackNo}`;
        eligibleStackKeys.add(key);
      }
    }

    for (const receipt of receipts) {
      if (fromDate && receipt.receiptDate < fromDate) {
        continue;
      }
      if (toDate && receipt.receiptDate > toDate) {
        continue;
      }

      const registration = registrationLookup.get(receipt.cropRegistrationId);
      for (const line of receipt.lines) {
        const normalizedStackNo = String(line.stackNo ?? "").trim();
        if (!normalizedStackNo) {
          continue;
        }

        const key = `${line.godownId}::${normalizedStackNo}`;
        if (!eligibleStackKeys.has(key)) {
          continue;
        }

        const group =
          stackRows.get(key) ??
          {
            godownId: line.godownId,
            godownName: godownLookup.get(line.godownId) ?? "Unknown Godown",
            stackNo: normalizedStackNo,
            rows: new Map()
          };
        stackRows.set(key, group);

        const current =
          group.rows.get(receipt.cropRegistrationCode) ??
          {
            regCode: receipt.cropRegistrationCode,
            farmerName: receipt.farmerName,
            village: registration?.village || "-",
            district: registration?.district || "-",
            qtyQtl: 0,
            bags: 0
          };

        current.qtyQtl = roundQtl(current.qtyQtl + Number(line.netWeightQtl ?? line.qtyQtl ?? 0));
        current.bags += Number(line.noOfBags ?? 0);
        group.rows.set(receipt.cropRegistrationCode, current);
      }
    }

    return Array.from(stackRows.values())
      .map((group) => {
        const originalRows = Array.from(group.rows.values()).sort((left, right) =>
          left.regCode.localeCompare(right.regCode, "en", { numeric: true, sensitivity: "base" })
        );
        const sameStackAccommodations = stackAccommodations.filter(
          (item) =>
            item.godownId === group.godownId && String(item.stackNo ?? "").trim() === String(group.stackNo).trim()
        );
        const hasAdjustment = sameStackAccommodations.length > 0;

        const rows = originalRows.map((row) => {
          const adjustedInRows = sameStackAccommodations.filter((item) => item.targetRegistrationCode === row.regCode);
          const adjustedOutRows = sameStackAccommodations.filter(
            (item) => item.sourceRegistrationCode === row.regCode
          );
          const adjustedInQtyQtl = roundQtl(
            adjustedInRows.reduce((sum, item) => sum + Number(item.adjustedQtyQtl ?? 0), 0)
          );
          const adjustedOutQtyQtl = roundQtl(
            adjustedOutRows.reduce((sum, item) => sum + Number(item.adjustedQtyQtl ?? 0), 0)
          );
          const adjustedInBags = adjustedInRows.reduce((sum, item) => sum + Number(item.adjustedBags ?? 0), 0);
          const adjustedOutBags = adjustedOutRows.reduce((sum, item) => sum + Number(item.adjustedBags ?? 0), 0);
          const changed =
            adjustedInQtyQtl > 0 ||
            adjustedOutQtyQtl > 0 ||
            adjustedInBags > 0 ||
            adjustedOutBags > 0;

          return hasAdjustment
            ? {
                regCode: row.regCode,
                farmerName: row.farmerName,
                village: row.village,
                district: row.district,
                qtyQtl: roundQtl(row.qtyQtl + adjustedInQtyQtl - adjustedOutQtyQtl),
                bags: row.bags + adjustedInBags - adjustedOutBags,
                mark: changed ? "*" : ""
              }
            : {
                regCode: row.regCode,
                farmerName: row.farmerName,
                village: row.village,
                district: row.district,
                qtyQtl: roundQtl(row.qtyQtl),
                bags: row.bags,
                mark: ""
              };
        });

        return {
          key: `${group.godownName}::${group.stackNo}`,
          godownId: group.godownId,
          godownName: group.godownName,
          stackNo: group.stackNo,
          displayMode: hasAdjustment ? "FINAL_ADJUSTED" : "ORIGINAL",
          rows,
          totalQtyQtl: roundQtl(rows.reduce((sum, row) => sum + Number(row.qtyQtl ?? 0), 0)),
          totalBags: rows.reduce((sum, row) => sum + Number(row.bags ?? 0), 0),
          changedFarmerCount: rows.filter((row) => row.mark === "*").length
        } satisfies StackCardRegisterSection;
      })
      .sort((left, right) => {
        const godownCompare = left.godownName.localeCompare(right.godownName, "en", { sensitivity: "base" });
        if (godownCompare !== 0) {
          return godownCompare;
        }
        return left.stackNo.localeCompare(right.stackNo, "en", { numeric: true, sensitivity: "base" });
      });
  }

  function buildStackCardRegisterPreview(): ReportPreview {
    const sections = buildStackCardRegisterSections();
    const previewRows = sections.flatMap((section) =>
      section.rows.map((row, index) => ({
        Godown: section.godownName,
        "Stack No.": section.stackNo,
        View: section.displayMode === "FINAL_ADJUSTED" ? "Final Adjusted" : "Original",
        "S.No.": index + 1,
        "Reg. Code": row.regCode,
        "Farmer Name": row.mark === "*" ? `* ${row.farmerName}` : row.farmerName,
        "Changed (*)": row.mark,
        Village: row.village,
        District: row.district,
        "Qty (QTL)": row.qtyQtl,
        Bags: row.bags
      }))
    );

    return {
      reportType: "STACK_CARD_REGISTER",
      title: "Stack Card Register",
      columns: Object.keys(
        previewRows[0] ?? {
          Godown: "",
          "Stack No.": "",
          View: "",
          "S.No.": 1,
          "Reg. Code": "",
          "Farmer Name": "",
          "Changed (*)": "",
          Village: "",
          District: "",
          "Qty (QTL)": 0,
          Bags: 0
        }
      ),
      rows: previewRows,
      totals: {
        "Total Stacks": sections.length,
        "Adjusted Stacks": sections.filter((item) => item.displayMode === "FINAL_ADJUSTED").length,
        "Original Stacks": sections.filter((item) => item.displayMode === "ORIGINAL").length,
        "Changed Farmers": sections.reduce((sum, item) => sum + item.changedFarmerCount, 0),
        "Total Qty (QTL)": formatNumber(sections.reduce((sum, item) => sum + item.totalQtyQtl, 0)),
        "Total Bags": sections.reduce((sum, item) => sum + item.totalBags, 0),
        "Marker Note": "* means this farmer row changed because of stack adjustment.",
        Rule: "Non-adjusted stacks show original sequence. Adjusted stacks show final position with * mark."
      },
      generatedAt: new Date().toISOString(),
      fileName: `stack-card-register-${new Date().toISOString().slice(0, 10)}.xlsx`
    };
  }

  function buildStackCardRegisterWorkbook(sections: StackCardRegisterSection[], preview: ReportPreview) {
    const rows: (string | number)[][] = [
      [COMPANY_NAME],
      [preview.title],
      [`Generated At: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`],
      ["Rule", "Non-adjusted stacks show original sequence. Adjusted stacks show final position with * mark."],
      []
    ];

    sections.forEach((section, index) => {
      rows.push([`${index + 1}. ${section.godownName} - Stack ${section.stackNo}`]);
      rows.push([
        "View",
        section.displayMode === "FINAL_ADJUSTED" ? "Final Adjusted Stack Position" : "Original Stack Position"
      ]);
      rows.push(["S. No.", "Reg. Code", "Farmer Name", "Changed (*)", "Village", "District", "Qty (QTL)", "Bags"]);
      rows.push(
        ...section.rows.map((row, rowIndex) => [
          rowIndex + 1,
          row.regCode,
          row.mark === "*" ? `* ${row.farmerName}` : row.farmerName,
          row.mark,
          row.village,
          row.district,
          row.qtyQtl,
          row.bags
        ])
      );
      rows.push(["", "", "", "", "", "TOTAL", section.totalQtyQtl, section.totalBags]);
      if (section.displayMode === "FINAL_ADJUSTED") {
        rows.push(["", "", "", "", "", "Changed Farmers", section.changedFarmerCount, "*"]);
        rows.push(["", "", "", "", "", "Marker Note", "* means adjusted row", ""]);
      }
      rows.push([]);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Stack Card Register");
    return workbook;
  }

  function buildAdjustedLotFormationRows(): AdjustedLotFormationRow[] {
    const fromDate = reportFromDate.trim();
    const toDate = reportToDate.trim();
    const districtFilter = reportDistrict.trim().toLowerCase();
    const godownFilter = reportGodownId.trim();
    const stackFilter = reportStackNo.trim().toLowerCase();
    const registrationFilter = reportRegistrationCode.trim().toLowerCase();
    const farmerFilter = reportFarmerName.trim().toLowerCase();
    const cropFilter = reportCrop.trim().toLowerCase();
    const varietyFilter = reportVariety.trim().toLowerCase();
    const classFilter = reportClassStage.trim().toLowerCase();

    const registrationLookup = new Map(registrations.map((item) => [item.id, item] as const));
    const registrationCodeLookup = new Map(registrations.map((item) => [item.cropRegistrationCode, item] as const));
    const voucherCodeLookup = financialVouchers.reduce((map, voucher) => {
      const existing = map.get(voucher.cropRegistrationCode) ?? [];
      existing.push(voucher.voucherNo);
      map.set(voucher.cropRegistrationCode, existing);
      return map;
    }, new Map<string, string[]>());
    const godownLookup = new Map(godowns.map((item) => [item.id, item.name] as const));
    const lotRows = new Map<
      string,
      {
        regNo: string;
        farmerName: string;
        warehouseName: string;
        godownId: string;
        stackNumber: string;
        stackKey: string;
        bags: number;
        netWeightQtl: number;
        lotId: string;
        moistureWeightedTotal: number;
        mark: string;
      }
    >();

    for (const receipt of receipts) {
      if (fromDate && receipt.receiptDate < fromDate) {
        continue;
      }
      if (toDate && receipt.receiptDate > toDate) {
        continue;
      }
      if (registrationFilter && !receipt.cropRegistrationCode.toLowerCase().includes(registrationFilter)) {
        continue;
      }
      if (farmerFilter && !receipt.farmerName.toLowerCase().includes(farmerFilter)) {
        continue;
      }

      const registration = registrationLookup.get(receipt.cropRegistrationId);
      if (districtFilter && !(registration?.district || "").toLowerCase().includes(districtFilter)) {
        continue;
      }
      if (cropFilter && !(registration?.crop || "").toLowerCase().includes(cropFilter)) {
        continue;
      }
      if (varietyFilter && !(registration?.variety || "").toLowerCase().includes(varietyFilter)) {
        continue;
      }
      if (classFilter && !(registration?.classStage || "").toLowerCase().includes(classFilter)) {
        continue;
      }

      for (const line of receipt.lines) {
        const normalizedStackNo = String(line.stackNo ?? "").trim();
        if (!normalizedStackNo) {
          continue;
        }
        if (godownFilter && line.godownId !== godownFilter) {
          continue;
        }
        if (stackFilter && normalizedStackNo.toLowerCase() !== stackFilter) {
          continue;
        }

        const lineNet = roundQtl(Number(line.netWeightQtl ?? line.qtyQtl ?? 0));
        if (lineNet <= 0) {
          continue;
        }
        const lineBags = Number(line.noOfBags ?? 0);
        const stackKey = `${line.godownId}::${normalizedStackNo}`;
        const warehouseName = godownLookup.get(line.godownId) ?? "Unknown Godown";

        for (const allocation of line.allocations ?? []) {
          const allocationQty = roundQtl(Number(allocation.qtyQtl ?? 0));
          if (allocationQty <= 0) {
            continue;
          }
          const allocationBags = Math.round(lineBags * (allocationQty / lineNet));
          const lotId = String(allocation.lotCode || allocation.lotId || "").trim();
          if (!lotId) {
            continue;
          }
          const key = `${stackKey}::${lotId}::${receipt.cropRegistrationCode}`;
          const current =
            lotRows.get(key) ??
            {
              regNo: receipt.cropRegistrationCode,
              farmerName: receipt.farmerName,
              warehouseName,
              godownId: line.godownId,
              stackNumber: normalizedStackNo,
              stackKey,
              bags: 0,
              netWeightQtl: 0,
              lotId,
              moistureWeightedTotal: 0,
              mark: ""
            };
          current.bags += allocationBags;
          current.netWeightQtl = roundQtl(current.netWeightQtl + allocationQty);
          current.moistureWeightedTotal += Number(line.moisturePercent ?? 0) * allocationQty;
          lotRows.set(key, current);
        }
      }
    }

    const getOrCreateTargetRow = (
      sourceRow: {
        warehouseName: string;
        godownId: string;
        stackNumber: string;
        stackKey: string;
        lotId: string;
        moistureWeightedTotal: number;
        netWeightQtl: number;
      },
      targetRegNo: string,
      targetFarmerName: string
    ) => {
      const targetExistingInSameStack = Array.from(lotRows.values())
        .filter((row) => row.stackKey === sourceRow.stackKey && row.regNo === targetRegNo)
        .sort((left, right) =>
          left.lotId.localeCompare(right.lotId, "en", { numeric: true, sensitivity: "base" })
        )[0];
      if (targetExistingInSameStack) {
        targetExistingInSameStack.mark = "*";
        return targetExistingInSameStack;
      }

      const key = `${sourceRow.stackKey}::${sourceRow.lotId}::${targetRegNo}`;
      const existing = lotRows.get(key);
      if (existing) {
        return existing;
      }
      const created = {
        regNo: targetRegNo,
        farmerName: targetFarmerName,
        warehouseName: sourceRow.warehouseName,
        godownId: sourceRow.godownId,
        stackNumber: sourceRow.stackNumber,
        stackKey: sourceRow.stackKey,
        bags: 0,
        netWeightQtl: 0,
        lotId: sourceRow.lotId,
        moistureWeightedTotal: 0,
        mark: "*"
      };
      lotRows.set(key, created);
      return created;
    };

    for (const adjustment of stackAccommodations) {
      const stackKey = `${adjustment.godownId}::${String(adjustment.stackNo ?? "").trim()}`;
      let remainingQty = roundQtl(Number(adjustment.adjustedQtyQtl ?? 0));
      let remainingBags = Number(adjustment.adjustedBags ?? 0);
      if (remainingQty <= 0 && remainingBags <= 0) {
        continue;
      }

      const sourceRows = Array.from(lotRows.values())
        .filter(
          (row) =>
            row.stackKey === stackKey &&
            row.regNo === adjustment.sourceRegistrationCode &&
            row.netWeightQtl > 0
        )
        .sort((left, right) =>
          left.lotId.localeCompare(right.lotId, "en", { numeric: true, sensitivity: "base" })
        );

      sourceRows.forEach((sourceRow, index) => {
        if (remainingQty <= 0 && remainingBags <= 0) {
          return;
        }
        const isLast = index === sourceRows.length - 1;
        const movedQty = isLast ? remainingQty : roundQtl(Math.min(remainingQty, sourceRow.netWeightQtl));
        const movedBags = isLast
          ? remainingBags
          : Math.min(
              remainingBags,
              Math.round(Number(adjustment.adjustedBags ?? 0) * (movedQty / Math.max(Number(adjustment.adjustedQtyQtl ?? 0), 0.01)))
            );
        if (movedQty <= 0 && movedBags <= 0) {
          return;
        }

        const sourceMoisture =
          sourceRow.netWeightQtl > 0
            ? roundQtl(sourceRow.moistureWeightedTotal / sourceRow.netWeightQtl)
            : 0;
        sourceRow.netWeightQtl = roundQtl(sourceRow.netWeightQtl - movedQty);
        sourceRow.bags -= movedBags;
        sourceRow.moistureWeightedTotal = roundQtl(sourceRow.moistureWeightedTotal - sourceMoisture * movedQty);
        sourceRow.mark = "*";

        const targetRow = getOrCreateTargetRow(
          sourceRow,
          adjustment.targetRegistrationCode,
          adjustment.targetFarmerName
        );
        targetRow.netWeightQtl = roundQtl(targetRow.netWeightQtl + movedQty);
        targetRow.bags += movedBags;
        targetRow.moistureWeightedTotal = roundQtl(targetRow.moistureWeightedTotal + sourceMoisture * movedQty);
        targetRow.mark = "*";

        remainingQty = roundQtl(remainingQty - movedQty);
        remainingBags -= movedBags;
      });
    }

    const getLotPrefix = (lotId: string, regNo: string) => {
      const normalizedLotId = String(lotId || "").trim();
      const lotMatch = normalizedLotId.match(/^(.*\/L)\d+$/i);
      if (lotMatch?.[1]) {
        return lotMatch[1];
      }
      return `2025-26/${regNo}/L`;
    };

    const capacitySafeRows = new Map<
      string,
      {
        regNo: string;
        farmerName: string;
        warehouseName: string;
        godownId: string;
        stackNumber: string;
        stackKey: string;
        bags: number;
        netWeightQtl: number;
        lotId: string;
        moistureWeightedTotal: number;
        mark: string;
      }
    >();

    Array.from(lotRows.values())
      .filter((row) => row.netWeightQtl > 0.0001 || row.bags > 0)
      .sort((left, right) =>
        left.lotId.localeCompare(right.lotId, "en", { numeric: true, sensitivity: "base" })
      )
      .forEach((row) => {
        let remainingQty = roundQtl(Math.max(0, row.netWeightQtl));
        let remainingBags = Math.max(0, row.bags);
        const totalQty = Math.max(remainingQty, 0.0001);
        const moisturePercent = row.netWeightQtl > 0 ? row.moistureWeightedTotal / row.netWeightQtl : 0;
        const lotPrefix = getLotPrefix(row.lotId, row.regNo);
        let lotNo = Math.max(Number(row.lotId.match(/\/L(\d+)$/i)?.[1] ?? 1), 1);

        while (remainingQty > 0.0001 || remainingBags > 0) {
          const splitQty = remainingQty > LOT_CAPACITY_QTL ? LOT_CAPACITY_QTL : remainingQty;
          const splitBags =
            remainingQty <= LOT_CAPACITY_QTL
              ? remainingBags
              : Math.round(row.bags * (splitQty / totalQty));
          const lotId = `${lotPrefix}${lotNo}`;
          const key = `${row.stackKey}::${lotId}::${row.regNo}`;
          const current =
            capacitySafeRows.get(key) ?? {
              ...row,
              bags: 0,
              netWeightQtl: 0,
              lotId,
              moistureWeightedTotal: 0,
              mark: row.mark
            };
          current.bags += splitBags;
          current.netWeightQtl = roundQtl(current.netWeightQtl + splitQty);
          current.moistureWeightedTotal = roundQtl(
            current.moistureWeightedTotal + moisturePercent * splitQty
          );
          current.mark = current.mark === "*" || row.mark === "*" ? "*" : "";
          capacitySafeRows.set(key, current);

          remainingQty = roundQtl(remainingQty - splitQty);
          remainingBags -= splitBags;
          lotNo += 1;
        }
      });

    return Array.from(capacitySafeRows.values())
      .filter((row) => row.netWeightQtl > 0.0001 || row.bags > 0)
      .sort((left, right) => {
        const warehouseCompare = left.warehouseName.localeCompare(right.warehouseName, "en", {
          sensitivity: "base"
        });
        if (warehouseCompare !== 0) {
          return warehouseCompare;
        }
        const stackCompare = left.stackNumber.localeCompare(right.stackNumber, "en", {
          numeric: true,
          sensitivity: "base"
        });
        if (stackCompare !== 0) {
          return stackCompare;
        }
        const lotCompare = left.lotId.localeCompare(right.lotId, "en", {
          numeric: true,
          sensitivity: "base"
        });
        if (lotCompare !== 0) {
          return lotCompare;
        }
        return left.regNo.localeCompare(right.regNo, "en", { numeric: true, sensitivity: "base" });
      })
      .map((row, index) => ({
        ...(() => {
          const registration = registrationCodeLookup.get(row.regNo);
          return {
            paymentVoucherNo: voucherCodeLookup.get(row.regNo)?.join(", ") || "Not Generated",
            expectedYieldQtl: roundQtl(Number(registration?.expectedYieldQtl ?? 0)),
            totalNetIntakeQtl: roundQtl(Number(registration?.totalReceivedQtl ?? 0))
          };
        })(),
        srNo: index + 1,
        regNo: row.regNo,
        farmerName: row.mark === "*" ? `* ${row.farmerName}` : row.farmerName,
        warehouseName: row.warehouseName,
        stackNumber: row.stackNumber,
        bags: Math.max(0, row.bags),
        netWeightQtl: roundQtl(Math.max(0, row.netWeightQtl)),
        lotId: row.lotId,
        moisturePercent:
          row.netWeightQtl > 0 ? roundQtl(row.moistureWeightedTotal / row.netWeightQtl) : 0,
        mark: row.mark
      }));
  }

  function buildAdjustedLotFormationPreview(): ReportPreview {
    const rows = buildAdjustedLotFormationRows();
    const previewRows = rows.map((row) => ({
      "Sr No.": row.srNo,
      "Reg No.": row.regNo,
      "Farmer Name": row.farmerName,
      "Warehouse Name": row.warehouseName,
      "Stack Number": row.stackNumber,
      Bags: row.bags,
      "Adjusted Net Qty (QTL)": row.netWeightQtl,
      "Lot ID": row.lotId,
      "Moisture %": row.moisturePercent
    }));

    return {
      reportType: "ADJUSTED_LOT_FORMATION_REGISTER",
      title: "Adjusted Lot Formation Register",
      columns: Object.keys(
        previewRows[0] ?? {
          "Sr No.": 1,
          "Reg No.": "",
          "Farmer Name": "",
          "Warehouse Name": "",
          "Stack Number": "",
          Bags: 0,
          "Adjusted Net Qty (QTL)": 0,
          "Lot ID": "",
          "Moisture %": 0
        }
      ),
      rows: previewRows,
      totals: {
        "Total Rows": rows.length,
        "Adjusted Rows": rows.filter((row) => row.mark === "*").length,
        "Total Bags": rows.reduce((sum, row) => sum + row.bags, 0),
        "Total Adjusted Net Qty (QTL)": formatNumber(rows.reduce((sum, row) => sum + row.netWeightQtl, 0)),
        "Marker Note": "* means bags/weight changed due to adjusted stack accommodation."
      },
      generatedAt: new Date().toISOString(),
      fileName: `adjusted-lot-formation-${new Date().toISOString().slice(0, 10)}.xlsx`
    };
  }

  function buildAdjustedLotLedgerFarmerWisePreview(): ReportPreview {
    const rows = buildAdjustedLotFormationRows()
      .sort((left, right) => {
        const farmerCompare = left.farmerName.localeCompare(right.farmerName, "en", {
          sensitivity: "base"
        });
        if (farmerCompare !== 0) {
          return farmerCompare;
        }
        const regCompare = left.regNo.localeCompare(right.regNo, "en", {
          numeric: true,
          sensitivity: "base"
        });
        if (regCompare !== 0) {
          return regCompare;
        }
        const lotCompare = left.lotId.localeCompare(right.lotId, "en", {
          numeric: true,
          sensitivity: "base"
        });
        if (lotCompare !== 0) {
          return lotCompare;
        }
        const warehouseCompare = left.warehouseName.localeCompare(right.warehouseName, "en", {
          sensitivity: "base"
        });
        if (warehouseCompare !== 0) {
          return warehouseCompare;
        }
        return left.stackNumber.localeCompare(right.stackNumber, "en", {
          numeric: true,
          sensitivity: "base"
        });
      });

    const previewRows = rows.map((row, index) => ({
      "Sr No.": index + 1,
      "Reg No.": row.regNo,
      "Farmer Name": row.farmerName,
      "Payment Voucher No.": row.paymentVoucherNo,
      "Expected Yield (QTL)": row.expectedYieldQtl,
      "Total Net Intake (QTL)": row.totalNetIntakeQtl,
      "Lot ID": row.lotId,
      "Warehouse Name": row.warehouseName,
      "Stack Number": row.stackNumber,
      Bags: row.bags,
      "Adjusted Net Qty (QTL)": row.netWeightQtl,
      "Moisture %": row.moisturePercent
    }));

    return {
      reportType: "ADJUSTED_LOT_LEDGER_FARMER_WISE",
      title: "Adjusted Lot Ledger Farmer Wise",
      columns: Object.keys(
        previewRows[0] ?? {
          "Sr No.": 1,
          "Reg No.": "",
          "Farmer Name": "",
          "Payment Voucher No.": "",
          "Expected Yield (QTL)": 0,
          "Total Net Intake (QTL)": 0,
          "Lot ID": "",
          "Warehouse Name": "",
          "Stack Number": "",
          Bags: 0,
          "Adjusted Net Qty (QTL)": 0,
          "Moisture %": 0
        }
      ),
      rows: previewRows,
      totals: {
        "Total Farmers": new Set(rows.map((row) => row.regNo)).size,
        "Total Rows": rows.length,
        "Adjusted Rows": rows.filter((row) => row.mark === "*").length,
        "Total Expected Yield (QTL)": formatNumber(
          Array.from(new Map(rows.map((row) => [row.regNo, row.expectedYieldQtl] as const)).values()).reduce(
            (sum, value) => sum + value,
            0
          )
        ),
        "Total Farmer Net Intake (QTL)": formatNumber(
          Array.from(new Map(rows.map((row) => [row.regNo, row.totalNetIntakeQtl] as const)).values()).reduce(
            (sum, value) => sum + value,
            0
          )
        ),
        "Total Bags": rows.reduce((sum, row) => sum + row.bags, 0),
        "Total Adjusted Net Qty (QTL)": formatNumber(rows.reduce((sum, row) => sum + row.netWeightQtl, 0)),
        "Marker Note": "* means bags/weight changed due to adjusted stack accommodation."
      },
      generatedAt: new Date().toISOString(),
      fileName: `adjusted-lot-ledger-farmer-wise-${new Date().toISOString().slice(0, 10)}.xlsx`
    };
  }

  function applyVerticalMerges(
    worksheet: XLSX.WorkSheet,
    rowStart: number,
    dataRows: Record<string, string | number>[],
    columns: string[]
  ) {
    const mergeColumns = columns.filter((column) =>
      [
        "Reg No.",
        "Farmer Name",
        "Payment Voucher No.",
        "Expected Yield (QTL)",
        "Total Net Intake (QTL)",
        "Warehouse Name",
        "Stack Number",
        "Lot ID"
      ].includes(column)
    );
    const merges: NonNullable<XLSX.WorkSheet["!merges"]> = [...(worksheet["!merges"] ?? [])];
    for (const column of mergeColumns) {
      const columnIndex = columns.indexOf(column);
      let start = 0;
      while (start < dataRows.length) {
        let end = start;
        while (
          end + 1 < dataRows.length &&
          String(dataRows[end + 1]?.[column] ?? "") === String(dataRows[start]?.[column] ?? "")
        ) {
          end += 1;
        }
        if (end > start && String(dataRows[start]?.[column] ?? "").trim()) {
          merges.push({
            s: { r: rowStart + start, c: columnIndex },
            e: { r: rowStart + end, c: columnIndex }
          });
        }
        start = end + 1;
      }
    }
    worksheet["!merges"] = merges;
  }

  function buildMergedAdjustedLotPdfBody(preview: ReportPreview, shouldMerge = true) {
    if (!shouldMerge) {
      return preview.rows.map((row) => preview.columns.map((column) => String(row[column] ?? "")));
    }

    const mergeColumns = new Set([
      "Reg No.",
      "Farmer Name",
      "Payment Voucher No.",
      "Expected Yield (QTL)",
      "Total Net Intake (QTL)",
      "Warehouse Name",
      "Stack Number",
      "Lot ID"
    ]);
    const spans = new Map<string, { start: number; end: number }[]>();
    for (const column of preview.columns) {
      if (!mergeColumns.has(column)) {
        continue;
      }
      const columnSpans: { start: number; end: number }[] = [];
      let start = 0;
      while (start < preview.rows.length) {
        let end = start;
        while (
          end + 1 < preview.rows.length &&
          String(preview.rows[end + 1]?.[column] ?? "") === String(preview.rows[start]?.[column] ?? "")
        ) {
          end += 1;
        }
        if (end > start && String(preview.rows[start]?.[column] ?? "").trim()) {
          columnSpans.push({ start, end });
        }
        start = end + 1;
      }
      spans.set(column, columnSpans);
    }

    return preview.rows.map((row, rowIndex) =>
      preview.columns.map((column) => {
        const span = spans.get(column)?.find((item) => rowIndex >= item.start && rowIndex <= item.end);
        if (!span) {
          return String(row[column] ?? "");
        }
        if (rowIndex === span.start) {
          return {
            content: String(row[column] ?? ""),
            rowSpan: span.end - span.start + 1
          };
        }
        return "";
      })
    );
  }

  function polishAdjustedLotWorksheet(
    worksheet: XLSX.WorkSheet,
    preview: ReportPreview,
    widthOverrides: Record<string, number> = {}
  ) {
    worksheet["!cols"] = preview.columns.map((column) => ({
      wch:
        widthOverrides[column] ??
        (column.includes("Farmer") ? 28 : column.includes("Warehouse") ? 24 : column.includes("Adjusted") ? 18 : 14)
    }));
    worksheet["!autofilter"] = {
      ref: XLSX.utils.encode_range({
        s: { r: 5, c: 0 },
        e: { r: Math.max(5, 5 + preview.rows.length), c: Math.max(0, preview.columns.length - 1) }
      })
    };
    worksheet["!freeze"] = {
      xSplit: 0,
      ySplit: 6,
      topLeftCell: "A7",
      activePane: "bottomLeft",
      state: "frozen"
    } as never;
  }

  function buildAdjustedLotFormationWorkbook(preview: ReportPreview) {
    const rows: (string | number)[][] = [
      [COMPANY_NAME],
      [preview.title],
      [`Generated At: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`],
      ["Rule", "Rows marked with * are adjusted reporting rows. Original intake, lots, and vouchers are unchanged."],
      [],
      preview.columns,
      ...preview.rows.map((row) => preview.columns.map((column) => row[column] ?? ""))
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    applyVerticalMerges(worksheet, 6, preview.rows, preview.columns);
    polishAdjustedLotWorksheet(worksheet, preview, {
      "Sr No.": 8,
      "Reg No.": 14,
      "Stack Number": 12,
      Bags: 10,
      "Adjusted Net Qty (QTL)": 20,
      "Moisture %": 12
    });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Adjusted Lot Formation");
    return workbook;
  }

  function buildAdjustedLotFarmerWiseWorkbook(preview: ReportPreview) {
    const groupedRows: (string | number)[][] = [];
    const groupSpans: { start: number; end: number }[] = [];
    let currentRegNo = "";
    let currentStart = -1;

    preview.rows.forEach((row) => {
      const regNo = String(row["Reg No."] ?? "");
      if (currentRegNo && regNo !== currentRegNo) {
        groupSpans.push({ start: currentStart, end: groupedRows.length - 1 });
        groupedRows.push([]);
        currentStart = groupedRows.length;
      }
      if (!currentRegNo || regNo !== currentRegNo) {
        currentRegNo = regNo;
        currentStart = groupedRows.length;
      }
      groupedRows.push(preview.columns.map((column) => row[column] ?? ""));
    });
    if (currentRegNo) {
      groupSpans.push({ start: currentStart, end: groupedRows.length - 1 });
    }

    const rows: (string | number)[][] = [
      [COMPANY_NAME],
      [preview.title],
      [`Generated At: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`],
      [
        "Rule",
        "Farmer-wise adjusted lot ledger. Common farmer fields are merged per registration; blank row separates each farmer."
      ],
      [],
      preview.columns,
      ...groupedRows
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const commonColumns = [
      "Reg No.",
      "Farmer Name",
      "Payment Voucher No.",
      "Expected Yield (QTL)",
      "Total Net Intake (QTL)"
    ];
    const merges: NonNullable<XLSX.WorkSheet["!merges"]> = [...(worksheet["!merges"] ?? [])];
    for (const span of groupSpans) {
      if (span.end <= span.start) {
        continue;
      }
      for (const column of commonColumns) {
        const columnIndex = preview.columns.indexOf(column);
        if (columnIndex < 0) {
          continue;
        }
        merges.push({
          s: { r: 6 + span.start, c: columnIndex },
          e: { r: 6 + span.end, c: columnIndex }
        });
      }
    }
    worksheet["!merges"] = merges;
    polishAdjustedLotWorksheet(worksheet, preview, {
      "Sr No.": 8,
      "Reg No.": 14,
      "Payment Voucher No.": 22,
      "Expected Yield (QTL)": 18,
      "Total Net Intake (QTL)": 20,
      "Lot ID": 14,
      "Stack Number": 12,
      Bags: 10,
      "Adjusted Net Qty (QTL)": 20,
      "Moisture %": 12
    });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Farmer Wise Lot Ledger");
    return workbook;
  }


  function buildCustomDatePaymentRegisterPreview(): ReportPreview {
    const fromDate = reportFromDate.trim();
    const toDate = reportToDate.trim();
    const districtFilter = reportDistrict.trim().toLowerCase();
    const villageFilter = reportVillage.trim().toLowerCase();
    const farmerFilter = reportFarmerName.trim().toLowerCase();
    const statusFilterValue = reportPaymentStatus.trim().toLowerCase();
    const seasonLabel = reportSeasonLabel.trim() || "RABI 2025-26";

    const paymentRows = financialVouchers
      .flatMap((voucher) =>
        (voucher.payments ?? []).map((payment) => ({
          date: payment.paymentDate,
          voucherNo: voucher.voucherNo,
          farmerName: voucher.farmerName,
          regCode: voucher.cropRegistrationCode,
          village: voucher.village || "-",
          district: voucher.district || "-",
          status: voucher.status || "DRAFT",
          transactionNo: payment.transactionNo || "-",
          mode: payment.mode || "RTGS/NEFT",
          amount: Number(payment.amount ?? 0)
        }))
      )
      .filter((row) => !fromDate || row.date >= fromDate)
      .filter((row) => !toDate || row.date <= toDate)
      .filter((row) => !districtFilter || row.district.toLowerCase() === districtFilter)
      .filter((row) => !villageFilter || row.village.toLowerCase() === villageFilter)
      .filter((row) => !farmerFilter || row.farmerName.toLowerCase() === farmerFilter)
      .filter((row) => !statusFilterValue || row.status.toLowerCase() === statusFilterValue)
      .sort((left, right) => {
        const dateCompare = left.date.localeCompare(right.date);
        if (dateCompare !== 0) {
          return dateCompare;
        }
        const voucherCompare = left.voucherNo.localeCompare(right.voucherNo, "en", {
          sensitivity: "base",
          numeric: true
        });
        if (voucherCompare !== 0) {
          return voucherCompare;
        }
        return left.farmerName.localeCompare(right.farmerName, "en", {
          sensitivity: "base"
        });
      });

    const previewRows = paymentRows.map((row, index) => ({
      "S.No.": index + 1,
      Date: formatDateDisplay(row.date),
      "Voucher No.": row.voucherNo,
      "Farmer Name": row.farmerName,
      "Reg. Code": row.regCode,
      Village: row.village,
      District: row.district,
      Status: row.status,
      "Transaction No.": row.transactionNo,
      Mode: row.mode,
      Amount: roundQtl(row.amount)
    }));

    const districtSummary = Array.from(
      paymentRows.reduce((map, row) => {
        const current = map.get(row.district) ?? { count: 0, amount: 0 };
        current.count += 1;
        current.amount = roundQtl(current.amount + row.amount);
        map.set(row.district, current);
        return map;
      }, new Map<string, { count: number; amount: number }>())
    )
      .sort(([left], [right]) => left.localeCompare(right, "en", { sensitivity: "base" }))
      .map(([district, summary]) => `${district}: ${summary.count} / ${formatNumber(summary.amount)}`)
      .join(" | ") || "-";

    const statusSummary = Array.from(
      paymentRows.reduce((map, row) => {
        const current = map.get(row.status) ?? { count: 0, amount: 0 };
        current.count += 1;
        current.amount = roundQtl(current.amount + row.amount);
        map.set(row.status, current);
        return map;
      }, new Map<string, { count: number; amount: number }>())
    )
      .sort(([left], [right]) => left.localeCompare(right, "en", { sensitivity: "base" }))
      .map(([status, summary]) => `${status}: ${summary.count} / ${formatNumber(summary.amount)}`)
      .join(" | ") || "-";

    return {
      reportType: "CUSTOM_DATE_PAYMENT_REGISTER",
      title: "Custom Date Payment Register",
      columns: Object.keys(
        previewRows[0] ?? {
          "S.No.": 1,
          Date: "",
          "Voucher No.": "",
          "Farmer Name": "",
          "Reg. Code": "",
          Village: "",
          District: "",
          Status: "",
          "Transaction No.": "",
          Mode: "",
          Amount: 0
        }
      ),
      rows: previewRows,
      totals: {
        "From Date": fromDate ? formatDateDisplay(fromDate) : "ALL",
        "To Date": toDate ? formatDateDisplay(toDate) : "ALL",
        "Total Entries": previewRows.length,
        "Total Amount": formatNumber(paymentRows.reduce((sum, row) => sum + row.amount, 0)),
        "Unique Farmers": new Set(paymentRows.map((row) => row.regCode)).size,
        "District Summary": districtSummary,
        "Status Summary": statusSummary
      },
      generatedAt: new Date().toISOString(),
      fileName: `payment-register-${seasonLabel.replace(/\s+/g, "-").toLowerCase()}.xlsx`
    };
  }

  function buildOrganizerFarmerPaymentRegisterPreview(): ReportPreview {
    const fromDate = reportFromDate.trim();
    const toDate = reportToDate.trim();
    const districtFilter = reportDistrict.trim().toLowerCase();
    const organizerFilter = reportOrganizerName.trim().toLowerCase();
    const seasonLabel = reportSeasonLabel.trim() || "RABI 2025-26";

    const rows = financialVouchers
      .map((voucher) => {
        const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
        const organizerName = registration?.organizerName?.trim() || "Direct Farmer";
        return {
          organizerName,
          farmerName: voucher.farmerName,
          regCode: voucher.cropRegistrationCode,
          village: voucher.village || registration?.village || "-",
          district: voucher.district || registration?.district || "-",
          seasonLabel: `${voucher.season} ${voucher.year}`.trim(),
          voucherDate: voucher.voucherDate,
          netAmount: roundQtl(getVoucherFinalPayable(voucher)),
          paidAmount: roundQtl(getVoucherTotalPaid(voucher)),
          balanceAmount: roundQtl(getVoucherBalance(voucher)),
          status: voucher.status || "DRAFT"
        };
      })
      .filter((row) => !seasonLabel || row.seasonLabel.toLowerCase() === seasonLabel.toLowerCase())
      .filter((row) => !fromDate || row.voucherDate >= fromDate)
      .filter((row) => !toDate || row.voucherDate <= toDate)
      .filter((row) => !districtFilter || row.district.toLowerCase() === districtFilter)
      .filter((row) => !organizerFilter || row.organizerName.toLowerCase() === organizerFilter)
      .sort((left, right) => {
        const organizerCompare = left.organizerName.localeCompare(right.organizerName, "en", {
          sensitivity: "base"
        });
        if (organizerCompare !== 0) {
          return organizerCompare;
        }
        const farmerCompare = left.farmerName.localeCompare(right.farmerName, "en", {
          sensitivity: "base"
        });
        if (farmerCompare !== 0) {
          return farmerCompare;
        }
        return left.regCode.localeCompare(right.regCode, "en", {
          sensitivity: "base",
          numeric: true
        });
      });

    const previewRows = rows.map((row, index) => ({
      "S.No.": index + 1,
      Organizer: row.organizerName,
      "Farmer Name": row.farmerName,
      "Reg. Code": row.regCode,
      Village: row.village,
      District: row.district,
      "Net Amount": row.netAmount,
      "Paid Amount": row.paidAmount,
      Balance: row.balanceAmount,
      Status: row.status
    }));

    const organizerSummary = Array.from(
      rows.reduce((map, row) => {
        const current = map.get(row.organizerName) ?? { count: 0, balance: 0 };
        current.count += 1;
        current.balance = roundQtl(current.balance + row.balanceAmount);
        map.set(row.organizerName, current);
        return map;
      }, new Map<string, { count: number; balance: number }>())
    )
      .sort(([left], [right]) => left.localeCompare(right, "en", { sensitivity: "base" }))
      .map(([organizerName, summary]) => `${organizerName}: ${summary.count} / ${formatNumber(summary.balance)}`)
      .join(" | ") || "-";

    return {
      reportType: "ORGANIZER_FARMER_PAYMENT_REGISTER",
      title: "Organizer Wise Farmer Payment Register",
      columns: Object.keys(
        previewRows[0] ?? {
          "S.No.": 1,
          Organizer: "",
          "Farmer Name": "",
          "Reg. Code": "",
          Village: "",
          District: "",
          "Net Amount": 0,
          "Paid Amount": 0,
          Balance: 0,
          Status: ""
        }
      ),
      rows: previewRows,
      totals: {
        "From Date": fromDate ? formatDateDisplay(fromDate) : "ALL",
        "To Date": toDate ? formatDateDisplay(toDate) : "ALL",
        Organizer: reportOrganizerName || "ALL",
        "Total Farmers": previewRows.length,
        "Total Net Amount": formatNumber(rows.reduce((sum, row) => sum + row.netAmount, 0)),
        "Total Paid Amount": formatNumber(rows.reduce((sum, row) => sum + row.paidAmount, 0)),
        "Total Balance": formatNumber(rows.reduce((sum, row) => sum + row.balanceAmount, 0)),
        "Organizer Summary": organizerSummary
      },
      generatedAt: new Date().toISOString(),
      fileName: `organizer-farmer-payment-${seasonLabel.replace(/\s+/g, "-").toLowerCase()}.xlsx`
    };
  }

  function buildOrganizerPaymentTransactionReportPreview(): ReportPreview {
    const fromDate = reportFromDate.trim();
    const toDate = reportToDate.trim();
    const districtFilter = reportDistrict.trim().toLowerCase();
    const organizerFilter = reportOrganizerName.trim().toLowerCase();
    const seasonLabel = reportSeasonLabel.trim() || "RABI 2025-26";

    const blocks: OrganizerPaymentTransactionBlock[] = financialVouchers
      .map((voucher) => {
        const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
        const organizerName = registration?.organizerName?.trim() || "Direct Farmer";
        const voucherSeasonLabel = `${voucher.season} ${voucher.year}`.trim();
        const payments = (voucher.payments ?? [])
          .filter((payment) => !fromDate || payment.paymentDate >= fromDate)
          .filter((payment) => !toDate || payment.paymentDate <= toDate)
          .sort((left, right) => {
            const dateCompare = left.paymentDate.localeCompare(right.paymentDate);
            if (dateCompare !== 0) {
              return dateCompare;
            }
            return (left.transactionNo || "").localeCompare(right.transactionNo || "", "en", {
              sensitivity: "base",
              numeric: true
            });
          })
          .map((payment) => ({
            transactionDate: payment.paymentDate,
            payment: roundQtl(Number(payment.amount ?? 0)),
            transactionNumber: payment.transactionNo || "-",
            transactionRemark: payment.remarks || "-"
          }));

        return {
          organizerName,
          regCode: voucher.cropRegistrationCode,
          name: voucher.farmerName,
          village: voucher.village || registration?.village || "-",
          district: voucher.district || registration?.district || "-",
          seasonLabel: voucherSeasonLabel,
          totalBags: Number(voucher.totalBags ?? 0),
          totalNetQtyQtl: roundQtl(Number(voucher.totalNetQtyQtl ?? 0)),
          ratePerQtl: roundQtl(Number(voucher.certifiedRatePerQtl ?? 0)),
          grossAmount: roundQtl(Number(voucher.grossPayableAmount ?? 0)),
          deduction: roundQtl(Number(voucher.deductionAmount ?? 0)),
          finalAmount: roundQtl(getVoucherFinalPayable(voucher)),
          payments
        };
      })
      .filter((block) => block.payments.length > 0)
      .filter((block) => !seasonLabel || block.seasonLabel.toLowerCase() === seasonLabel.toLowerCase())
      .filter((block) => !districtFilter || block.district.toLowerCase() === districtFilter)
      .filter((block) => !organizerFilter || block.organizerName.toLowerCase() === organizerFilter)
      .sort((left, right) => {
        const organizerCompare = left.organizerName.localeCompare(right.organizerName, "en", {
          sensitivity: "base"
        });
        if (organizerCompare !== 0) {
          return organizerCompare;
        }
        const nameCompare = left.name.localeCompare(right.name, "en", { sensitivity: "base" });
        if (nameCompare !== 0) {
          return nameCompare;
        }
        return left.regCode.localeCompare(right.regCode, "en", {
          sensitivity: "base",
          numeric: true
        });
      });

    const previewRows = blocks.flatMap((block) => {
      const paymentRows = block.payments.map((payment, paymentIndex) => ({
        "reg code": paymentIndex === 0 ? block.regCode : "",
        name: paymentIndex === 0 ? block.name : "",
        village: paymentIndex === 0 ? block.village : "",
        distt: paymentIndex === 0 ? block.district : "",
        "total bags": paymentIndex === 0 ? block.totalBags : "",
        "total net qty": paymentIndex === 0 ? block.totalNetQtyQtl : "",
        rate: paymentIndex === 0 ? block.ratePerQtl : "",
        "gross amount": paymentIndex === 0 ? block.grossAmount : "",
        deduction: paymentIndex === 0 ? block.deduction : "",
        "final amount": paymentIndex === 0 ? block.finalAmount : "",
        "transaction date": payment.transactionDate,
        payment: payment.payment,
        "transaction number": payment.transactionNumber,
        "transaction remark": payment.transactionRemark
      }));
      const paidAmount = roundQtl(
        block.payments.reduce((sum, payment) => sum + payment.payment, 0)
      );
      const balanceAmount = roundQtl(block.finalAmount - paidAmount);
      return [
        ...paymentRows,
        {
          "reg code": "",
          name: "",
          village: "",
          distt: "",
          "total bags": "",
          "total net qty": "",
          rate: "",
          "gross amount": "",
          deduction: "",
          "final amount": "Farmer total / balance",
          "transaction date": "",
          payment: paidAmount,
          "transaction number": `Balance: ${formatNumber(balanceAmount)}`,
          "transaction remark":
            balanceAmount < 0 ? "OVERPAID" : balanceAmount === 0 ? "SETTLED" : "PENDING"
        }
      ];
    });

    const organizerSummary =
      Array.from(
        blocks.reduce((map, block) => {
          const current = map.get(block.organizerName) ?? { count: 0, payment: 0 };
          current.count += block.payments.length;
          current.payment = roundQtl(
            current.payment + block.payments.reduce((sum, payment) => sum + payment.payment, 0)
          );
          map.set(block.organizerName, current);
          return map;
        }, new Map<string, { count: number; payment: number }>())
      )
        .sort(([left], [right]) => left.localeCompare(right, "en", { sensitivity: "base" }))
        .map(([organizerName, summary]) => `${organizerName}: ${summary.count} / ${formatNumber(summary.payment)}`)
        .join(" | ") || "-";

    const totalPayment = roundQtl(
      blocks.reduce(
        (sum, block) =>
          sum + block.payments.reduce((paymentSum, payment) => paymentSum + payment.payment, 0),
        0
      )
    );

    return {
      reportType: "ORGANIZER_PAYMENT_TRANSACTION_REPORT",
      title: "Organizer Payment Transaction Report",
      columns: Object.keys(
        previewRows[0] ?? {
          "reg code": "",
          name: "",
          village: "",
          distt: "",
          "total bags": 0,
          "total net qty": 0,
          rate: 0,
          "gross amount": 0,
          deduction: 0,
          "final amount": 0,
          "transaction date": "",
          payment: 0,
          "transaction number": "",
          "transaction remark": ""
        }
      ),
      rows: previewRows,
      totals: {
        "From Date": fromDate ? formatDateDisplay(fromDate) : "ALL",
        "To Date": toDate ? formatDateDisplay(toDate) : "ALL",
        Organizer: reportOrganizerName || "ALL",
        "Total Farmers": blocks.length,
        "Total Payment Entries": blocks.reduce((sum, block) => sum + block.payments.length, 0),
        "Total Bags": blocks.reduce((sum, block) => sum + block.totalBags, 0),
        "Total Net Qty": formatNumber(blocks.reduce((sum, block) => sum + block.totalNetQtyQtl, 0)),
        "Total Gross Amount": formatNumber(blocks.reduce((sum, block) => sum + block.grossAmount, 0)),
        "Total Deduction": formatNumber(blocks.reduce((sum, block) => sum + block.deduction, 0)),
        "Total Final Amount": formatNumber(blocks.reduce((sum, block) => sum + block.finalAmount, 0)),
        "Total Payment": formatNumber(totalPayment),
        "Total Balance": formatNumber(blocks.reduce((sum, block) => sum + block.finalAmount, 0) - totalPayment),
        "Organizer Summary": organizerSummary
      },
      generatedAt: new Date().toISOString(),
      fileName: `organizer-payment-transactions-${seasonLabel.replace(/\s+/g, "-").toLowerCase()}.xlsx`
    };
  }

  function buildOverpaidFarmerReportPreview(): ReportPreview {
    const fromDate = reportFromDate.trim();
    const toDate = reportToDate.trim();
    const districtFilter = reportDistrict.trim().toLowerCase();
    const farmerFilter = reportFarmerName.trim().toLowerCase();
    const registrationFilter = reportRegistrationCode.trim().toLowerCase();
    const seasonLabel = reportSeasonLabel.trim() || "RABI 2025-26";

    const rows = financialVouchers
      .map((voucher) => {
        const registration = registrations.find((item) => item.id === voucher.cropRegistrationId);
        const organizerName = registration?.organizerName?.trim() || "Direct Farmer";
        const finalPayable = roundQtl(getVoucherFinalPayable(voucher));
        const paidAmount = roundQtl(getVoucherTotalPaid(voucher));
        const balanceAmount = roundQtl(getVoucherBalance(voucher));
        return {
          voucherNo: voucher.voucherNo,
          voucherDate: voucher.voucherDate,
          organizerName,
          farmerName: voucher.farmerName,
          regCode: voucher.cropRegistrationCode,
          village: voucher.village || registration?.village || "-",
          district: voucher.district || registration?.district || "-",
          grossPayableAmount: roundQtl(Number(voucher.grossPayableAmount ?? 0)),
          deductionAmount: roundQtl(Number(voucher.deductionAmount ?? 0)),
          finalPayable,
          paidAmount,
          balanceAmount,
          status: voucher.status || "DRAFT",
          seasonEntry: `${voucher.season} ${voucher.year}`.trim()
        };
      })
      .filter((row) => row.balanceAmount < 0 || row.status === "OVERPAID")
      .filter((row) => !seasonLabel || row.seasonEntry.toLowerCase() === seasonLabel.toLowerCase())
      .filter((row) => !fromDate || row.voucherDate >= fromDate)
      .filter((row) => !toDate || row.voucherDate <= toDate)
      .filter((row) => !districtFilter || row.district.toLowerCase() === districtFilter)
      .filter((row) => !farmerFilter || row.farmerName.toLowerCase().includes(farmerFilter))
      .filter((row) => !registrationFilter || row.regCode.toLowerCase().includes(registrationFilter))
      .sort((left, right) => {
        const balanceCompare = left.balanceAmount - right.balanceAmount;
        if (balanceCompare !== 0) {
          return balanceCompare;
        }
        return left.farmerName.localeCompare(right.farmerName, "en", { sensitivity: "base" });
      });

    const previewRows = rows.map((row, index) => ({
      "S.No.": index + 1,
      "Voucher No.": row.voucherNo,
      Date: formatDateDisplay(row.voucherDate),
      Organizer: row.organizerName,
      "Farmer Name": row.farmerName,
      "Reg. Code": row.regCode,
      Village: row.village,
      District: row.district,
      "Gross Payable": row.grossPayableAmount,
      Deduction: row.deductionAmount,
      "Final Payable": row.finalPayable,
      "Paid Amount": row.paidAmount,
      Balance: row.balanceAmount,
      Status: row.status
    }));

    const organizerSummary = Array.from(
      rows.reduce((map, row) => {
        const current = map.get(row.organizerName) ?? { count: 0, amount: 0 };
        current.count += 1;
        current.amount = roundQtl(current.amount + Math.abs(row.balanceAmount));
        map.set(row.organizerName, current);
        return map;
      }, new Map<string, { count: number; amount: number }>())
    )
      .sort(([left], [right]) => left.localeCompare(right, "en", { sensitivity: "base" }))
      .map(([organizerName, summary]) => `${organizerName}: ${summary.count} / ${formatNumber(summary.amount)}`)
      .join(" | ") || "-";

    return {
      reportType: "OVERPAID_FARMER_REPORT",
      title: "Overpaid Farmer Report",
      columns: Object.keys(
        previewRows[0] ?? {
          "S.No.": 1,
          "Voucher No.": "",
          Date: "",
          Organizer: "",
          "Farmer Name": "",
          "Reg. Code": "",
          Village: "",
          District: "",
          "Gross Payable": 0,
          Deduction: 0,
          "Final Payable": 0,
          "Paid Amount": 0,
          Balance: 0,
          Status: ""
        }
      ),
      rows: previewRows,
      totals: {
        Season: seasonLabel,
        "Total Overpaid Farmers": previewRows.length,
        "Overpaid Exposure": formatNumber(rows.reduce((sum, row) => sum + Math.abs(row.balanceAmount), 0)),
        "Most Negative Balance": formatNumber(Math.abs(Math.min(...rows.map((row) => row.balanceAmount), 0))),
        "Organizer Summary": organizerSummary
      },
      generatedAt: new Date().toISOString(),
      fileName: `overpaid-farmer-report-${seasonLabel.replace(/\s+/g, "-").toLowerCase()}.xlsx`
    };
  }

  function buildReceiptVoucherTraceabilityPreview(): ReportPreview {
    const fromDate = reportFromDate.trim();
    const toDate = reportToDate.trim();
    const districtFilter = reportDistrict.trim().toLowerCase();
    const farmerFilter = reportFarmerName.trim().toLowerCase();
    const registrationFilter = reportRegistrationCode.trim().toLowerCase();
    const seasonLabel = reportSeasonLabel.trim() || "RABI 2025-26";

    const rows = receipts
      .map((receipt) => {
        const registration = registrations.find((item) => item.id === receipt.cropRegistrationId);
        const voucher = financialVouchers.find((item) => item.cropRegistrationId === receipt.cropRegistrationId) ?? null;
        return {
          receiptNo: receipt.receiptNo,
          receiptDate: receipt.receiptDate,
          regCode: receipt.cropRegistrationCode,
          farmerName: receipt.farmerName,
          village: registration?.village || "-",
          district: registration?.district || "-",
          netIntakeQtl: roundQtl(
            receipt.lines.reduce((sum, line) => sum + Number(line.netWeightQtl ?? line.qtyQtl ?? 0), 0)
          ),
          voucherNo: voucher?.voucherNo || "-",
          voucherStatus: voucher?.status || "NOT GENERATED",
          finalPayable: voucher ? roundQtl(getVoucherFinalPayable(voucher)) : 0,
          paidAmount: voucher ? roundQtl(getVoucherTotalPaid(voucher)) : 0,
          balanceAmount: voucher ? roundQtl(getVoucherBalance(voucher)) : 0,
          seasonEntry: registration ? `${registration.season} ${registration.year}` : ""
        };
      })
      .filter((row) => !seasonLabel || row.seasonEntry.toLowerCase() === seasonLabel.toLowerCase())
      .filter((row) => !fromDate || row.receiptDate >= fromDate)
      .filter((row) => !toDate || row.receiptDate <= toDate)
      .filter((row) => !districtFilter || row.district.toLowerCase() === districtFilter)
      .filter((row) => !farmerFilter || row.farmerName.toLowerCase().includes(farmerFilter))
      .filter((row) => !registrationFilter || row.regCode.toLowerCase().includes(registrationFilter))
      .sort((left, right) => {
        const dateCompare = left.receiptDate.localeCompare(right.receiptDate);
        if (dateCompare !== 0) {
          return dateCompare;
        }
        const receiptCompare = left.receiptNo.localeCompare(right.receiptNo, "en", {
          sensitivity: "base",
          numeric: true
        });
        if (receiptCompare !== 0) {
          return receiptCompare;
        }
        return left.regCode.localeCompare(right.regCode, "en", { sensitivity: "base", numeric: true });
      });

    const previewRows = rows.map((row, index) => ({
      "S.No.": index + 1,
      "Receipt No.": row.receiptNo,
      "Receipt Date": formatDateDisplay(row.receiptDate),
      "Reg. Code": row.regCode,
      "Farmer Name": row.farmerName,
      Village: row.village,
      District: row.district,
      "Net Intake (QTL)": row.netIntakeQtl,
      "Voucher No.": row.voucherNo,
      "Voucher Status": row.voucherStatus,
      "Final Payable": row.finalPayable,
      "Paid Amount": row.paidAmount,
      Balance: row.balanceAmount
    }));

    return {
      reportType: "RECEIPT_VOUCHER_TRACEABILITY_REPORT",
      title: "Receipt to Voucher Traceability Report",
      columns: Object.keys(
        previewRows[0] ?? {
          "S.No.": 1,
          "Receipt No.": "",
          "Receipt Date": "",
          "Reg. Code": "",
          "Farmer Name": "",
          Village: "",
          District: "",
          "Net Intake (QTL)": 0,
          "Voucher No.": "",
          "Voucher Status": "",
          "Final Payable": 0,
          "Paid Amount": 0,
          Balance: 0
        }
      ),
      rows: previewRows,
      totals: {
        Season: seasonLabel,
        "Total Receipts": previewRows.length,
        "Receipts Linked to Voucher": rows.filter((row) => row.voucherNo !== "-").length,
        "Receipts Without Voucher": rows.filter((row) => row.voucherNo === "-").length,
        "Total Net Intake (QTL)": formatNumber(rows.reduce((sum, row) => sum + row.netIntakeQtl, 0))
      },
      generatedAt: new Date().toISOString(),
      fileName: `receipt-voucher-traceability-${seasonLabel.replace(/\s+/g, "-").toLowerCase()}.xlsx`
    };
  }

  function buildOrganizerIntakePaymentCommissionPreview(): ReportPreview {
    const fromDate = reportFromDate.trim();
    const toDate = reportToDate.trim();
    const districtFilter = reportDistrict.trim().toLowerCase();
    const organizerFilter = reportOrganizerName.trim().toLowerCase();
    const seasonLabel = reportSeasonLabel.trim() || "RABI 2025-26";

    const rows = organizers
      .filter((organizer) => !organizerFilter || organizer.name.toLowerCase() === organizerFilter)
      .filter((organizer) => !districtFilter || organizer.district.toLowerCase() === districtFilter)
      .map((organizer) => {
        const linkedRegistrations = registrations.filter(
          (item) =>
            item.organizerId === organizer.id &&
            `${item.season} ${item.year}`.trim().toLowerCase() === seasonLabel.toLowerCase() &&
            (!districtFilter || item.district.toLowerCase() === districtFilter)
        );
        const registrationIds = new Set(linkedRegistrations.map((item) => item.id));
        const linkedVouchers = financialVouchers.filter(
          (item) =>
            registrationIds.has(item.cropRegistrationId) &&
            (!fromDate || item.voucherDate >= fromDate) &&
            (!toDate || item.voucherDate <= toDate)
        );
        const organizerPaymentRows = organizerPayments.filter(
          (item) =>
            item.organizerId === organizer.id &&
            (!fromDate || item.paymentDate >= fromDate) &&
            (!toDate || item.paymentDate <= toDate)
        );
        const expectedYieldQtl = roundQtl(
          linkedRegistrations.reduce((sum, item) => sum + Number(item.expectedYieldQtl ?? 0), 0)
        );
        const receivedQtyQtl = roundQtl(
          linkedRegistrations.reduce((sum, item) => sum + Number(item.totalReceivedQtl ?? 0), 0)
        );
        const farmerNetPayable = roundQtl(
          linkedVouchers.reduce((sum, item) => sum + getVoucherFinalPayable(item), 0)
        );
        const farmerPaid = roundQtl(linkedVouchers.reduce((sum, item) => sum + getVoucherTotalPaid(item), 0));
        const farmerBalance = roundQtl(linkedVouchers.reduce((sum, item) => sum + getVoucherBalance(item), 0));
        const commissionRatePerQtl = roundQtl(Number(organizer.commissionRatePerQtl ?? 0));
        const grossCommission = roundQtl(receivedQtyQtl * commissionRatePerQtl);
        const deductionAmount = roundQtl(Number(organizer.deductionAmount ?? 0));
        const netCommissionPayable = roundQtl(grossCommission - deductionAmount);
        const organizerPaid = roundQtl(
          organizerPaymentRows.reduce((sum, item) => sum + Number(item.amount ?? 0), 0)
        );
        const organizerBalance = roundQtl(netCommissionPayable - organizerPaid);
        return {
          organizerName: organizer.name,
          district: organizer.district || "-",
          farmerCount: linkedRegistrations.length,
          expectedYieldQtl,
          receivedQtyQtl,
          farmerNetPayable,
          farmerPaid,
          farmerBalance,
          commissionRatePerQtl,
          grossCommission,
          deductionAmount,
          netCommissionPayable,
          organizerPaid,
          organizerBalance
        };
      })
      .filter((row) => row.farmerCount > 0)
      .sort((left, right) => left.organizerName.localeCompare(right.organizerName, "en", { sensitivity: "base" }));

    const previewRows = rows.map((row, index) => ({
      "S.No.": index + 1,
      Organizer: row.organizerName,
      District: row.district,
      Farmers: row.farmerCount,
      "Expected Yield": row.expectedYieldQtl,
      "Received Qty": row.receivedQtyQtl,
      "Farmer Net Payable": row.farmerNetPayable,
      "Farmer Paid": row.farmerPaid,
      "Farmer Balance": row.farmerBalance,
      "Commission Rate/QTL": row.commissionRatePerQtl,
      "Gross Commission": row.grossCommission,
      Deduction: row.deductionAmount,
      "Net Commission Payable": row.netCommissionPayable,
      "Organizer Paid": row.organizerPaid,
      "Organizer Balance": row.organizerBalance
    }));

    return {
      reportType: "ORGANIZER_INTAKE_PAYMENT_COMMISSION_REPORT",
      title: "Organizer Intake vs Payment vs Commission",
      columns: Object.keys(
        previewRows[0] ?? {
          "S.No.": 1,
          Organizer: "",
          District: "",
          Farmers: 0,
          "Expected Yield": 0,
          "Received Qty": 0,
          "Farmer Net Payable": 0,
          "Farmer Paid": 0,
          "Farmer Balance": 0,
          "Commission Rate/QTL": 0,
          "Gross Commission": 0,
          Deduction: 0,
          "Net Commission Payable": 0,
          "Organizer Paid": 0,
          "Organizer Balance": 0
        }
      ),
      rows: previewRows,
      totals: {
        Season: seasonLabel,
        "Total Organizers": previewRows.length,
        "Expected Yield": formatNumber(rows.reduce((sum, row) => sum + row.expectedYieldQtl, 0)),
        "Received Qty": formatNumber(rows.reduce((sum, row) => sum + row.receivedQtyQtl, 0)),
        "Farmer Net Payable": formatNumber(rows.reduce((sum, row) => sum + row.farmerNetPayable, 0)),
        "Farmer Paid": formatNumber(rows.reduce((sum, row) => sum + row.farmerPaid, 0)),
        "Gross Commission": formatNumber(rows.reduce((sum, row) => sum + row.grossCommission, 0)),
        "Organizer Balance": formatNumber(rows.reduce((sum, row) => sum + row.organizerBalance, 0))
      },
      generatedAt: new Date().toISOString(),
      fileName: `organizer-intake-payment-commission-${seasonLabel.replace(/\s+/g, "-").toLowerCase()}.xlsx`
    };
  }

  function buildLocalReportWorkbook(preview: ReportPreview) {
    const rows: (string | number)[][] = [];
    rows.push([preview.title]);
    rows.push([`Generated At: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`]);
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
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, preview.title.slice(0, 31) || "Report");
    return workbook;
  }

  function buildOrganizerPaymentTransactionWorkbook(preview: ReportPreview) {
    const rows: (string | number)[][] = [];
    rows.push([COMPANY_NAME]);
    rows.push([preview.title]);
    rows.push([`Generated At: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`]);
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
    const merges: XLSX.Range[] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(preview.columns.length - 1, 0) } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(preview.columns.length - 1, 0) } }
    ];

    let activeBlockStart = -1;
    preview.rows.forEach((row, index) => {
      const sheetRow = index + 5;
      if (String(row["reg code"] ?? "").trim()) {
        activeBlockStart = sheetRow;
      }
      if (row["final amount"] === "Farmer total / balance" && activeBlockStart >= 0) {
        const lastPaymentRow = sheetRow - 1;
        if (lastPaymentRow > activeBlockStart) {
          for (let columnIndex = 0; columnIndex <= 9; columnIndex += 1) {
            merges.push({
              s: { r: activeBlockStart, c: columnIndex },
              e: { r: lastPaymentRow, c: columnIndex }
            });
          }
        }
        activeBlockStart = -1;
      }
    });
    worksheet["!merges"] = merges;
    worksheet["!cols"] = [
      { wch: 14 },
      { wch: 26 },
      { wch: 20 },
      { wch: 12 },
      { wch: 11 },
      { wch: 14 },
      { wch: 12 },
      { wch: 15 },
      { wch: 13 },
      { wch: 19 },
      { wch: 16 },
      { wch: 14 },
      { wch: 30 },
      { wch: 28 }
    ];
    worksheet["!autofilter"] = {
      ref: `A5:${XLSX.utils.encode_col(Math.max(preview.columns.length - 1, 0))}${Math.max(
        5,
        preview.rows.length + 5
      )}`
    };

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Payment Transactions");
    return workbook;
  }

  function runDatabaseBackup() {
    if (!requirePermission("canMaintenance", "Only Admin can run database backup.")) {
      return;
    }
    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/database/backup`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          backupDirectory
        })
      });
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to take database backup.");
      }
      const data = (await response.json()) as { backupDirectory: string; message: string };
      setMaintenanceMessage(`Backup saved to: ${data.backupDirectory}`);
      notifyUser(data.message);
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to take database backup.");
    });
  }

  function runDatabaseRestore() {
    if (!requirePermission("canMaintenance", "Only Admin can run database restore.")) {
      return;
    }
    if (!restoreDirectory.trim()) {
      notifyUser("Enter restore folder or directory first.");
      return;
    }
    void (async () => {
      const response = await fetchWithAuth(`${API_BASE}/api/seed/database/restore`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          restoreDirectory
        })
      });
      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to restore database.");
      }
      const data = (await response.json()) as { restoreDirectory: string; message: string };
      setMaintenanceMessage(`Restore executed from: ${data.restoreDirectory}`);
      await loadBootstrap();
      notifyUser(data.message);
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to restore database.");
    });
  }

  async function fetchReportPreviewData() {
    if (reportType === "CUSTOM_DATE_PAYMENT_REGISTER") {
      return buildCustomDatePaymentRegisterPreview();
    }
    if (reportType === "ORGANIZER_FARMER_PAYMENT_REGISTER") {
      return buildOrganizerFarmerPaymentRegisterPreview();
    }
    if (reportType === "ORGANIZER_PAYMENT_TRANSACTION_REPORT") {
      return buildOrganizerPaymentTransactionReportPreview();
    }
    if (reportType === "OVERPAID_FARMER_REPORT") {
      return buildOverpaidFarmerReportPreview();
    }
    if (reportType === "RECEIPT_VOUCHER_TRACEABILITY_REPORT") {
      return buildReceiptVoucherTraceabilityPreview();
    }
    if (reportType === "ORGANIZER_INTAKE_PAYMENT_COMMISSION_REPORT") {
      return buildOrganizerIntakePaymentCommissionPreview();
    }
    if (reportType === "ADJUSTED_LOT_FORMATION_REGISTER") {
      return buildAdjustedLotFormationPreview();
    }
    if (reportType === "ADJUSTED_LOT_LEDGER_FARMER_WISE") {
      return buildAdjustedLotLedgerFarmerWisePreview();
    }
    if (reportType === "STACK_CARD_REGISTER") {
      return buildStackCardRegisterPreview();
    }

    const response = await fetchWithAuth(`${API_BASE}/api/seed/reports/preview`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(buildReportRequest())
    });

    if (!response.ok) {
      throw new Error("Unable to preview report.");
    }

    return (await response.json()) as ReportPreview;
  }

  function previewReportModule() {
    void (async () => {
      const data = await fetchReportPreviewData();
      setReportPreview(data);
      notifyUser(`${data.title} preview loaded.`);
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to preview report.");
    });
  }

  function downloadReportWorkbook() {
    void (async () => {
      if (
        reportType === "CUSTOM_DATE_PAYMENT_REGISTER" ||
        reportType === "ORGANIZER_FARMER_PAYMENT_REGISTER" ||
        reportType === "ORGANIZER_PAYMENT_TRANSACTION_REPORT" ||
        reportType === "OVERPAID_FARMER_REPORT" ||
        reportType === "RECEIPT_VOUCHER_TRACEABILITY_REPORT" ||
        reportType === "ORGANIZER_INTAKE_PAYMENT_COMMISSION_REPORT" ||
        reportType === "ADJUSTED_LOT_FORMATION_REGISTER" ||
        reportType === "ADJUSTED_LOT_LEDGER_FARMER_WISE" ||
        reportType === "STACK_CARD_REGISTER"
      ) {
        const preview =
          reportPreview ??
          (reportType === "CUSTOM_DATE_PAYMENT_REGISTER"
            ? buildCustomDatePaymentRegisterPreview()
            : reportType === "ORGANIZER_FARMER_PAYMENT_REGISTER"
              ? buildOrganizerFarmerPaymentRegisterPreview()
              : reportType === "ORGANIZER_PAYMENT_TRANSACTION_REPORT"
                ? buildOrganizerPaymentTransactionReportPreview()
                : reportType === "OVERPAID_FARMER_REPORT"
                  ? buildOverpaidFarmerReportPreview()
                  : reportType === "RECEIPT_VOUCHER_TRACEABILITY_REPORT"
                    ? buildReceiptVoucherTraceabilityPreview()
                    : reportType === "ORGANIZER_INTAKE_PAYMENT_COMMISSION_REPORT"
                      ? buildOrganizerIntakePaymentCommissionPreview()
                      : reportType === "ADJUSTED_LOT_FORMATION_REGISTER"
                        ? buildAdjustedLotFormationPreview()
                        : reportType === "ADJUSTED_LOT_LEDGER_FARMER_WISE"
                          ? buildAdjustedLotLedgerFarmerWisePreview()
                        : buildStackCardRegisterPreview());
        if (!reportPreview) {
          setReportPreview(preview);
        }
        const workbook =
          reportType === "STACK_CARD_REGISTER"
            ? buildStackCardRegisterWorkbook(buildStackCardRegisterSections(), preview)
            : reportType === "ADJUSTED_LOT_FORMATION_REGISTER"
              ? buildAdjustedLotFormationWorkbook(preview)
              : reportType === "ADJUSTED_LOT_LEDGER_FARMER_WISE"
                ? buildAdjustedLotFarmerWiseWorkbook(preview)
            : reportType === "ORGANIZER_PAYMENT_TRANSACTION_REPORT"
              ? buildOrganizerPaymentTransactionWorkbook(preview)
            : buildLocalReportWorkbook(preview);
        XLSX.writeFile(workbook, preview.fileName || "payment-register.xlsx");
        notifyUser("Report workbook downloaded.");
        return;
      }

      const response = await fetchWithAuth(`${API_BASE}/api/seed/reports/export`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(buildReportRequest())
      });

      if (!response.ok) {
        throw new Error("Unable to export report workbook.");
      }

      const blob = await response.blob();
      const header = response.headers.get("Content-Disposition") ?? "";
      const matchedName = header.match(/filename="?([^"]+)"?/i)?.[1];
      const fileName = matchedName || "report.xlsx";
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = fileName;
      document.body.append(anchor);
      anchor.click();
      anchor.remove();
      window.URL.revokeObjectURL(url);
      notifyUser("Report workbook downloaded.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to export report workbook.");
    });
  }

  function downloadReportPdf() {
    void (async () => {
      const preview = reportPreview ?? (await fetchReportPreviewData());
      if (!reportPreview) {
        setReportPreview(preview);
      }

      if (reportType === "ADJUSTED_LOT_FORMATION_REGISTER" || reportType === "ADJUSTED_LOT_LEDGER_FARMER_WISE") {
        const pdf = new jsPDF({
          orientation: "landscape",
          unit: "mm",
          format: "legal"
        });
        const pageWidth = pdf.internal.pageSize.getWidth();
        let reportY = renderPdfBrandHeader(pdf, preview.title, {
          left: 14,
          right: pageWidth - 14,
          y: 10,
          compact: false,
          logoDataUrl: pdfLogoDataUrl
        });

        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(8);
        pdf.text(`Season: ${reportSeasonLabel}`, 14, reportY);
        pdf.text(`Generated: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`, 120, reportY);
        reportY += 5;
        pdf.text(
          "Rule: Adjusted Net Qty is the reporting quantity after stack/lot adjustment. Original intake, lot, and voucher data is unchanged.",
          14,
          reportY
        );
        reportY += 4;
        pdf.text("* means bags/weight changed due to adjusted stack accommodation.", 14, reportY);

        const adjustedLotColumnStyles = preview.columns.reduce(
          (styles, column, index) => {
            if (column === "Sr No.") styles[index] = { halign: "center", cellWidth: 10 };
            if (column === "Reg No.") styles[index] = { cellWidth: 22 };
            if (column === "Farmer Name") styles[index] = { cellWidth: 42 };
            if (column === "Payment Voucher No.") styles[index] = { cellWidth: 28 };
            if (column === "Expected Yield (QTL)") styles[index] = { halign: "right", cellWidth: 22 };
            if (column === "Total Net Intake (QTL)") styles[index] = { halign: "right", cellWidth: 24 };
            if (column === "Lot ID") styles[index] = { cellWidth: 22 };
            if (column === "Warehouse Name") styles[index] = { cellWidth: 36 };
            if (column === "Stack Number") styles[index] = { halign: "center", cellWidth: 16 };
            if (column === "Bags") styles[index] = { halign: "right", cellWidth: 14 };
            if (column === "Adjusted Net Qty (QTL)") styles[index] = { halign: "right", cellWidth: 24 };
            if (column === "Moisture %") styles[index] = { halign: "right", cellWidth: 16 };
            return styles;
          },
          {} as Record<number, { halign?: "left" | "center" | "right"; cellWidth?: number }>
        );

        autoTable(pdf, {
          startY: reportY + 5,
          head: [preview.columns],
          body: buildMergedAdjustedLotPdfBody(
            preview,
            reportType !== "ADJUSTED_LOT_LEDGER_FARMER_WISE"
          ) as never,
          theme: "grid",
          styles: {
            font: "helvetica",
            fontSize: 6.8,
            cellPadding: 1.2,
            overflow: "linebreak",
            textColor: [47, 30, 18],
            valign: "middle"
          },
          headStyles: {
            fillColor: [48, 86, 61],
            textColor: [255, 255, 255],
            fontStyle: "bold",
            fontSize: 6.5,
            halign: "center",
            valign: "middle"
          },
          alternateRowStyles: {
            fillColor: [252, 248, 241]
          },
          columnStyles: adjustedLotColumnStyles,
          margin: { left: 14, right: 14, bottom: 10 }
        });

        pdf.save((preview.fileName || "adjusted-lot-formation.xlsx").replace(/\.xlsx$/i, ".pdf"));
        notifyUser("Report PDF downloaded.");
        return;
      }

      if (reportType === "STACK_CARD_REGISTER") {
        const sections = buildStackCardRegisterSections();
        const pdf = new jsPDF({
          orientation: "landscape",
          unit: "mm",
          format: "legal"
        });
        const pageWidth = pdf.internal.pageSize.getWidth();
        let currentY = renderPdfBrandHeader(pdf, preview.title, {
          left: 14,
          right: pageWidth - 14,
          y: 10,
          compact: false,
          logoDataUrl: pdfLogoDataUrl
        });

        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(8);
        pdf.text(`Generated: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`, 14, currentY);
        currentY += 5;
        pdf.text("Rule: Non-adjusted stacks show original sequence. Adjusted stacks show final position with * mark.", 14, currentY);

        currentY += 7;
        sections.forEach((section, index) => {
          if (currentY > 180) {
            pdf.addPage();
            currentY = 18;
          }

          pdf.setFont("helvetica", "bold");
          pdf.setFontSize(10);
          pdf.text(`${index + 1}. ${section.godownName} - Stack ${section.stackNo}`, 14, currentY);
          pdf.setFont("helvetica", "normal");
          pdf.setFontSize(8);
          pdf.text(
            `View: ${section.displayMode === "FINAL_ADJUSTED" ? "Final Adjusted Stack Position" : "Original Stack Position"}`,
            14,
            currentY + 5
          );

          autoTable(pdf, {
            startY: currentY + 8,
            head: [["S. No.", "Reg. Code", "Farmer Name", "Changed (*)", "Village", "District", "Qty (QTL)", "Bags"]],
            body: [
              ...section.rows.map((row, rowIndex) => [
                String(rowIndex + 1),
                row.regCode,
                row.mark === "*" ? `* ${row.farmerName}` : row.farmerName,
                row.mark,
                row.village,
                row.district,
                formatNumber(row.qtyQtl),
                String(row.bags)
              ]),
              ["", "", "", "", "", "TOTAL", formatNumber(section.totalQtyQtl), String(section.totalBags)]
            ],
            theme: "grid",
            styles: {
              font: "helvetica",
              fontSize: 6.8,
              cellPadding: 1.3,
              overflow: "linebreak",
              textColor: [47, 30, 18]
            },
            headStyles: {
              fillColor: section.displayMode === "FINAL_ADJUSTED" ? [48, 86, 61] : [127, 75, 38],
              fontSize: 7,
              halign: "center"
            },
            bodyStyles: {
              valign: "middle"
            },
            margin: { left: 14, right: 14, bottom: 10 }
          });

          const sectionTable = pdf as jsPDF & { lastAutoTable?: { finalY?: number } };
          currentY = (sectionTable.lastAutoTable?.finalY ?? currentY + 8) + 6;
          if (section.displayMode === "FINAL_ADJUSTED" && section.changedFarmerCount > 0) {
            pdf.setFontSize(7.5);
            pdf.text("* means this farmer row changed because of stack adjustment.", 14, currentY);
            currentY += 5;
          }
        });

        pdf.save((preview.fileName || "stack-card-register.xlsx").replace(/\.xlsx$/i, ".pdf"));
        notifyUser("Report PDF downloaded.");
        return;
      }

      const pdf = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "legal"
      });
      const pageWidth = pdf.internal.pageSize.getWidth();
      let reportY = renderPdfBrandHeader(pdf, preview.title, {
        left: 14,
        right: pageWidth - 14,
        y: 10,
        compact: false,
        logoDataUrl: pdfLogoDataUrl
      });

      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(8);
      pdf.text(`Season: ${reportSeasonLabel}`, 14, reportY);
      pdf.text(`Generated: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`, 120, reportY);
      reportY += 4;

      autoTable(pdf, {
        startY: reportY,
        head: [["Metric", "Value"]],
        body: Object.entries(preview.totals).map(([label, value]) => [label, String(value)]),
        theme: "grid",
        styles: {
          font: "helvetica",
          fontSize: 7.5,
          cellPadding: 1.6,
          textColor: [47, 30, 18]
        },
        headStyles: {
          fillColor: [127, 75, 38],
          fontSize: 7.5
        },
        tableWidth: 88,
        margin: { left: 14, right: 14 }
      });

      const metricsTable = pdf as jsPDF & { lastAutoTable?: { finalY?: number } };
      autoTable(pdf, {
        startY: (metricsTable.lastAutoTable?.finalY ?? reportY) + 4,
        head: [preview.columns],
        body: preview.rows.map((row) => preview.columns.map((column) => String(row[column] ?? ""))),
        theme: "grid",
        styles: {
          font: "helvetica",
          fontSize: 6.8,
          cellPadding: 1.3,
          overflow: "linebreak",
          textColor: [47, 30, 18]
        },
        headStyles: {
          fillColor: [127, 75, 38],
          fontSize: 7,
          halign: "center"
        },
        bodyStyles: {
          valign: "middle"
        },
        margin: { left: 14, right: 14, bottom: 10 }
      });

      pdf.save((preview.fileName || "report.xlsx").replace(/\.xlsx$/i, ".pdf"));
      notifyUser("Report PDF downloaded.");
    })().catch((error) => {
      notifyUser(error instanceof Error ? error.message : "Unable to download report PDF.");
    });
  }

  function resetReportFilters() {
    setReportType("GODOWN_WISE_DETAIL");
    setReportSeasonLabel("RABI 2025-26");
    setReportFromDate("");
    setReportToDate("");
    setReportCrop("");
    setReportVariety("");
    setReportClassStage("");
    setReportDistrict("");
    setReportGodownId("");
    setReportStackNo("");
    setReportRegistrationCode("");
    setReportFarmerName("");
    setReportOrganizerName("");
    setReportVillage("");
    setReportPaymentStatus("");
    setReportMode("ALL");
    setReportPreview(null);
  }

  function buildSlipPreview(
    overrides?: Partial<{
      slipType: SlipType;
      slipRegistrationId: string;
      slipReceiptNo: string;
      slipDate: string;
    }>
  ): SlipPreview | null {
    const nextSlipType = overrides?.slipType ?? slipType;
    const nextRegistrationId = overrides?.slipRegistrationId ?? slipRegistrationId;
    const nextReceiptNo = overrides?.slipReceiptNo ?? slipReceiptNo;
    const nextSlipDate = overrides?.slipDate ?? slipDate;
    const nextRegistration = registrations.find((item) => item.id === nextRegistrationId) ?? null;

    if (!nextRegistration) {
      throw new Error("Select a registration first.");
    }

    const nextSlipReceipts = receipts
      .filter((item) => item.cropRegistrationId === nextRegistration.id)
      .slice()
      .sort((left, right) =>
        `${left.receiptDate}-${left.receiptNo}`.localeCompare(`${right.receiptDate}-${right.receiptNo}`)
      );
    const nextSlipReceipt = nextSlipReceipts.find((item) => item.receiptNo === nextReceiptNo) ?? null;
    const nextSlipDailyReceipts = nextSlipDate
      ? nextSlipReceipts.filter((item) => item.receiptDate === nextSlipDate)
      : [];
    const nextSlipDiscrepancies = discrepancies.filter(
      (item) => item.cropRegistrationId === nextRegistration.id
    );
    const nextSlipShifts = discrepancyShifts.filter(
      (item) => item.cropRegistrationCode === nextRegistration.cropRegistrationCode
    );
    const nextSlipLots = lotLedgerRows.filter(
      (item) => item.cropRegistrationId === nextRegistration.id && item.status !== "VOID"
    );
    const totalNetQtl = nextSlipReceipts.reduce((sum, receipt) => sum + sumReceiptNetQty(receipt), 0);
    const totalGrossQtl = nextSlipReceipts.reduce(
      (sum, receipt) =>
        sum + receipt.lines.reduce((lineSum, line) => lineSum + Number(line.grossWeightQtl ?? 0), 0),
      0
    );
    const totalBags = nextSlipReceipts.reduce(
      (sum, receipt) =>
        sum + receipt.lines.reduce((lineSum, line) => lineSum + Number(line.noOfBags ?? 0), 0),
      0
    );

    if (nextSlipType === "FARMER_SINGLE_RECEIPT") {
      if (!nextSlipReceipt) {
        throw new Error("Select a receipt for farmer intake slip.");
      }

      const receiptGross = nextSlipReceipt.lines.reduce(
        (sum, line) => sum + Number(line.grossWeightQtl ?? 0),
        0
      );
      const receiptNet = sumReceiptNetQty(nextSlipReceipt);
      const receiptBags = nextSlipReceipt.lines.reduce(
        (sum, line) => sum + Number(line.noOfBags ?? 0),
        0
      );
      const tableRows = nextSlipReceipt.lines.map((line, index) => ({
        No: index + 1,
        "Rcpt No": nextSlipReceipt.receiptNo,
        Vehicle: line.vehicleNo || "-",
        Stack: line.stackNo,
        Bags: Number(line.noOfBags ?? 0),
        Gross: roundQtl(Number(line.grossWeightQtl ?? 0)),
        Net: roundQtl(Number(line.netWeightQtl ?? line.qtyQtl ?? 0))
      }));

      return {
        slipType: nextSlipType,
        title: "FARMER INTAKE RECEIPT",
        pageSize: "A5",
        orientation: "portrait",
        template: "FARMER_SINGLE_CLASSIC",
        slipNo: `FIS-${nextSlipReceipt.receiptDate}-${nextSlipReceipt.receiptNo}`,
        summary: [
          { label: "Date", value: formatDateDisplay(nextSlipReceipt.receiptDate) },
          { label: "Season", value: `${nextRegistration.season} ${nextRegistration.year}` },
          { label: "Receipt No.", value: nextSlipReceipt.receiptNo },
          { label: "Reg. Code", value: nextRegistration.cropRegistrationCode },
          { label: "Farmer Name", value: nextRegistration.farmerName },
          { label: "F/H Name", value: nextRegistration.fatherName || "-" },
          { label: "Village", value: nextRegistration.village || "-" },
          { label: "Block", value: nextRegistration.block || "-" },
          { label: "District", value: nextRegistration.district || "-" },
          { label: "Crop", value: nextRegistration.crop },
          { label: "Variety", value: nextRegistration.variety },
          { label: "Class", value: nextRegistration.classStage },
          { label: "Exp. Yield", value: `${formatNumber(nextRegistration.expectedYieldQtl)} QTL` }
        ],
        tableColumns: Object.keys(
          tableRows[0] ?? { No: 1, "Rcpt No": "", Vehicle: "", Stack: "", Bags: 0, Gross: 0, Net: 0 }
        ),
        tableRows,
        totals: [
          { label: "Prev. Received", value: `${formatNumber(Math.max(roundQtl(nextRegistration.totalReceivedQtl - receiptNet), 0))} QTL` },
          { label: "Today Received", value: `${formatNumber(receiptNet)} QTL` },
          { label: "Balance", value: `${formatNumber(Math.max(roundQtl(nextRegistration.allowedIntakeQtl - nextRegistration.totalReceivedQtl), 0))} QTL` },
          { label: "Total Bags", value: String(receiptBags) },
          { label: "Total Gross", value: `${formatNumber(receiptGross)} QTL` },
          { label: "Total Net", value: `${formatNumber(receiptNet)} QTL` }
        ],
        footerNote:
          nextSlipReceipt.lines.map((line) => line.remarks?.trim()).filter(Boolean).join(" | ") ||
          "Farmer intake receipt issued."
      };
    }

    if (nextSlipType === "FARMER_OVERALL") {
      if (nextSlipReceipts.length === 0) {
        throw new Error("No intake found for the selected registration.");
      }

      const totalDiscrepancyQtl = roundQtl(
        nextSlipDiscrepancies.reduce((sum, item) => sum + item.excessQtyQtl, 0)
      );
      const totalDiscrepancyBags = nextSlipDiscrepancies.reduce(
        (sum, item) => sum + Number(item.estimatedExcessBags ?? 0),
        0
      );
      const totalShiftedQtl = roundQtl(
        nextSlipShifts.reduce((sum, item) => sum + item.shiftedQtyQtl, 0)
      );
      const pendingDiscrepancyQtl = Math.max(roundQtl(totalDiscrepancyQtl - totalShiftedQtl), 0);
      const discrepancyStatus =
        nextSlipDiscrepancies.length === 0
          ? "CLEAR"
          : pendingDiscrepancyQtl > 0
            ? "UNDER REVIEW"
            : "RESOLVED";

      const tableRows = nextSlipReceipts.map((receipt, index) => ({
        No: index + 1,
        Date: receipt.receiptDate,
        "Rcpt No": receipt.receiptNo,
        Vehicle:
          Array.from(
            new Set(receipt.lines.map((line) => line.vehicleNo?.trim()).filter(Boolean))
        ).join(", ") || "-",
        Bags: receipt.lines.reduce((sum, line) => sum + Number(line.noOfBags ?? 0), 0),
        Gross: roundQtl(
          receipt.lines.reduce((sum, line) => sum + Number(line.grossWeightQtl ?? 0), 0)
        ),
        Net: roundQtl(sumReceiptNetQty(receipt))
      }));
      const lotSummaryLines = nextSlipLots.length
        ? nextSlipLots.map((lot) => ({
            label: `L${lot.lotNo}`,
            value: `Stack ${lot.stackNo}   ${lot.bags} Bags   ${formatNumber(roundQtl(lot.displayQtyQtl))} QTL`
          }))
        : [{ label: "L-", value: "No lots created" }];
      const discrepancyLines = [
        { label: "Excess Qty", value: `${formatNumber(totalDiscrepancyQtl)} QTL` },
        { label: "Excess Bags", value: String(totalDiscrepancyBags) },
        { label: "Status", value: discrepancyStatus }
      ];
      if (totalShiftedQtl > 0) {
        discrepancyLines.push({
          label: "Shifted Qty",
          value: `${formatNumber(totalShiftedQtl)} QTL`
        });
      }
      const finalStatusLines = [
        { label: "Certification Status", value: discrepancyStatus }
      ];
      if (totalDiscrepancyQtl > 0) {
        finalStatusLines.push({
          label: "Excess Qty",
          value: `${formatNumber(totalDiscrepancyQtl)} QTL`
        });
      }
      const extraSections = [
        { title: "Lot Summary", lines: lotSummaryLines }
      ];
      if (totalDiscrepancyQtl > 0) {
        extraSections.push({ title: "Discrepancy", lines: discrepancyLines });
      }
      extraSections.push({ title: "Final Status", lines: finalStatusLines });

      return {
        slipType: nextSlipType,
        title: "FARMER OVERALL CONSOLIDATED INTAKE SLIP",
        pageSize: "A5",
        orientation: "portrait",
        template: "FARMER_OVERALL_CLASSIC",
        slipNo: `FOIS-${nextRegistration.cropRegistrationCode}`,
        summary: [
          { label: "Date", value: formatDateDisplay(new Date().toISOString().slice(0, 10)) },
          { label: "Season", value: `${nextRegistration.season} ${nextRegistration.year}` },
          { label: "Statement No.", value: `FOIS-${nextRegistration.cropRegistrationCode}` },
          { label: "Reg. Code", value: nextRegistration.cropRegistrationCode },
          { label: "Farmer Name", value: nextRegistration.farmerName },
          { label: "F/H Name", value: nextRegistration.fatherName || "-" },
          { label: "Village", value: nextRegistration.village || "-" },
          { label: "Block", value: nextRegistration.block || "-" },
          { label: "District", value: nextRegistration.district || "-" },
          { label: "Crop", value: nextRegistration.crop },
          { label: "Variety", value: nextRegistration.variety },
          { label: "Class", value: nextRegistration.classStage },
          { label: "Exp. Yield", value: `${formatNumber(nextRegistration.expectedYieldQtl)} QTL` }
        ],
        tableColumns: Object.keys(
          tableRows[0] ?? { No: 1, Date: "", "Rcpt No": "", Vehicle: "", Bags: 0, Gross: 0, Net: 0 }
        ),
        tableRows,
        totals: [
          { label: "Total Net", value: `${formatNumber(totalNetQtl)} QTL` },
          { label: "Total Bags", value: String(totalBags) },
          { label: "Total Gross", value: `${formatNumber(totalGrossQtl)} QTL` },
          { label: "Balance", value: `${formatNumber(Math.max(roundQtl(nextRegistration.allowedIntakeQtl - totalNetQtl), 0))} QTL` }
        ],
        extraSections,
        footerNote: `Overall farmer intake consolidated across ${nextSlipReceipts.length} receipt(s).`
      };
    }

    if (!nextSlipDate) {
      throw new Error("Select a date for daily consolidated intake slip.");
    }

    if (nextSlipDailyReceipts.length === 0) {
      throw new Error("No intake found for the selected registration and date.");
    }

    const previousReceivedQtl = roundQtl(
      nextSlipReceipts
        .filter((receipt) => receipt.receiptDate < nextSlipDate)
        .reduce((sum, receipt) => sum + sumReceiptNetQty(receipt), 0)
    );
    const dayGrossQtl = roundQtl(
      nextSlipDailyReceipts.reduce(
        (sum, receipt) =>
          sum + receipt.lines.reduce((lineSum, line) => lineSum + Number(line.grossWeightQtl ?? 0), 0),
        0
      )
    );
    const dayNetQtl = roundQtl(
      nextSlipDailyReceipts.reduce((sum, receipt) => sum + sumReceiptNetQty(receipt), 0)
    );
    const dayBags = nextSlipDailyReceipts.reduce(
      (sum, receipt) =>
        sum + receipt.lines.reduce((lineSum, line) => lineSum + Number(line.noOfBags ?? 0), 0),
      0
    );
      const tableRows = nextSlipDailyReceipts.flatMap((receipt) =>
        receipt.lines.map((line, index) => ({
          No: `${receipt.receiptNo}-${index + 1}`,
          "Receipt No.": receipt.receiptNo,
          Vehicle: line.vehicleNo || "-",
          Stack: line.stackNo,
          Bags: Number(line.noOfBags ?? 0),
          Gross: roundQtl(Number(line.grossWeightQtl ?? 0)),
          Net: roundQtl(Number(line.netWeightQtl ?? line.qtyQtl ?? 0))
        }))
      );
    const remarks = Array.from(
      new Set(
        nextSlipDailyReceipts.flatMap((receipt) =>
          receipt.lines.map((line) => (line.remarks?.trim() ? line.remarks.trim() : ""))
        )
      )
    )
      .filter(Boolean)
      .join(" | ");
    const nextSlipNo = `DCIS-${nextSlipDate}-${String(nextSlipDailyReceipts.length).padStart(3, "0")}`;

    return {
      slipType: nextSlipType,
      title: "Daily Consolidated Intake Slip",
      pageSize: "A4",
      orientation: "portrait",
      template: "DAILY_CONSOLIDATED_CLASSIC",
      slipNo: nextSlipNo,
      summary: [
        { label: "Date", value: formatDateDisplay(nextSlipDate) },
        { label: "Season", value: `${nextRegistration.season} ${nextRegistration.year}` },
        { label: "Reg. Code", value: nextRegistration.cropRegistrationCode },
        { label: "Farmer Name", value: nextRegistration.farmerName },
        { label: "F/H Name", value: nextRegistration.fatherName || "-" },
        { label: "Village", value: nextRegistration.village || "-" },
        { label: "Block", value: nextRegistration.block || "-" },
        { label: "District", value: nextRegistration.district || "-" },
        { label: "Crop", value: nextRegistration.crop },
        { label: "Variety", value: nextRegistration.variety },
        { label: "Class", value: nextRegistration.classStage },
        { label: "Exp. Yield", value: `${formatNumber(nextRegistration.expectedYieldQtl)} QTL` }
      ],
      tableColumns: Object.keys(
        tableRows[0] ?? { No: 1, "Receipt No.": "", Vehicle: "", Stack: "", Bags: 0, Gross: 0, Net: 0 }
      ),
      tableRows,
      totals: [
        { label: "Prev. Received", value: `${formatNumber(previousReceivedQtl)} QTL` },
        { label: "Today Received", value: `${formatNumber(dayNetQtl)} QTL` },
        {
          label: "Balance",
          value: `${formatNumber(Math.max(roundQtl(nextRegistration.allowedIntakeQtl - previousReceivedQtl - dayNetQtl), 0))} QTL`
        },
        { label: "Total Bags", value: String(dayBags) },
        { label: "Total Gross", value: `${formatNumber(dayGrossQtl)} QTL` },
        { label: "Total Net", value: `${formatNumber(dayNetQtl)} QTL` }
      ],
      footerNote: remarks || "Daily consolidated intake issued."
    };
  }

  function renderSlipPreviewPdf(doc: jsPDF, preview: SlipPreview, addNewPage = false) {
    if (addNewPage) {
      doc.addPage();
    }

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const left = 10;
    const right = pageWidth - 10;
    let y = 10;
    const titleText =
      preview.template === "DAILY_CONSOLIDATED_CLASSIC"
        ? "DAILY CONSOLIDATED INTAKE SLIP"
        : preview.template === "FARMER_OVERALL_CLASSIC"
          ? "FARMER OVERALL CONSOLIDATED INTAKE SLIP"
          : "FARMER INTAKE RECEIPT";
    const totalMap = new Map(preview.totals.map((item) => [item.label, item.value]));
    const metaEntries = preview.summary.slice(0, 4);
    const detailEntries = preview.summary.slice(4);

    y = renderPdfBrandHeader(doc, titleText, {
      left,
      right,
      y,
      compact: preview.pageSize === "A5",
      logoDataUrl: pdfLogoDataUrl
    });

    doc.setFontSize(preview.pageSize === "A5" ? 9 : 10);
    metaEntries.forEach((item, index) => {
      const label = index === 0 ? "Slip No." : item.label;
      const value = index === 0 ? preview.slipNo ?? item.value : item.value;
      doc.setFont("helvetica", "bold");
      doc.text(`${label}`, left, y);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${value}`, left + 26, y);
      y += 4.5;
    });

    y += 1;
    const detailColGap = 8;
    const detailColWidth = (right - left - detailColGap) / 2;
    detailEntries.forEach((item, index) => {
      const col = index % 2;
      const row = Math.floor(index / 2);
      const rowY = y + row * 4.5;
      const x = left + col * (detailColWidth + detailColGap);
      doc.setFont("helvetica", "bold");
      doc.text(item.label, x, rowY);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${item.value}`, x + 23, rowY, { maxWidth: detailColWidth - 23 });
    });
    y += Math.ceil(detailEntries.length / 2) * 4.5 + 2;

    doc.line(left, y, right, y);
    y += 4;
    doc.setFont("helvetica", "bold");
    doc.text("Receipt Detail", left, y);
    y += 3;
    doc.line(left, y, right, y);
    y += 1;

    const body = preview.tableRows.map((row) =>
      preview.tableColumns.map((column) => String(row[column] ?? ""))
    );
    body.push([
      "Total",
      ...preview.tableColumns.slice(1, Math.max(preview.tableColumns.length - 3, 1)).map(() => ""),
      totalMap.get("Total Bags") ?? "",
      totalMap.get("Total Gross") ?? "",
      totalMap.get("Total Net") ?? ""
    ]);

    autoTable(doc, {
      startY: y,
      head: [preview.tableColumns],
      body,
      theme: "grid",
      styles: {
        fontSize: preview.pageSize === "A5" ? 7.5 : 8,
        cellPadding: 1.5,
        textColor: [47, 30, 18],
        lineColor: [216, 200, 183]
      },
      headStyles: {
        fillColor: [243, 237, 229],
        textColor: [47, 30, 18],
        fontStyle: "bold"
      },
      bodyStyles: {
        valign: "middle"
      },
      margin: { left, right: pageWidth - right },
      didParseCell: (data) => {
        if (data.row.index === body.length - 1) {
          data.cell.styles.fontStyle = "bold";
        }
      }
    });

    y = (doc as jsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY
      ? ((doc as jsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? y) + 5
      : y + 30;

    if (y > pageHeight - 35) {
      doc.addPage();
      y = 12;
    }

    const bottomLines = [
      totalMap.get("Prev. Received")
        ? { label: "Prev. Received", value: totalMap.get("Prev. Received") ?? "" }
        : null,
      totalMap.get("Today Received")
        ? { label: "Today Received", value: totalMap.get("Today Received") ?? "" }
        : null,
      totalMap.get("Total Net")
        ? { label: "Net Received", value: totalMap.get("Total Net") ?? "" }
        : null,
      totalMap.get("Balance")
        ? { label: "Balance", value: totalMap.get("Balance") ?? "" }
        : null
    ].filter((item): item is { label: string; value: string } => Boolean(item));

    bottomLines.forEach((line) => {
      doc.setFont("helvetica", "bold");
      doc.text(line.label, left, y);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${line.value}`, left + 26, y);
      y += 4.5;
    });
    y += 1.5;

    doc.setFont("helvetica", "bold");
    doc.text("Remarks:", left, y);
    y += 4;
    doc.setFont("helvetica", "normal");
    const extraSections = preview.extraSections ?? [];
    if (extraSections.length > 0) {
      extraSections.forEach((section) => {
        if (y > pageHeight - 28) {
          doc.addPage();
          y = 12;
        }
        doc.setFont("helvetica", "bold");
        doc.text(section.title, left, y);
        y += 4;
        section.lines.forEach((line) => {
          doc.setFont("helvetica", "bold");
          doc.text(line.label, left, y);
          doc.setFont("helvetica", "normal");
          const valueLines = doc.splitTextToSize(`: ${line.value}`, right - left - 24);
          doc.text(valueLines, left + 24, y);
          y += valueLines.length * 4;
        });
        y += 3;
      });
    }

    if (preview.footerNote.trim()) {
      if (y > pageHeight - 24) {
        doc.addPage();
        y = 12;
      }
      doc.setFont("helvetica", "bold");
      doc.text("Remarks:", left, y);
      y += 4;
      doc.setFont("helvetica", "normal");
      const remarksLines = doc.splitTextToSize(preview.footerNote, right - left);
      doc.text(remarksLines, left, y);
      y += remarksLines.length * 4 + 4;
    }

    if (y > pageHeight - 18) {
      doc.addPage();
      y = 12;
    }

    const signatureWidth = (right - left) / 3;
    doc.setFont("helvetica", "normal");
    doc.text("Operator Sign : ______________", left, y);
    doc.text("Farmer Sign : ______________", left + signatureWidth, y);
    doc.text("Godown Incharge : ______________", left + signatureWidth * 2, y);
  }

  function previewSlip(overrides?: Partial<{
    slipType: SlipType;
    slipRegistrationId: string;
    slipReceiptNo: string;
    slipDate: string;
  }>) {
    try {
      const preview = buildSlipPreview(overrides);
      setSlipPreview(preview);
      notifyUser(`${preview?.title ?? "Slip"} preview loaded.`);
    } catch (error) {
      notifyUser(error instanceof Error ? error.message : "Unable to build slip preview.");
    }
  }

  function openSlipModal(nextSlipType: SlipType, registrationId: string) {
    const nextRegistration = registrations.find((item) => item.id === registrationId) ?? null;
    const nextReceipts = receipts
      .filter((item) => item.cropRegistrationId === registrationId)
      .slice()
      .sort((left, right) =>
        `${left.receiptDate}-${left.receiptNo}`.localeCompare(`${right.receiptDate}-${right.receiptNo}`)
      );
    const defaultReceiptNo = nextReceipts[0]?.receiptNo ?? "";
    const defaultDate = Array.from(new Set(nextReceipts.map((item) => item.receiptDate)))[0] ?? "";

    setSlipType(nextSlipType);
    setSlipRegistrationId(registrationId);
    setSlipReceiptNo(defaultReceiptNo);
    setSlipDate(defaultDate);
    setSlipModalOpen(true);

    if (nextRegistration) {
      try {
        const preview = buildSlipPreview({
          slipType: nextSlipType,
          slipRegistrationId: registrationId,
          slipReceiptNo: defaultReceiptNo,
          slipDate: defaultDate
        });
        setSlipPreview(preview);
      } catch {
        setSlipPreview(null);
      }
    }
  }

  function closeSlipModal() {
    setSlipModalOpen(false);
    setSlipPreview(null);
  }

  function resetSlipFilters() {
    setSlipSearch("");
    setSlipDistrictFilter("");
    setSlipVillageFilter("");
    setSlipClassFilter("");
    setSlipCropFilter("");
    setSlipOnlyWithIntake(true);
    setSlipType("FARMER_SINGLE_RECEIPT");
    setSlipRegistrationId(registrations[0]?.id ?? "");
    setSlipReceiptNo("");
    setSlipDate(receipts[0]?.receiptDate ?? "");
    setSlipPreview(null);
    setSlipModalOpen(false);
  }

  function promptAdminPassword(voucherNo: string, actionLabel: string) {
    if (typeof window === "undefined") {
      return null;
    }
    const value = window.prompt(
      `${actionLabel} for paid voucher ${voucherNo} requires admin password.`
    );
    if (value === null) {
      return null;
    }
    return value.trim();
  }

  function openVoucherModal(registrationId: string, adminPassword = "") {
    if (!requirePermission("canVoucher", "Your role cannot generate or edit financial vouchers.")) {
      return;
    }
    const existingVoucher = voucherByRegistrationId.get(registrationId) ?? null;
    let overridePassword = adminPassword;
    if (existingVoucher && isVoucherLockedStatus(existingVoucher.status) && !overridePassword) {
      const prompted = promptAdminPassword(existingVoucher.voucherNo, "Editing");
      if (prompted === null) {
        return;
      }
      overridePassword = prompted;
    }
    setVoucherRegistrationId(registrationId);
    setVoucherDate(existingVoucher?.voucherDate ?? new Date().toISOString().slice(0, 10));
    setCertifiedRate(existingVoucher ? String(existingVoucher.certifiedRatePerQtl) : "");
    setDiscrepancyRate(existingVoucher ? String(existingVoucher.discrepancyRatePerQtl) : "");
    setVoucherDeduction(existingVoucher ? String(existingVoucher.deductionAmount) : "0");
    setVoucherAdminPassword(overridePassword);
    setVoucherRemarks(existingVoucher?.remarks ?? "");
    setVoucherPreview(
      existingVoucher
        ? {
            voucher: existingVoucher,
            hasDiscrepancy: Number(existingVoucher.discrepancyQtyQtl ?? 0) > 0
          }
        : null
    );
    setVoucherModalOpen(true);
  }

  function closeVoucherModal() {
    setVoucherModalOpen(false);
    setVoucherPreview(null);
    setVoucherAdminPassword("");
  }

  function openPaymentLedger(voucher: FinancialVoucher) {
    if (!requirePermission("canVoucher", "Your role cannot manage voucher payments.")) {
      return;
    }
    setPaymentLedgerVoucherId(voucher.id);
    setPaymentDate(new Date().toISOString().slice(0, 10));
    setPaymentAmount("");
    setPaymentTransactionNo("");
    setPaymentRemarks("");
    setEditingVoucherPaymentId("");
    setPaymentAdminPassword("");
  }

  function closePaymentLedger() {
    setPaymentLedgerVoucherId("");
    setPaymentDate(new Date().toISOString().slice(0, 10));
    setPaymentAmount("");
    setPaymentTransactionNo("");
    setPaymentRemarks("");
    setEditingVoucherPaymentId("");
    setPaymentAdminPassword("");
  }

  function startEditVoucherPayment(voucher: FinancialVoucher, payment: FinancialVoucherPayment) {
    if (!requirePermission("canVoucher", "Your role cannot manage voucher payments.")) {
      return;
    }
    setPaymentLedgerVoucherId(voucher.id);
    setEditingVoucherPaymentId(payment.id);
    setPaymentDate(payment.paymentDate);
    setPaymentAmount(String(payment.amount ?? ""));
    setPaymentTransactionNo(payment.transactionNo ?? "");
    setPaymentRemarks(payment.remarks ?? "");
    setPaymentAdminPassword("");
  }

  async function generateVoucher() {
    if (!requirePermission("canVoucher", "Your role cannot generate or edit financial vouchers.")) {
      return;
    }
    if (!voucherRegistrationId) {
      notifyUser("Select a registration first.");
      return;
    }
    const existingVoucher = voucherByRegistrationId.get(voucherRegistrationId) ?? null;

    const certifiedRateValue = Number(certifiedRate);
    const discrepancyRateValue = Number(discrepancyRate || 0);
    const deductionAmount = Number(voucherDeduction || 0);
    if (!Number.isFinite(certifiedRateValue) || certifiedRateValue < 0) {
      notifyUser("Enter certified rate.");
      return;
    }

    const response = await fetchWithAuth(
      existingVoucher
        ? `${API_BASE}/api/seed/financial-vouchers/${existingVoucher.id}`
        : `${API_BASE}/api/seed/financial-vouchers`,
      {
      method: existingVoucher ? "PUT" : "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        voucherDate,
        cropRegistrationId: voucherRegistrationId,
        certifiedRatePerQtl: certifiedRateValue,
        discrepancyRatePerQtl: discrepancyRateValue,
        deductionAmount,
        adminPassword: voucherAdminPassword,
        remarks: voucherRemarks
      })
    });

    if (!response.ok) {
      const errorBody = await response.json().catch(() => null);
      throw new Error(errorBody?.message || "Unable to generate financial voucher.");
    }

    const data = (await response.json()) as AppBootstrapPayload;

    setRegistrations(data.registrations ?? []);
    setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
    setStacks(data.stacks?.length ? data.stacks : defaultStacks);
    setLots(data.lots ?? []);
    setDiscrepancies(data.discrepancies ?? []);
    setDiscrepancyShifts(data.discrepancyShifts ?? []);
    setStackAccommodations(data.stackAccommodations ?? []);
    setFinancialVouchers(data.financialVouchers ?? []);
    setReceipts(data.receipts ?? []);
    setFeatures({
      discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
    });

    const createdVoucher =
      (data.financialVouchers ?? []).find((item) => item.cropRegistrationId === voucherRegistrationId) ??
      (data.financialVouchers ?? [])[0];

    if (!createdVoucher) {
      throw new Error("Voucher was generated but could not be loaded.");
    }

    setVoucherPreview({
      voucher: createdVoucher,
      hasDiscrepancy: Number(createdVoucher.discrepancyQtyQtl ?? 0) > 0
    });
    notifyUser(
      existingVoucher
        ? `Voucher ${createdVoucher.voucherNo} updated.`
        : `Voucher ${createdVoucher.voucherNo} generated.`
    );
  }

  async function markVoucherPaid(voucher: FinancialVoucher) {
    if (!requirePermission("canVoucher", "Your role cannot update financial vouchers.")) {
      return;
    }

    const response = await fetchWithAuth(`${API_BASE}/api/seed/financial-vouchers/${voucher.id}/paid`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({})
    });

    if (!response.ok) {
      const errorBody = await response.json().catch(() => null);
      throw new Error(errorBody?.message || "Unable to mark voucher as paid.");
    }

    const data = (await response.json()) as AppBootstrapPayload;

    setRegistrations(data.registrations ?? []);
    setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
    setStacks(data.stacks?.length ? data.stacks : defaultStacks);
    setLots(data.lots ?? []);
    setDiscrepancies(data.discrepancies ?? []);
    setDiscrepancyShifts(data.discrepancyShifts ?? []);
    setStackAccommodations(data.stackAccommodations ?? []);
    setFinancialVouchers(data.financialVouchers ?? []);
    setReceipts(data.receipts ?? []);
    setFeatures({
      discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
    });

    if (voucherPreview?.voucher.id === voucher.id) {
      const refreshed =
        (data.financialVouchers ?? []).find((item) => item.id === voucher.id) ?? null;
      setVoucherPreview(
        refreshed
          ? {
              voucher: refreshed,
              hasDiscrepancy: Number(refreshed.discrepancyQtyQtl ?? 0) > 0
            }
          : null
      );
    }

    notifyUser(`Voucher ${voucher.voucherNo} marked as paid.`);
  }

  async function deleteVoucher(voucher: FinancialVoucher, adminPassword = "") {
    if (!requirePermission("canVoucher", "Your role cannot delete financial vouchers.")) {
      return;
    }

    const confirmed = confirmDestructiveAction({
      itemLabel: `Voucher No.: ${voucher.voucherNo}\nFarmer: ${voucher.farmerName}`
    });
    if (!confirmed) {
      notifyUser("Delete cancelled. No financial voucher was deleted.", false);
      return;
    }

    let overridePassword = adminPassword;
    if (isVoucherLockedStatus(voucher.status) && !overridePassword) {
      const prompted = promptAdminPassword(voucher.voucherNo, "Deleting");
      if (prompted === null) {
        return;
      }
      overridePassword = prompted;
    }

    const response = await fetchWithAuth(`${API_BASE}/api/seed/financial-vouchers/${voucher.id}`, {
      method: "DELETE",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        adminPassword: overridePassword
      })
    });

    if (!response.ok) {
      const errorBody = await response.json().catch(() => null);
      throw new Error(errorBody?.message || "Unable to delete financial voucher.");
    }

    const data = (await response.json()) as AppBootstrapPayload;

    setRegistrations(data.registrations ?? []);
    setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
    setStacks(data.stacks?.length ? data.stacks : defaultStacks);
    setLots(data.lots ?? []);
    setDiscrepancies(data.discrepancies ?? []);
    setDiscrepancyShifts(data.discrepancyShifts ?? []);
    setStackAccommodations(data.stackAccommodations ?? []);
    setFinancialVouchers(data.financialVouchers ?? []);
    setReceipts(data.receipts ?? []);
    setFeatures({
      discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
    });

    if (voucherPreview?.voucher.id === voucher.id) {
      closeVoucherModal();
    }

    notifyUser(`Voucher ${voucher.voucherNo} deleted.`);
  }

  async function addVoucherPayment() {
    if (!requirePermission("canVoucher", "Your role cannot manage voucher payments.")) {
      return;
    }
    if (!paymentLedgerVoucher) {
      notifyUser("Select a voucher first.");
      return;
    }
    if (isAddingVoucherPayment) {
      notifyUser("Payment save is already in progress. Please wait.");
      return;
    }
    const amount = Number(paymentAmount || 0);
    if (!Number.isFinite(amount) || amount <= 0) {
      notifyUser("Enter payment amount.");
      return;
    }
    if (!paymentTransactionNo.trim()) {
      notifyUser("Enter transaction number.");
      return;
    }

    setIsAddingVoucherPayment(true);
    try {
      const isEditingPayment = Boolean(editingVoucherPaymentId);
      const response = await fetchWithAuth(
        isEditingPayment
          ? `${API_BASE}/api/seed/financial-vouchers/${paymentLedgerVoucher.id}/payments/${editingVoucherPaymentId}`
          : `${API_BASE}/api/seed/financial-vouchers/${paymentLedgerVoucher.id}/payments`,
        {
        method: isEditingPayment ? "PUT" : "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          paymentDate,
          amount,
          transactionNo: paymentTransactionNo.trim(),
          remarks: paymentRemarks,
          adminPassword: paymentAdminPassword
        })
      });

      if (!response.ok) {
        const errorBody = await response.json().catch(() => null);
        throw new Error(errorBody?.message || "Unable to save voucher payment.");
      }

      const data = (await response.json()) as FullBootstrapPayload;
      applyFullBootstrap(data);
      if (voucherPreview?.voucher.id === paymentLedgerVoucher.id) {
        const refreshed =
          (data.financialVouchers ?? []).find((item) => item.id === paymentLedgerVoucher.id) ?? null;
        setVoucherPreview(
          refreshed
            ? {
                voucher: refreshed,
                hasDiscrepancy: Number(refreshed.discrepancyQtyQtl ?? 0) > 0
              }
            : null
        );
      }
      setPaymentDate(new Date().toISOString().slice(0, 10));
      setPaymentAmount("");
      setPaymentTransactionNo("");
      setPaymentRemarks("");
      setEditingVoucherPaymentId("");
      setPaymentAdminPassword("");
      notifyUser(
        isEditingPayment
          ? `Payment entry updated for voucher ${paymentLedgerVoucher.voucherNo}.`
          : `Payment recorded for voucher ${paymentLedgerVoucher.voucherNo}.`
      );
    } finally {
      setIsAddingVoucherPayment(false);
    }
  }

  function renderVoucherPdfPage(doc: jsPDF, voucher: FinancialVoucher, addPage = false) {
    if (addPage) {
      doc.addPage("a5", "portrait");
    }
    const left = 10;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const right = pageWidth - 10;
    const hasDiscrepancy = Number(voucher.discrepancyQtyQtl ?? 0) > 0;
    const voucherMetaFontSize = 8.8;
    let y = 8;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    y = renderPdfBrandHeader(doc, "FARMER OVERALL INTAKE CUM PAYMENT VOUCHER", {
      left,
      right,
      y,
      compact: true,
      logoDataUrl: pdfLogoDataUrl
    });

    const topLines = [
      ["Voucher No.", voucher.voucherNo],
      ["Voucher Date", formatDateDisplay(voucher.voucherDate)],
      ["Season", `${voucher.season} ${voucher.year}`],
      ["Reg. Code", voucher.cropRegistrationCode]
    ];
    doc.setFontSize(voucherMetaFontSize);
    topLines.forEach(([label, value]) => {
      doc.setFont("helvetica", "bold");
      doc.text(label, left, y);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${value}`, left + 24, y);
      y += 4.2;
    });
    doc.setFontSize(9);

    const detailLines = [
      ["Farmer Name", voucher.farmerName],
      ["F/H Name", voucher.fatherName || "-"],
      ["Village", voucher.village || "-"],
      ["Block", voucher.block || "-"],
      ["District", voucher.district || "-"],
      ["Crop", voucher.crop],
      ["Variety", voucher.variety],
      ["Class", voucher.classStage],
      ["Exp. Yield", `${formatNumber(voucher.expectedYieldQtl)} QTL`]
    ];
    const colWidth = (right - left - 8) / 2;
    detailLines.forEach((line, index) => {
      const col = index % 2;
      const row = Math.floor(index / 2);
      const rowY = y + row * 4.2;
      const x = left + col * (colWidth + 8);
      doc.setFont("helvetica", "bold");
      doc.text(line[0], x, rowY);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${line[1]}`, x + 20, rowY, { maxWidth: colWidth - 20 });
    });
    y += Math.ceil(detailLines.length / 2) * 4.2 + 1.5;

    doc.line(left, y, right, y);
    y += 4;
    doc.setFont("helvetica", "bold");
    doc.text("Overall Intake Detail", left, y);
    y += 2.5;
    doc.line(left, y, right, y);
    y += 1;

    autoTable(doc, {
      startY: y,
      head: [["Date", "Rcpt No", "Vehicle No", "Stack", "Bags", "Gross", "Net"]],
      body: [
        ...voucher.lines.map((line) => [
          formatDateDisplay(line.receiptDate),
          line.receiptNo,
          line.vehicleNo,
          line.stackNo,
          String(line.bags),
          formatNumber(line.grossQtyQtl),
          formatNumber(line.netQtyQtl)
        ]),
        [
          "Total",
          "",
          "",
          "",
          String(voucher.totalBags),
          formatNumber(voucher.totalGrossQtyQtl),
          formatNumber(voucher.totalNetQtyQtl)
        ]
      ],
      theme: "grid",
      styles: {
        fontSize: 7.4,
        cellPadding: 1.2,
        textColor: [47, 30, 18],
        lineColor: [216, 200, 183]
      },
      headStyles: {
        fillColor: [243, 237, 229],
        textColor: [47, 30, 18],
        fontStyle: "bold"
      },
      margin: { left, right: pageWidth - right },
      didParseCell: (data) => {
        if (data.row.index === voucher.lines.length) {
          data.cell.styles.fontStyle = "bold";
        }
      }
    });

    y = ((doc as jsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? y) + 4;
    if (hasDiscrepancy) {
      doc.setFont("helvetica", "bold");
      doc.text("Discrepancy Summary", left, y);
      y += 4.2;
      doc.setFont("helvetica", "bold");
      doc.text("Excess Qty", left, y);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${formatNumber(voucher.discrepancyQtyQtl)} QTL`, left + 24, y);
      y += 4.2;
      doc.setFont("helvetica", "bold");
      doc.text("Excess Bags", left, y);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${voucher.discrepancyBags}`, left + 24, y);
      y += 5;
    }

    doc.line(left, y, right, y);
    y += 4;
    doc.setFont("helvetica", "bold");
    doc.text("Payment Calculation", left, y);
    y += 2.5;
    doc.line(left, y, right, y);
    y += 1;

    const paymentBody = [
      [
        "Certified Qty",
        formatNumber(voucher.certifiedQtyQtl),
        formatNumber(voucher.certifiedRatePerQtl),
        formatNumber(voucher.certifiedAmount)
      ]
    ];
    if (hasDiscrepancy) {
      paymentBody.push([
        "Discrepancy Qty",
        formatNumber(voucher.discrepancyQtyQtl),
        formatNumber(voucher.discrepancyRatePerQtl),
        formatNumber(voucher.discrepancyAmount)
      ]);
    }
    paymentBody.push(
      ["Gross Payable Amount", "", "", formatNumber(voucher.grossPayableAmount)],
      ["Seed Payment", "", "", formatNumber(voucher.deductionAmount)],
      ["Net Payable Amount", "", "", formatNumber(voucher.netPayableAmount)],
      ["Rounded Off", "", "", formatNumber(Number(voucher.roundedOffAmount ?? 0))],
      ["Final Payable Amount", "", "", formatNumber(getVoucherFinalPayable(voucher))]
    );

    autoTable(doc, {
      startY: y,
      head: [["Type", "Qty (QTL)", "Rate/QTL", "Amount"]],
      body: paymentBody,
      theme: "grid",
      styles: {
        fontSize: 7.4,
        cellPadding: 1.2,
        textColor: [47, 30, 18],
        lineColor: [216, 200, 183]
      },
      headStyles: {
        fillColor: [243, 237, 229],
        textColor: [47, 30, 18],
        fontStyle: "bold"
      },
      margin: { left, right: pageWidth - right },
      didParseCell: (data) => {
        if (data.row.index >= paymentBody.length - 3) {
          data.cell.styles.fontStyle = "bold";
        }
      }
    });

    y = ((doc as jsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? y) + 5;
    if (y > pageHeight - 15) {
      doc.addPage();
      y = 10;
    }
    doc.setFont("helvetica", "bold");
    doc.text("Approval Status", left, y);
    y += 5;
    doc.setFont("helvetica", "normal");
    doc.text("Prepared By        : __________________", left, y);
    doc.text("Checked By         : __________________", left + 55, y);
    y += 5;
    doc.text("Approved By        : __________________", left, y);
    doc.text("Farmer Signature   : __________________", left + 55, y);
    y += 5;
    doc.line(left, y, right, y);
    y += 4;
    doc.line(left, y, right, y);
  }

  function downloadVoucherPdf(voucher: FinancialVoucher) {
    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a5"
    });
    renderVoucherPdfPage(doc, voucher);

    doc.save(`${voucher.voucherNo.replace(/[^A-Za-z0-9_-]+/g, "_")}.pdf`);
    notifyUser(`Voucher ${voucher.voucherNo} downloaded.`);
  }

  function downloadBulkVoucherPdf(vouchers: FinancialVoucher[]) {
    if (!vouchers.length) {
      notifyUser("No vouchers match the current filter for bulk download.");
      return;
    }

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a5"
    });

    vouchers.forEach((voucher, index) => {
      renderVoucherPdfPage(doc, voucher, index > 0);
    });

    const districtLabel = voucherRegisterDistrictFilter || "ALL_DISTRICTS";
    doc.save(`financial_vouchers_${districtLabel.replace(/[^A-Za-z0-9_-]+/g, "_")}.pdf`);
    notifyUser(`${vouchers.length} voucher(s) downloaded in one PDF.`);
  }

  function renderPaymentLedgerPdfPage(doc: jsPDF, voucher: FinancialVoucher, addPage = false) {
    if (addPage) {
      doc.addPage("a5", "portrait");
    }

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const left = 6;
    const right = pageWidth - 6;
    let y = 5;
    const grossPayableAmount = roundQtl(Number(voucher.grossPayableAmount ?? 0));
    const deductionAmount = roundQtl(Number(voucher.deductionAmount ?? 0));
    const roundedAmount = roundQtl(Number(voucher.roundedOffAmount ?? 0));
    const sortedPayments = (voucher.payments ?? [])
      .slice()
      .sort((leftPayment, rightPayment) => {
        const dateCompare = leftPayment.paymentDate.localeCompare(rightPayment.paymentDate);
        if (dateCompare !== 0) {
          return dateCompare;
        }
        return (leftPayment.transactionNo || "").localeCompare(rightPayment.transactionNo || "", "en", {
          sensitivity: "base"
        });
      });

    const ledgerRows: Array<[string, string, string, string, string, string]> = [];
    let runningBalance = 0;

    runningBalance = roundQtl(runningBalance + grossPayableAmount);
    ledgerRows.push([
      formatDateDisplay(voucher.voucherDate),
      "Voucher Raised",
      "-",
      formatNumber(grossPayableAmount),
      "-",
      formatNumber(runningBalance)
    ]);

    if (deductionAmount !== 0) {
      runningBalance = roundQtl(runningBalance - deductionAmount);
      ledgerRows.push([
        formatDateDisplay(voucher.voucherDate),
        "Deduction Applied",
        "-",
        "-",
        formatNumber(deductionAmount),
        formatNumber(runningBalance)
      ]);
    }

    if (roundedAmount !== 0) {
      const roundedTransactionType = roundedAmount > 0 ? "Rounded Off Added" : "Rounded Off Less";
      if (roundedAmount > 0) {
        runningBalance = roundQtl(runningBalance + roundedAmount);
        ledgerRows.push([
          formatDateDisplay(voucher.voucherDate),
          roundedTransactionType,
          "-",
          formatNumber(roundedAmount),
          "-",
          formatNumber(runningBalance)
        ]);
      } else {
        runningBalance = roundQtl(runningBalance - Math.abs(roundedAmount));
        ledgerRows.push([
          formatDateDisplay(voucher.voucherDate),
          roundedTransactionType,
          "-",
          "-",
          formatNumber(Math.abs(roundedAmount)),
          formatNumber(runningBalance)
        ]);
      }
    }

    sortedPayments.forEach((payment) => {
      const amount = roundQtl(Number(payment.amount ?? 0));
      runningBalance = roundQtl(runningBalance - amount);
      ledgerRows.push([
        formatDateDisplay(payment.paymentDate),
        payment.mode ? `Payment by ${payment.mode}` : "Payment",
        payment.transactionNo || "-",
        "-",
        formatNumber(amount),
        formatNumber(runningBalance)
      ]);
    });

    y = renderPdfBrandHeader(doc, "FARMER VOUCHER LEDGER", {
      left,
      right,
      y,
      compact: true,
      logoDataUrl: pdfLogoDataUrl
    });

    const metaLines = [
      ["Voucher No.", voucher.voucherNo],
      ["Voucher Date", formatDateDisplay(voucher.voucherDate)],
      ["Season", `${voucher.season} ${voucher.year}`],
      ["Farmer Name", voucher.farmerName],
      ["F/H Name", voucher.fatherName || "-"],
      ["Reg. Code", voucher.cropRegistrationCode],
      ["Village", voucher.village || "-"],
      ["Block", voucher.block || "-"],
      ["District", voucher.district || "-"]
    ];

    doc.setFontSize(6.6);
    metaLines.forEach(([label, value], index) => {
      const x = index % 2 === 0 ? left : left + 70;
      const rowY = y + Math.floor(index / 2) * 3.6;
      doc.setFont("helvetica", "bold");
      doc.text(label, x, rowY);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${value}`, x + 22, rowY, { maxWidth: 42 });
    });
    y += Math.ceil(metaLines.length / 2) * 3.6 + 1.5;

    const voucherRemark = String(voucher.remarks ?? "").trim();
    if (voucherRemark) {
      doc.setFontSize(6.5);
      doc.setFont("helvetica", "bold");
      doc.text("Voucher Remark", left, y);
      doc.text(`: ${voucherRemark}`, left + 24, y, { maxWidth: right - (left + 24) });
      y += 3.8;
    }

    y += 0.6;
    doc.line(left, y, right, y);
    y += 3;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7);
    doc.text("Receipt / Intake Detail", left, y);
    y += 2;

    autoTable(doc, {
      startY: y,
      head: [["S.No.", "Receipt No.", "Date", "Vehicle No.", "Stack", "Bags", "Gross Qty", "Net Qty"]],
      body: [
        ...voucher.lines.map((line, index) => [
          String(index + 1),
          line.receiptNo,
          formatDateDisplay(line.receiptDate),
          line.vehicleNo,
          line.stackNo,
          String(line.bags),
          formatNumber(line.grossQtyQtl),
          formatNumber(line.netQtyQtl)
        ]),
        [
          "",
          "",
          "",
          "Total",
          "",
          String(voucher.totalBags),
          formatNumber(voucher.totalGrossQtyQtl),
          formatNumber(voucher.totalNetQtyQtl)
        ]
      ],
      margin: { left, right },
      styles: {
        font: "helvetica",
        fontSize: 4.8,
        cellPadding: 0.55,
        lineColor: [160, 140, 120],
        lineWidth: 0.1
      },
      headStyles: {
        fillColor: [91, 61, 38],
        textColor: 255,
        fontStyle: "bold"
      },
      columnStyles: {
        0: { cellWidth: 8, halign: "center" },
        1: { cellWidth: 15 },
        2: { cellWidth: 14 },
        3: { cellWidth: 34 },
        4: { cellWidth: 9 },
        5: { cellWidth: 9, halign: "right" },
        6: { cellWidth: 17, halign: "right" },
        7: { cellWidth: 17, halign: "right" }
      },
      didParseCell: (data) => {
        if (data.row.index === voucher.lines.length) {
          data.cell.styles.fontStyle = "bold";
        }
      }
    });

    y = ((doc as jsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? y) + 2.5;
    if (y > pageHeight - 34) {
      doc.addPage("a5", "portrait");
      y = 6;
    }

    doc.line(left, y, right, y);
    y += 3;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7);
    doc.text("Payment Calculation", left, y);
    y += 2;

    autoTable(doc, {
      startY: y,
      head: [["Particulars", "Qty (QTL)", "Rate/QTL", "Amount"]],
      body: [
        [
          "Certified Seed",
          formatNumber(voucher.certifiedQtyQtl),
          formatNumber(voucher.certifiedRatePerQtl),
          formatNumber(voucher.certifiedAmount)
        ],
        [
          "Discrepancy Seed",
          formatNumber(voucher.discrepancyQtyQtl),
          formatNumber(voucher.discrepancyRatePerQtl),
          formatNumber(voucher.discrepancyAmount)
        ],
        ["Gross Payable", "", "", formatNumber(grossPayableAmount)],
      ],
      margin: { left, right },
      styles: {
        font: "helvetica",
        fontSize: 4.9,
        cellPadding: 0.55,
        lineColor: [160, 140, 120],
        lineWidth: 0.1
      },
      headStyles: {
        fillColor: [91, 61, 38],
        textColor: 255,
        fontStyle: "bold"
      },
      columnStyles: {
        0: { cellWidth: 47 },
        1: { cellWidth: 22, halign: "right" },
        2: { cellWidth: 22, halign: "right" },
        3: { cellWidth: 36, halign: "right" }
      },
      didParseCell: (data) => {
        if (data.row.index >= 2) {
          data.cell.styles.fontStyle = data.row.index === 2 ? "bold" : "normal";
        }
        if (data.row.index === 2) {
          data.cell.styles.fillColor = [244, 238, 229];
        }
      }
    });

    y = ((doc as jsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? y) + 2.5;
    if (y > pageHeight - 34) {
      doc.addPage("a5", "portrait");
      y = 6;
    }

    doc.line(left, y, right, y);
    y += 3;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7);
    doc.text("Financial Transactions", left, y);
    y += 2;

    autoTable(doc, {
      startY: y,
      head: [["S.No.", "Date", "Transaction Type", "Transaction No.", "Debit", "Credit", "Balance"]],
      body: ledgerRows.map((row, index) => [String(index + 1), ...row]),
      margin: { left, right },
      styles: {
        font: "helvetica",
        fontSize: 4.8,
        cellPadding: 0.55,
        lineColor: [160, 140, 120],
        lineWidth: 0.1
      },
      headStyles: {
        fillColor: [91, 61, 38],
        textColor: 255,
        fontStyle: "bold"
      },
      columnStyles: {
        0: { cellWidth: 7, halign: "center" },
        1: { cellWidth: 13 },
        2: { cellWidth: 31 },
        3: { cellWidth: 29 },
        4: { cellWidth: 15, halign: "right" },
        5: { cellWidth: 15, halign: "right" },
        6: { cellWidth: 15, halign: "right" }
      }
    });
  }

  function downloadPaymentLedgerPdf(voucher: FinancialVoucher) {
    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a5"
    });
    renderPaymentLedgerPdfPage(doc, voucher);
    doc.save(`${voucher.voucherNo.replace(/[^A-Za-z0-9_-]+/g, "_")}_ledger.pdf`);
    notifyUser(`Payment ledger for ${voucher.voucherNo} downloaded.`);
  }

  function downloadBulkPaymentLedgerPdf(vouchers: FinancialVoucher[]) {
    if (!vouchers.length) {
      notifyUser("No vouchers match the current filter for bulk ledger download.");
      return;
    }

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a5"
    });

    vouchers.forEach((voucher, index) => {
      renderPaymentLedgerPdfPage(doc, voucher, index > 0);
    });

    const districtLabel = voucherRegisterDistrictFilter || "ALL_DISTRICTS";
    doc.save(`financial_ledgers_${districtLabel.replace(/[^A-Za-z0-9_-]+/g, "_")}.pdf`);
    notifyUser(`${vouchers.length} ledger(s) downloaded in one PDF.`);
  }

  function downloadOrganizerWiseBulkPaymentLedgerPdf(vouchers: FinancialVoucher[]) {
    if (!voucherRegisterOrganizerFilter) {
      notifyUser("Select an organizer first for organizer-wise bulk ledger PDF.");
      return;
    }
    const organizerVouchers = vouchers.filter(
      (voucher) => getVoucherOrganizerName(voucher) === voucherRegisterOrganizerFilter
    );
    if (!organizerVouchers.length) {
      notifyUser("No vouchers match this organizer and current filters.");
      return;
    }

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a5"
    });

    organizerVouchers.forEach((voucher, index) => {
      renderPaymentLedgerPdfPage(doc, voucher, index > 0);
    });

    const organizerLabel = voucherRegisterOrganizerFilter.replace(/[^A-Za-z0-9_-]+/g, "_");
    doc.save(`financial_ledgers_organizer_${organizerLabel}.pdf`);
    notifyUser(`${organizerVouchers.length} organizer-wise ledger(s) downloaded in one PDF.`);
  }

  function downloadOrganizerCommissionVoucherPdf(summary: OrganizerCommissionRow) {
    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4"
    });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const left = 12;
    const right = pageWidth - 12;
    let y = 12;
    const payments = organizerPayments
      .filter((item) => item.organizerId === summary.organizer.id)
      .slice()
      .sort((leftPayment, rightPayment) => {
        const dateCompare = leftPayment.paymentDate.localeCompare(rightPayment.paymentDate);
        if (dateCompare !== 0) {
          return dateCompare;
        }
        return leftPayment.transactionNo.localeCompare(rightPayment.transactionNo, "en", {
          sensitivity: "base"
        });
      });

    y = renderPdfBrandHeader(doc, "ORGANIZER COMMISSION VOUCHER", {
      left,
      right,
      y,
      compact: false,
      logoDataUrl: pdfLogoDataUrl
    });

    const summaryLines: Array<[string, string]> = [
      ["Organizer Name", summary.organizer.name],
      ["Mobile", summary.organizer.mobile || "-"],
      ["Village", summary.organizer.village || "-"],
      ["District", summary.organizer.district || "-"],
      ["Commission Rate / QTL", formatNumber(summary.ratePerQtl)],
      ["Linked Farmers", String(summary.farmerCount)],
      ["Total Intake Qty", `${formatNumber(summary.totalIntakeQtl)} QTL`],
      ["Gross Commission", formatNumber(summary.grossCommissionAmount)],
      ["Deduction Amount", formatNumber(summary.deductionAmount)],
      ["Net Payable", formatNumber(summary.netPayableAmount)],
      ["Commission Paid", formatNumber(summary.paidAmount)],
      ["Balance", formatNumber(summary.balanceAmount)]
    ];

    summaryLines.forEach(([label, value], index) => {
      const x = index % 2 === 0 ? left : left + 95;
      const rowY = y + Math.floor(index / 2) * 5.2;
      doc.setFont("helvetica", "bold");
      doc.setFontSize(9);
      doc.text(label, x, rowY);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${value}`, x + 34, rowY, { maxWidth: 58 });
    });
    y += Math.ceil(summaryLines.length / 2) * 5.2 + 4;

    doc.line(left, y, right, y);
    y += 5;
    doc.setFont("helvetica", "bold");
    doc.text("Linked Farmer Intake Detail", left, y);
    y += 3;

    autoTable(doc, {
      startY: y,
      head: [["S.No.", "Reg. Code", "Farmer Name", "Village", "Received QTL", "Rate/QTL", "Commission"]],
      body: (summary.linkedRegistrations.length
        ? summary.linkedRegistrations
            .slice()
            .sort((leftReg, rightReg) =>
              leftReg.cropRegistrationCode.localeCompare(rightReg.cropRegistrationCode, "en", {
                sensitivity: "base",
                numeric: true
              })
            )
            .map((registration, index) => [
              String(index + 1),
              registration.cropRegistrationCode,
              registration.farmerName,
              registration.village || "-",
              formatNumber(Number(registration.totalReceivedQtl ?? 0)),
              formatNumber(
                Number(registration.organizerCommissionRatePerQtl ?? summary.ratePerQtl)
              ),
              formatNumber(
                roundQtl(
                  Number(registration.totalReceivedQtl ?? 0) *
                    Number(registration.organizerCommissionRatePerQtl ?? summary.ratePerQtl)
                )
              )
            ])
        : [["-", "-", "No linked farmers", "-", "-", "-", "-"]]),
      margin: { left, right },
      styles: {
        font: "helvetica",
        fontSize: 7.2,
        cellPadding: 1.3,
        lineColor: [160, 140, 120],
        lineWidth: 0.1
      },
      headStyles: {
        fillColor: [91, 61, 38],
        textColor: 255,
        fontStyle: "bold"
      },
      columnStyles: {
        0: { cellWidth: 12 },
        1: { cellWidth: 24 },
        2: { cellWidth: 50 },
        3: { cellWidth: 30 },
        4: { cellWidth: 22, halign: "right" },
        5: { cellWidth: 20, halign: "right" },
        6: { cellWidth: 24, halign: "right" }
      }
    });

    y = ((doc as jsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? y) + 6;
    if (y > pageHeight - 70) {
      doc.addPage();
      y = 12;
    }

    doc.line(left, y, right, y);
    y += 5;
    doc.setFont("helvetica", "bold");
    doc.text("Commission Payment History", left, y);
    y += 3;

    autoTable(doc, {
      startY: y,
      head: [["S.No.", "Date", "Transaction No.", "Amount", "Remarks", "Entered By"]],
      body: (payments.length
        ? payments.map((payment, index) => [
            String(index + 1),
            formatDateDisplay(payment.paymentDate),
            payment.transactionNo,
            formatNumber(Number(payment.amount ?? 0)),
            payment.remarks || "-",
            payment.createdBy || "-"
          ])
        : [["-", "-", "No payment recorded", "-", "-", "-"]]),
      margin: { left, right },
      styles: {
        font: "helvetica",
        fontSize: 7.1,
        cellPadding: 1.3,
        lineColor: [160, 140, 120],
        lineWidth: 0.1
      },
      headStyles: {
        fillColor: [91, 61, 38],
        textColor: 255,
        fontStyle: "bold"
      },
      columnStyles: {
        0: { cellWidth: 12 },
        1: { cellWidth: 22 },
        2: { cellWidth: 38 },
        3: { cellWidth: 24, halign: "right" },
        4: { cellWidth: 58 },
        5: { cellWidth: 24 }
      }
    });

    y = ((doc as jsPDF & { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? y) + 6;
    if (y > pageHeight - 30) {
      doc.addPage();
      y = 12;
    }

    doc.line(left, y, right, y);
    y += 5;
    doc.setFont("helvetica", "bold");
    doc.text("Voucher Totals", left, y);
    y += 5;

    [
      ["Gross Commission", formatNumber(summary.grossCommissionAmount)],
      ["Deduction Amount", formatNumber(summary.deductionAmount)],
      ["Net Commission Payable", formatNumber(summary.netPayableAmount)],
      ["Total Commission Paid", formatNumber(summary.paidAmount)],
      ["Closing Balance", formatNumber(summary.balanceAmount)]
    ].forEach(([label, value]) => {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(9);
      doc.text(label, left, y);
      doc.setFont("helvetica", "normal");
      doc.text(`: ${value}`, left + 46, y);
      y += 5;
    });

    y += 4;
    doc.line(left, y, right, y);
    y += 8;
    doc.setFont("helvetica", "normal");
    doc.text("Prepared By          : __________________", left, y);
    doc.text("Checked By          : __________________", left + 70, y);
    y += 7;
    doc.text("Approved By         : __________________", left, y);
    doc.text("Organizer Sign      : __________________", left + 70, y);
    y += 5;
    doc.line(left, y, right, y);

    doc.save(
      `${summary.organizer.name.replace(/[^A-Za-z0-9_-]+/g, "_")}_commission_voucher.pdf`
    );
    notifyUser(`Commission voucher for ${summary.organizer.name} downloaded.`);
  }

  function printSlipPdf() {
    if (!slipPreview) {
      notifyUser("Load a slip preview first.");
      return;
    }

    try {
      const doc = new jsPDF({
        orientation: slipPreview.orientation,
        unit: "mm",
        format: slipPreview.pageSize.toLowerCase() as "a4" | "a5"
      });
      const fileName = `${(slipPreview.slipNo ?? slipPreview.title).replace(/[^A-Za-z0-9_-]+/g, "_")}.pdf`;
      renderSlipPreviewPdf(doc, slipPreview);
      doc.save(fileName);
      notifyUser("Slip PDF downloaded.");
    } catch (error) {
      notifyUser(error instanceof Error ? error.message : "Unable to generate slip PDF.");
    }
  }

  function downloadAllOverallSlips() {
    const targetRegistrations = slipRegistrationRows.filter((item) => item.totalReceivedQtl > 0);
    if (targetRegistrations.length === 0) {
      notifyUser("No registration with intake found for overall slip download.");
      return;
    }

    try {
      const doc = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a5"
      });

      let renderedCount = 0;
      targetRegistrations.forEach((registration, index) => {
        const preview = buildSlipPreview({
          slipType: "FARMER_OVERALL",
          slipRegistrationId: registration.id
        });
        if (!preview) {
          return;
        }
        renderSlipPreviewPdf(doc, preview, renderedCount > 0 || index > 0);
        renderedCount += 1;
      });

      if (renderedCount === 0) {
        notifyUser("Unable to build overall slips for the selected filters.");
        return;
      }

      const scopeParts = [
        slipDistrictFilter || "ALLDIST",
        slipVillageFilter || "ALLVILL",
        slipClassFilter || "ALLCLASS"
      ];
      const fileName = `overall_slips_${scopeParts
        .join("_")
        .replace(/[^A-Za-z0-9_-]+/g, "_")}.pdf`;
      doc.save(fileName);
      notifyUser(`${renderedCount} overall slip(s) downloaded.`);
    } catch (error) {
      notifyUser(error instanceof Error ? error.message : "Unable to download overall slips.");
    }
  }

  function exportImportViewToExcel() {
    const exportRows = sortedImportRegistrations.map((item, index) => ({
      "S.No.": index + 1,
      "Registration Code": item.cropRegistrationCode,
      "Farmer Name": item.farmerName,
      "Father Name": item.fatherName,
      Village: item.village,
      District: item.district,
      Crop: item.crop,
      Variety: item.variety,
      "Seed Class": item.classStage,
      "Certified Area (Ha)": item.certifiedAreaHa,
      "Expected Yield (QTL)": item.expectedYieldQtl,
      "Received (QTL)": item.totalReceivedQtl,
      "Balance (QTL)": item.balanceQtl,
      Status: item.status
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Farmer Master");
    XLSX.writeFile(workbook, "farmer-master-view.xlsx");
    notifyUser("Farmer master list downloaded.");
  }

  function toggleRegistrationSort(column: RegistrationSortKey) {
    if (registrationSortBy === column) {
      setRegistrationSortDirection((current) => (current === "asc" ? "desc" : "asc"));
      return;
    }

    setRegistrationSortBy(column);
    setRegistrationSortDirection("asc");
  }

  function exportRegistrationMasterToExcel() {
    const exportRows = sortedRegistrationRows.map((item, index) => ({
      "S.No.": index + 1,
      "Registration Code": item.cropRegistrationCode,
      "Farmer Name": item.farmerName,
      Village: item.village,
      Crop: item.crop,
      Variety: item.variety,
      "Seed Class": item.classStage,
      "Expected Yield (QTL)": item.expectedYieldQtl,
      "Received (QTL)": item.totalReceivedQtl,
      "Balance (QTL)": item.balanceQtl,
      Status: item.status
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Registration Master");
    XLSX.writeFile(workbook, "registration-master.xlsx");
    notifyUser("Registration master downloaded.");
  }

  if (!currentUser || !currentPermissions) {
    return (
      <main className="loginShell">
        <section className="loginCard">
          <div className="loginTitleBlock">
            <img className="loginLogo" src={BRAND_LOGO_SRC} alt="Krishiv Seeds" />
            <h1>Seed Intake Platform</h1>
            <p>Login to continue.</p>
          </div>

          <div className="formGrid">
            <label>
              <span>User ID</span>
              <input value={loginEmail} onChange={(event) => setLoginEmail(event.target.value)} />
            </label>
            <label>
              <span>Password</span>
              <input
                type="password"
                value={loginPassword}
                onChange={(event) => setLoginPassword(event.target.value)}
              />
            </label>
          </div>

          {loginError ? <p className="loginError">{loginError}</p> : null}

          <div className="actionsFooter">
            <button
              className="primaryButton"
              type="button"
              onClick={() => {
                void handleLogin().catch((error) => {
                  const message = error instanceof Error ? error.message : "Unable to login.";
                  setLoginError(message);
                  notifyUser(message);
                });
              }}
            >
              Login
            </button>
          </div>
        </section>
      </main>
    );
  }

  return (
    <main className="appShell">
      <header className="topBar">
        <div className="topBarBrand">
          <img className="topBarLogo" src={BRAND_LOGO_SRC} alt="Krishiv Seeds" />
          <div>
            <h1>Seed Intake Platform</h1>
            <p>{reportSeasonLabel}</p>
          </div>
        </div>
        <div className="headerMeta">
          <span>{currentUser.name} ({currentUser.role})</span>
          <span>{registrations.length} registrations</span>
          <span>{lots.length} lots</span>
          <span>{receipts.length} receipts</span>
          <button className="ghostButton" type="button" onClick={handleLogout}>
            Logout
          </button>
        </div>
      </header>

      {!hasLoadedCoreData ? (
        <div className="infoStrip">Loading core data...</div>
      ) : isLoadingOperationalData && !hasLoadedOperationalData ? (
        <div className="infoStrip">Loading detailed operational data in the background...</div>
      ) : null}

      <div className="layout">
        <aside className="sidebar">
          <div className="sidebarBrand">
            <img className="sidebarLogo" src={BRAND_LOGO_SRC} alt="Krishiv Seeds" />
            <div>
              <strong>{COMPANY_NAME}</strong>
              <p>Raw seed intake and certification traceability</p>
            </div>
          </div>

          <nav className="navList">
            {visibleNavSections.map((section) => {
              const isExpanded = expandedNavSections[section.key] || section.key === activeNavSectionKey;
              const hasActiveChild = section.items.some((item) => item.key === activeView);
              return (
                <div className={`navSection ${hasActiveChild ? "active" : ""}`} key={section.key}>
                  <button
                    className={`navSectionButton ${hasActiveChild ? "active" : ""}`}
                    onClick={() => toggleNavSection(section.key)}
                    type="button"
                  >
                    <span>{section.label}</span>
                    <span className="navSectionToggle">{isExpanded ? "−" : "+"}</span>
                  </button>
                  {isExpanded ? (
                    <div className="navSubList">
                      {section.items.map((item) => (
                        <button
                          className={`navSubButton ${activeView === item.key ? "active" : ""}`}
                          disabled={isNavItemDisabled(item.key)}
                          key={item.key}
                          onClick={() => openViewFromNav(item.key, section.key)}
                          type="button"
                        >
                          {item.label}
                        </button>
                      ))}
                    </div>
                  ) : null}
                </div>
              );
            })}
          </nav>

        </aside>

        <section className="contentArea">
          <div className="contentHeader">
            <h2>{activeNavItem?.label ?? "KRISHIV Seed Intake"}</h2>
            <p>
              {toast ||
                (features.discrepancyWorkflow
                  ? "Discrepancy workflow is enabled. Excess intake is saved and flagged for later shift."
                  : "Allowed intake in v1 is capped at expected yield. No override is enabled.")}
            </p>
            {activeView === "dashboard" ? <p>{shortSnapshot}</p> : null}
          </div>

          {activeView === "dashboard" && (
            <div className="contentStack">
              <section className="panel dashboardHero">
                <div>
                  <p className="eyebrow">Operational Control</p>
                  <h3>Seed Intake Command Dashboard</h3>
                  <p className="dashboardIntro">
                    Track intake progress, stock pressure, payment status, and discrepancy exposure from one
                    screen. Use this view as the daily working dashboard for officers and management.
                  </p>
                </div>
                <div className="dashboardHeroStats">
                  <div className="dashboardHeroCard">
                    <span>Today Intake</span>
                    <strong>{formatNumber(todayReceiptNet)} QTL</strong>
                    <small>{todayReceipts.length} receipts | {todayReceiptBags} bags</small>
                  </div>
                  <div className="dashboardHeroCard">
                    <span>Intake Coverage</span>
                    <strong>{formatNumber(intakeCoveragePct)}%</strong>
                    <small>
                      {formatNumber(dashboardMetrics.netReceived)} of {formatNumber(dashboardMetrics.expectedYield)} QTL
                    </small>
                  </div>
                  <div className="dashboardHeroCard">
                    <span>Pending Payable</span>
                    <strong>{formatNumber(voucherBalanceOutstanding)} INR</strong>
                    <small>{draftVoucherCount} draft | {paidVoucherCount} paid/overpaid</small>
                  </div>
                  <div className="dashboardHeroCard">
                    <span>Discrepancy Exposure</span>
                    <strong>{formatNumber(discrepancyExposurePct)}%</strong>
                    <small>{dashboardMetrics.discrepancyCount} open cases</small>
                  </div>
                  <div className="dashboardHeroCard">
                    <span>Organizer Net Payable</span>
                    <strong>{formatNumber(organizerNetPayableTotal)} INR</strong>
                    <small>{organizerPendingCount} organizer(s) pending settlement</small>
                  </div>
                </div>
              </section>

              <section className="panel dashboardProgressGrid">
                <article className="dashboardProgressCard">
                  <div className="panelHeader">
                    <h3>Intake Progress</h3>
                    <span>{formatNumber(intakeCoveragePct)}% complete</span>
                  </div>
                  <div className="dashboardProgress">
                    <div
                      className="dashboardProgressBar"
                      style={{ width: `${Math.max(Math.min(intakeCoveragePct, 100), 0)}%` }}
                    />
                  </div>
                  <div className="dashboardProgressMeta">
                    <span>Received {formatNumber(dashboardMetrics.netReceived)} QTL</span>
                    <span>Pending {formatNumber(dashboardMetrics.pending)} QTL</span>
                  </div>
                </article>
                <article className="dashboardProgressCard">
                  <div className="panelHeader">
                    <h3>Lot Utilization</h3>
                    <span>{formatNumber(lotUtilizationPct)}% utilized</span>
                  </div>
                  <div className="dashboardProgress">
                    <div
                      className="dashboardProgressBar dashboardProgressBarWarning"
                      style={{ width: `${Math.max(Math.min(lotUtilizationPct, 100), 0)}%` }}
                    />
                  </div>
                  <div className="dashboardProgressMeta">
                    <span>{dashboardMetrics.totalLots} lots in use</span>
                    <span>{dashboardMetrics.fullLots} full lots</span>
                  </div>
                </article>
                <article className="dashboardProgressCard">
                  <div className="panelHeader">
                    <h3>Voucher Settlement</h3>
                    <span>{financialVouchers.length} vouchers</span>
                  </div>
                  <div className="dashboardProgress">
                    <div
                      className="dashboardProgressBar dashboardProgressBarFinance"
                      style={{
                        width: `${
                          financialVouchers.length > 0
                            ? Math.max(
                                Math.min((paidVoucherCount / financialVouchers.length) * 100, 100),
                                0
                              )
                            : 0
                        }%`
                      }}
                    />
                  </div>
                  <div className="dashboardProgressMeta">
                    <span>{draftVoucherCount} draft / unpaid working vouchers</span>
                    <span>{paidVoucherCount} paid or overpaid</span>
                  </div>
                </article>
                <article className="dashboardProgressCard">
                  <div className="panelHeader">
                    <h3>Organizer Settlement</h3>
                    <span>{organizerCommissionRows.length} organizers</span>
                  </div>
                  <div className="dashboardProgress">
                    <div
                      className="dashboardProgressBar"
                      style={{
                        width: `${
                          organizerNetPayableTotal > 0
                            ? Math.max(
                                Math.min((organizerPaidTotal / organizerNetPayableTotal) * 100, 100),
                                0
                              )
                            : 0
                        }%`
                      }}
                    />
                  </div>
                  <div className="dashboardProgressMeta">
                    <span>{formatNumber(organizerDeductionTotal)} INR deducted</span>
                    <span>{formatNumber(organizerBalanceTotal)} INR balance</span>
                  </div>
                </article>
              </section>

              <section className="panel metricsPanel dashboardMetricsPanel">
                <div className="metricBox">
                  <span>Expected Yield</span>
                  <strong>{formatNumber(dashboardMetrics.expectedYield)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Total Gross Weight</span>
                  <strong>{formatNumber(dashboardMetrics.grossReceived)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Total Net Weight</span>
                  <strong>{formatNumber(dashboardMetrics.netReceived)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Average Purchase Rate</span>
                  <strong>{formatNumber(averagePurchaseRate)} INR/QTL</strong>
                  <small>
                    Gross farmer payable + organizer commission / net intake qty
                  </small>
                </div>
                <div className="metricBox">
                  <span>Pending Balance</span>
                  <strong>{formatNumber(dashboardMetrics.pending)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Total Intake Bags</span>
                  <strong>{dashboardMetrics.intakeBags}</strong>
                </div>
                <div className="metricBox">
                  <span>Avg. Net / Receipt</span>
                  <strong>{formatNumber(averageNetPerReceipt)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Avg. Bag Weight</span>
                  <strong>{formatNumber(averageWeightPerBagKg)} KG</strong>
                </div>
                <div className="metricBox">
                  <span>Open Lots</span>
                  <strong>{dashboardMetrics.openLots}</strong>
                </div>
                <div className="metricBox">
                  <span>Full Lots</span>
                  <strong>{dashboardMetrics.fullLots}</strong>
                </div>
                <div className="metricBox">
                  <span>Active Registrations</span>
                  <strong>{dashboardMetrics.activeRegistrations}</strong>
                </div>
                <div className="metricBox">
                  <span>Open Discrepancy Bags</span>
                  <strong>{dashboardMetrics.discrepancyBags}</strong>
                </div>
                <div className="metricBox">
                  <span>Shifted Qty</span>
                  <strong>{formatNumber(dashboardMetrics.shiftedQty)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Organizer Linked Farmers</span>
                  <strong>{organizerLinkedFarmerCount}</strong>
                </div>
                <div className="metricBox">
                  <span>Organizer Expected Yield</span>
                  <strong>{formatNumber(organizerExpectedYieldTotal)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Organizer Deposited Qty</span>
                  <strong>{formatNumber(organizerDepositedTotal)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Organizer Pending Qty</span>
                  <strong>{formatNumber(organizerPendingYieldTotal)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Gross Organizer Commission</span>
                  <strong>{formatNumber(organizerGrossCommission)} INR</strong>
                </div>
                <div className="metricBox">
                  <span>Organizer Deduction</span>
                  <strong>{formatNumber(organizerDeductionTotal)} INR</strong>
                </div>
                <div className="metricBox">
                  <span>Organizer Balance</span>
                  <strong>{formatNumber(organizerBalanceTotal)} INR</strong>
                </div>
                <div className="metricBox">
                  <span>Organizer Farmers Not Yet Intake</span>
                  <strong>{organizerNoIntakeFarmerCount}</strong>
                </div>
                <div className="metricBox">
                  <span>No-Intake Pending Qty</span>
                  <strong>{formatNumber(organizerNoIntakePendingQty)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Organizer Farmer Payment Pending</span>
                  <strong>{formatNumber(organizerFarmerPaymentPendingTotal)} INR</strong>
                  <small>
                    Net {formatNumber(organizerFarmerPaymentNetTotal)} - Paid{" "}
                    {formatNumber(organizerFarmerPaymentPaidTotal)}
                  </small>
                </div>
              </section>

              <section className="panel twoColumn">
                <article className="infoCard">
                  <div className="panelHeader">
                    <h3>Quick Actions</h3>
                    <span>Open work modules directly</span>
                  </div>
                  <div className="dashboardActionGrid">
                    <button className="secondaryButton" type="button" onClick={() => setActiveView("intake")}>
                      New Intake Entry
                    </button>
                    <button className="secondaryButton" type="button" onClick={() => setActiveView("registrations")}>
                      Registration Master
                    </button>
                    <button className="secondaryButton" type="button" onClick={() => setActiveView("lots")}>
                      Lot Tracking
                    </button>
                    <button className="secondaryButton" type="button" onClick={() => setActiveView("finance")}>
                      Voucher Register
                    </button>
                    <button className="secondaryButton" type="button" onClick={() => setActiveView("commission")}>
                      Organizer Commission
                    </button>
                    {features.discrepancyWorkflow ? (
                      <button
                        className="secondaryButton"
                        type="button"
                        onClick={() => setActiveView("discrepancies")}
                      >
                        Discrepancy Register
                      </button>
                    ) : null}
                    <button className="secondaryButton" type="button" onClick={() => setActiveView("reports")}>
                      Reports
                    </button>
                  </div>
                </article>
                <article className="infoCard">
                  <div className="panelHeader">
                    <h3>Watchlist</h3>
                    <span>Immediate points needing attention</span>
                  </div>
                  <ul>
                    <li>{dashboardMetrics.discrepancyCount} discrepancy cases remain open across {stackHotspotRows.length} stack hotspot groups</li>
                    <li>{pendingRegistrationRows.length} major registrations still have substantial balance pending intake</li>
                    <li>{unpaidVoucherCount} vouchers are not fully settled and outstanding balance is {formatNumber(voucherBalanceOutstanding)} INR</li>
                    <li>{organizerFarmerPaymentPendingRows.length} organizers carry {formatNumber(organizerFarmerPaymentPendingTotal)} INR net farmer-payment balance after overpaid adjustment</li>
                    <li>{organizerPendingCount} organizers still have {formatNumber(organizerBalanceTotal)} INR pending after {formatNumber(organizerDeductionTotal)} INR deduction</li>
                    <li>{organizerNoIntakeFarmerCount} organizer-linked farmers still have zero intake against {formatNumber(organizerNoIntakePendingQty)} QTL pending quantity</li>
                    <li>{dashboardMetrics.shiftedCases} discrepancy shift entries have already moved {formatNumber(dashboardMetrics.shiftedQty)} QTL</li>
                    <li>{features.discrepancyWorkflow ? "Over-intake is being tracked through discrepancy workflow." : "Over-intake is fully blocked at intake save level."}</li>
                  </ul>
                </article>
              </section>

              <section className="dashboardBoard">
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.organizers ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Organizer Commission Board</h3>
                      <p>Net payable, deduction, and settlement position organizer-wise</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("organizers")}
                      >
                        {dashboardExpandedSections.organizers ? "Collapse" : "View Full"}
                      </button>
                      <button className="smallButton" type="button" onClick={() => setActiveView("commission")}>
                        Open Commission
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Organizer</th>
                          <th>District</th>
                          <th>Farmers</th>
                          <th>Gross</th>
                          <th>Deduction</th>
                          <th>Net Payable</th>
                          <th>Paid</th>
                          <th>Balance</th>
                          <th>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleOrganizerDashboardRows.length ? (
                          visibleOrganizerDashboardRows.map((row) => (
                            <tr
                              key={row.organizer.id}
                              className="clickableRow"
                              onClick={() => openOrganizerLedger(row.organizer.id)}
                            >
                              <td>{row.organizer.name}</td>
                              <td>{row.organizer.district || "-"}</td>
                              <td>{row.farmerCount}</td>
                              <td>{formatNumber(row.grossCommissionAmount)}</td>
                              <td>{formatNumber(row.deductionAmount)}</td>
                              <td>{formatNumber(row.netPayableAmount)}</td>
                              <td>{formatNumber(row.paidAmount)}</td>
                              <td>{formatNumber(row.balanceAmount)}</td>
                              <td>
                                <span className={`status ${row.balanceAmount <= 0 ? "full" : "active"}`}>
                                  {row.balanceAmount <= 0 ? "SETTLED" : row.paidAmount > 0 ? "PART PAID" : "PENDING"}
                                </span>
                              </td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={9} className="emptyStateCell">
                              No organizer commission record available.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </article>
              </section>

              <section className="dashboardBoard">
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.organizerFarmerPending ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Organizer-wise Farmer Payment Pending</h3>
                      <p>Net farmer-payment balance after paid and overpaid adjustment</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("organizerFarmerPending")}
                      >
                        {dashboardExpandedSections.organizerFarmerPending ? "Collapse" : "View Full"}
                      </button>
                      <button className="smallButton" type="button" onClick={() => setActiveView("reports")}>
                        Open Reports
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Organizer</th>
                          <th>District</th>
                          <th>Farmers</th>
                          <th>Vouchers</th>
                          <th>Net Payable</th>
                          <th>Paid</th>
                          <th>Net Pending</th>
                          <th>Payment Done</th>
                          <th>Overpaid</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleOrganizerFarmerPaymentPendingRows.length ? (
                          visibleOrganizerFarmerPaymentPendingRows.map((row) => (
                            <tr
                              key={row.organizerId}
                              className="clickableRow"
                              onClick={() => setActiveView("reports")}
                            >
                              <td>{row.organizerName}</td>
                              <td>{row.district || "-"}</td>
                              <td>{row.farmerCount}</td>
                              <td>{row.voucherCount}</td>
                              <td>{formatNumber(row.netPayableAmount)} INR</td>
                              <td>{formatNumber(row.paidAmount)} INR</td>
                              <td>{formatNumber(row.pendingAmount)} INR</td>
                              <td>
                                <div className="dashboardStatusCell">
                                  <div className="dashboardStatusBar">
                                    <div
                                      className={`dashboardStatusFill ${
                                        row.paymentCompletionPct >= 100
                                          ? "dashboardStatusFillHigh"
                                          : row.paymentCompletionPct >= 60
                                            ? "dashboardStatusFillMedium"
                                            : "dashboardStatusFillLow"
                                      }`}
                                      style={{
                                        width: `${Math.max(Math.min(row.paymentCompletionPct, 100), 0)}%`
                                      }}
                                    />
                                  </div>
                                  <small>
                                    {formatNumber(row.paymentCompletionPct)}% paid
                                    {row.pendingAmount < 0 ? " / overpaid" : ""}
                                  </small>
                                </div>
                              </td>
                              <td>{row.overpaidCount}</td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={9} className="emptyStateCell">
                              No organizer-wise farmer payment pending found.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </article>
              </section>

              <section className="dashboardBoard">
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.organizerPerformance ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Organizer Wise Farmer Performance</h3>
                      <p>Expected yield against deposited quantity organizer-wise</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("organizerPerformance")}
                      >
                        {dashboardExpandedSections.organizerPerformance ? "Collapse" : "View Full"}
                      </button>
                      <button className="smallButton" type="button" onClick={() => setActiveView("commission")}>
                        Open Commission
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Organizer</th>
                          <th>District</th>
                          <th>Farmers</th>
                          <th>Expected Yield</th>
                          <th>Deposited</th>
                          <th>Pending</th>
                          <th>Status</th>
                          <th>Coverage</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleOrganizerPerformanceRows.length ? (
                          visibleOrganizerPerformanceRows.map((row) => (
                            <tr
                              key={row.organizerId}
                              className="clickableRow"
                              onClick={() => openOrganizerLedger(row.organizerId)}
                            >
                              <td>{row.organizerName}</td>
                              <td>{row.district || "-"}</td>
                              <td>{row.farmerCount}</td>
                              <td>{formatNumber(row.expectedYieldQtl)} QTL</td>
                              <td>{formatNumber(row.depositedQtl)} QTL</td>
                              <td>{formatNumber(row.pendingQtl)} QTL</td>
                              <td>
                                <div className="dashboardStatusCell">
                                  <div className="dashboardStatusBar">
                                    <div
                                      className={`dashboardStatusFill ${
                                        row.coveragePct >= 95
                                          ? "dashboardStatusFillHigh"
                                          : row.coveragePct >= 60
                                            ? "dashboardStatusFillMedium"
                                            : "dashboardStatusFillLow"
                                      }`}
                                      style={{ width: `${Math.max(Math.min(row.coveragePct, 100), 0)}%` }}
                                    />
                                  </div>
                                  <small>
                                    {row.depositedQtl > 0
                                      ? `${formatNumber(row.depositedQtl)} / ${formatNumber(row.expectedYieldQtl)}`
                                      : "No intake yet"}
                                  </small>
                                </div>
                              </td>
                              <td>{formatNumber(row.coveragePct)}%</td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={8} className="emptyStateCell">
                              No organizer performance data is available yet.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </article>
              </section>

              <section className="dashboardBoard">
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.organizerNoIntake ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Organizer Farmers Not Yet Intake</h3>
                      <p>Organizer-wise list of linked farmers whose intake has not started yet</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("organizerNoIntake")}
                      >
                        {dashboardExpandedSections.organizerNoIntake ? "Collapse" : "View Full"}
                      </button>
                      <button className="smallButton" type="button" onClick={() => setActiveView("commission")}>
                        Open Commission
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Organizer</th>
                          <th>District</th>
                          <th>No-Intake Farmers</th>
                          <th>Pending Qty</th>
                          <th>Farmer Preview</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleOrganizerNoIntakeRows.length ? (
                          visibleOrganizerNoIntakeRows.map((row) => (
                            <tr
                              key={row.organizer.id}
                              className="clickableRow"
                              onClick={() => openOrganizerLedger(row.organizer.id)}
                            >
                              <td>{row.organizer.name}</td>
                              <td>{row.organizer.district || "-"}</td>
                              <td>{row.zeroIntakeFarmerCount}</td>
                              <td>{formatNumber(row.zeroIntakePendingQty)} QTL</td>
                              <td>{row.farmerPreview || "-"}</td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={5} className="emptyStateCell">
                              No organizer-linked farmer is pending first intake.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </article>
              </section>

              <section className="dashboardBoard">
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.districts ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>District Performance</h3>
                      <p>Top districts by received quantity</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("districts")}
                      >
                        {dashboardExpandedSections.districts ? "Collapse" : "View Full"}
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>District</th>
                          <th>Regs.</th>
                          <th>Received</th>
                          <th>Coverage</th>
                          <th>Pending</th>
                          <th>Discrepancies</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleDistrictDashboardRows.map((item) => (
                          <tr key={item.district}>
                            <td>{item.district}</td>
                            <td>{item.registrations}</td>
                            <td>{formatNumber(item.receivedNet)} QTL</td>
                            <td>{formatNumber(item.coveragePct)}%</td>
                            <td>{formatNumber(item.pending)} QTL</td>
                            <td>{item.discrepancyCases}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </article>
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.godowns ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Godown Utilization</h3>
                      <p>Storage position by godown</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("godowns")}
                      >
                        {dashboardExpandedSections.godowns ? "Collapse" : "View Full"}
                      </button>
                      <button className="smallButton" type="button" onClick={() => setActiveView("lots")}>
                        Open Lots
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Godown</th>
                          <th>Qty</th>
                          <th>Lots</th>
                          <th>Full Lots</th>
                          <th>Utilization</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleGodownStockRows.map((item) => (
                          <tr key={item.name}>
                            <td>{item.name}</td>
                            <td>{formatNumber(item.qtyQtl)} QTL</td>
                            <td>{item.lots}</td>
                            <td>{item.fullLots}</td>
                            <td>{formatNumber(item.utilizationPct)}%</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </article>
              </section>

              <section className="dashboardBoard">
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.pendingRegistrations ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Pending Registrations</h3>
                      <p>Highest remaining balance to intake</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("pendingRegistrations")}
                      >
                        {dashboardExpandedSections.pendingRegistrations ? "Collapse" : "View Full"}
                      </button>
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => setActiveView("registrations")}
                      >
                        Open Register
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Reg. Code</th>
                          <th>Farmer</th>
                          <th>District</th>
                          <th>Received</th>
                          <th>Pending</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visiblePendingRegistrationRows.map((item) => (
                          <tr
                            key={item.id}
                            className="clickableRow"
                            onClick={() => openRegistrationFromDashboard(item.id)}
                          >
                            <td>{item.cropRegistrationCode}</td>
                            <td>{item.farmerName}</td>
                            <td>{item.district}</td>
                            <td>{formatNumber(item.totalReceivedQtl)} QTL</td>
                            <td>{formatNumber(item.balanceQtl)} QTL</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </article>
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.stackHotspots ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Stack Hotspots</h3>
                      <p>Most affected discrepancy stacks</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("stackHotspots")}
                      >
                        {dashboardExpandedSections.stackHotspots ? "Collapse" : "View Full"}
                      </button>
                      {features.discrepancyWorkflow ? (
                        <button
                          className="smallButton"
                          type="button"
                          onClick={() => setActiveView("discrepancies")}
                        >
                          Open Register
                        </button>
                      ) : null}
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Godown</th>
                          <th>Stack</th>
                          <th>Cases</th>
                          <th>Bags</th>
                          <th>Qty</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleStackHotspotRows.length ? (
                          visibleStackHotspotRows.map((item) => (
                            <tr
                              key={item.key}
                              className="clickableRow"
                              onClick={() =>
                                openStackHotspotFromDashboard(item.godownName, item.stackNo)
                              }
                            >
                              <td>{item.godownName}</td>
                              <td>{item.stackNo}</td>
                              <td>{item.cases}</td>
                              <td>{item.bags}</td>
                              <td>{formatNumber(item.qtyQtl)} QTL</td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={5} className="emptyStateCell">
                              No discrepancy hotspot currently open.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </article>
              </section>

              <section className="dashboardBoard">
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.recentReceipts ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Recent Receipts</h3>
                      <p>Latest intake activity</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("recentReceipts")}
                      >
                        {dashboardExpandedSections.recentReceipts ? "Collapse" : "View Full"}
                      </button>
                      <button className="smallButton" type="button" onClick={() => setActiveView("intakeEdit")}>
                        Open Receipt Edit
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Receipt</th>
                          <th>Date</th>
                          <th>Reg. Code</th>
                          <th>Bags</th>
                          <th>Net</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleRecentReceiptRows.map((receipt) => (
                          <tr
                            key={receipt.id}
                            className="clickableRow"
                            onClick={() => openReceiptFromDashboard(receipt.id)}
                          >
                            <td>{receipt.receiptNo}</td>
                            <td>{receipt.receiptDate}</td>
                            <td>{receipt.cropRegistrationCode}</td>
                            <td>
                              {receipt.lines.reduce((sum, line) => sum + Number(line.noOfBags ?? 0), 0)}
                            </td>
                            <td>{formatNumber(sumReceiptNetQty(receipt))} QTL</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </article>
                <article
                  className={`panel infoCard dashboardCard ${
                    dashboardExpandedSections.recentVouchers ? "dashboardCardExpanded" : ""
                  }`}
                >
                  <div className="panelHeader dashboardCardHeader">
                    <div>
                      <h3>Recent Vouchers</h3>
                      <p>Latest farmer payment records</p>
                    </div>
                    <div className="panelHeaderActions">
                      <button
                        className="smallButton"
                        type="button"
                        onClick={() => toggleDashboardSection("recentVouchers")}
                      >
                        {dashboardExpandedSections.recentVouchers ? "Collapse" : "View Full"}
                      </button>
                      <button className="smallButton" type="button" onClick={() => setActiveView("finance")}>
                        Open Finance
                      </button>
                    </div>
                  </div>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Voucher</th>
                          <th>Date</th>
                          <th>Farmer</th>
                          <th>Status</th>
                          <th>Balance</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleRecentVoucherRows.length ? (
                          visibleRecentVoucherRows.map((voucher) => (
                            <tr
                              key={voucher.id}
                              className="clickableRow"
                              onClick={() => openVoucherFromDashboard(voucher.id)}
                            >
                              <td>{voucher.voucherNo}</td>
                              <td>{voucher.voucherDate}</td>
                              <td>{voucher.farmerName}</td>
                              <td>
                                <span className={`status ${voucher.status.toLowerCase()}`}>{voucher.status}</span>
                              </td>
                              <td>{formatNumber(voucher.balanceAmount)} INR</td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={5} className="emptyStateCell">
                              No financial vouchers generated yet.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </article>
              </section>

              <section className="panel twoColumn">
                <article className="infoCard">
                  <div className="panelHeader">
                    <h3>Control Notes</h3>
                    <span>System control assumptions</span>
                  </div>
                  <ul>
                    <li>{features.discrepancyWorkflow ? "Over-intake is saved and routed into discrepancy control." : "Over-intake is hard blocked during intake save."}</li>
                    <li>Lot cap remains fixed at 200 QTL per lot for continuity and reporting control.</li>
                    <li>Stack-wise segregation remains active for traceability and stack card reporting.</li>
                    <li>Certification-facing stock should exclude unresolved discrepancy quantity from clean stock view.</li>
                  </ul>
                </article>
                <article className="infoCard">
                  <div className="panelHeader">
                    <h3>System Coverage</h3>
                    <span>Modules contributing to dashboard data</span>
                  </div>
                  <ul>
                    <li>Farmer master import and registration launch workflow</li>
                    <li>Intake entry, edit, delete, and auto lot creation</li>
                    <li>Discrepancy register with shift history and hotspot traceability</li>
                    <li>Financial vouchers, payment ledger, reports, and stack/lot visibility</li>
                  </ul>
                </article>
              </section>
            </div>
          )}

          {activeView === "masters" && (
            <div className="contentStack">
              <section className="panel mastersPanel">
                <div className="mastersHero">
                  <div>
                    <p className="eyebrow">Admin Setup</p>
                    <h3>Masters</h3>
                    <p className="mastersIntro">
                      Create and review godowns and stacks from one clean workspace.
                    </p>
                  </div>
                  <div className="mastersStats">
                    <article className="mastersStatCard">
                      <span>Godowns</span>
                      <strong>{godowns.length}</strong>
                    </article>
                    <article className="mastersStatCard">
                      <span>Stacks</span>
                      <strong>{stacks.length}</strong>
                    </article>
                    <article className="mastersStatCard">
                      <span>Access</span>
                      <strong>{isAdminUser ? "Admin" : "Read only"}</strong>
                    </article>
                  </div>
                </div>

                <div className="mastersWorkspace">
                  <section className="mastersActionGrid">
                    <article className="mastersActionCard">
                      <div className="mastersCardHeader">
                        <div>
                          <h4>Add Godown</h4>
                          <p>Create a new storage location before assigning stacks.</p>
                        </div>
                        <span className="mastersBadge">Step 1</span>
                      </div>
                      <label>
                        <span>Godown name</span>
                        <input
                          onChange={(event) => setNewGodownName(event.target.value)}
                          placeholder="Add godown name"
                          value={newGodownName}
                          disabled={!isAdminUser}
                        />
                      </label>
                      <button
                        className="secondaryButton"
                        onClick={addGodown}
                        type="button"
                        disabled={!isAdminUser}
                      >
                        Add Godown
                      </button>
                    </article>

                    <article className="mastersActionCard">
                      <div className="mastersCardHeader">
                        <div>
                          <h4>Add Stack</h4>
                          <p>Attach a new stack under an existing godown.</p>
                        </div>
                        <span className="mastersBadge">Step 2</span>
                      </div>
                      <label>
                        <span>Stack under</span>
                        <select
                          onChange={(event) => setNewStackGodownId(event.target.value)}
                          value={newStackGodownId}
                          disabled={!isAdminUser}
                        >
                          {godowns.map((item) => (
                            <option key={item.id} value={item.id}>
                              {item.name}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label>
                        <span>Stack no.</span>
                        <input
                          onChange={(event) => setNewStackNo(event.target.value)}
                          placeholder="A-03"
                          value={newStackNo}
                          disabled={!isAdminUser}
                        />
                      </label>
                      <button
                        className="secondaryButton"
                        onClick={addStack}
                        type="button"
                        disabled={!isAdminUser}
                      >
                        Add Stack
                      </button>
                    </article>
                  </section>

                  <section className="mastersRegisterGrid">
                    <article className="infoCard mastersRegisterCard">
                      <div className="panelHeader">
                        <h3>Godown Directory</h3>
                        <span>{godowns.length} godown(s)</span>
                      </div>
                      <div className="tableWrap">
                        <table className="compactTable">
                          <thead>
                            <tr>
                              <th>S.No.</th>
                              <th>Godown Name</th>
                              <th>Linked Stacks</th>
                              <th>Total</th>
                            </tr>
                          </thead>
                          <tbody>
                            {godowns.map((godown, index) => {
                              const linkedStacks = stacks.filter((stack) => stack.godownId === godown.id);
                              return (
                                <tr key={godown.id}>
                                  <td>{index + 1}</td>
                                  <td>{godown.name}</td>
                                  <td>{linkedStacks.map((stack) => stack.stackNo).join(", ") || "-"}</td>
                                  <td>{linkedStacks.length}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </article>

                    <article className="infoCard mastersRegisterCard">
                      <div className="panelHeader">
                        <h3>Stack Register</h3>
                        <span>{stacks.length} stack(s)</span>
                      </div>
                      <div className="tableWrap">
                        <table className="registrationTable compactTable">
                          <thead>
                            <tr>
                              <th>S.No.</th>
                              <th>Stack No.</th>
                              <th>Godown</th>
                            </tr>
                          </thead>
                          <tbody>
                            {stacks.map((stack, index) => (
                              <tr key={stack.id}>
                                <td>{index + 1}</td>
                                <td>{stack.stackNo}</td>
                                <td>{godowns.find((item) => item.id === stack.godownId)?.name ?? "-"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </article>
                  </section>
                </div>
              </section>
            </div>
          )}

          {activeView === "import" && (
            <div className="contentStack">
              <section className="panel twoColumn">
                <article className="infoCard">
                  <h3>Farmer Master Excel</h3>
                  <p>{importMessage}</p>
                  <label>
                    <span>Admin Import Password</span>
                    <input
                      type="password"
                      value={importAdminPassword}
                      onChange={(event) => setImportAdminPassword(event.target.value)}
                      disabled={!effectivePermissions?.canImport}
                    />
                  </label>
                  <input
                    accept=".xlsx,.xls"
                    disabled={!effectivePermissions?.canImport}
                    onChange={(event) => {
                      const file = event.target.files?.[0];
                      if (file) {
                        void handleRegistrationImport(file);
                      }
                    }}
                    type="file"
                  />
                </article>
                <article className="infoCard">
                  <h3>Raw Seed Intake Report Template</h3>
                  <p>Optional in v1, but linked so final report structure stays aligned.</p>
                  <input
                    accept=".xlsx,.xls"
                    onChange={(event) => {
                      const file = event.target.files?.[0];
                      if (file) {
                        void handleReportTemplateImport(file);
                      }
                    }}
                    type="file"
                  />
                </article>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Imported Registrations</h3>
                  <div className="panelHeaderActions">
                    <span>{registrations.length} rows</span>
                    <select
                      onChange={(event) => {
                        setImportSortBy(event.target.value as ImportSortKey);
                        setImportPage(1);
                      }}
                      value={importSortBy}
                    >
                      <option value="farmerName">Sort by Farmer Name</option>
                      <option value="village">Sort by Village</option>
                      <option value="classStage">Sort by Seed Class</option>
                    </select>
                    <button className="secondaryButton" onClick={exportImportViewToExcel} type="button">
                      Download Excel
                    </button>
                  </div>
                </div>
                <div className="tableWrap">
                  <table className="registrationTable">
                    <thead>
                      <tr>
                        <th>Reg. Code</th>
                        <th>Farmer</th>
                        <th>Village</th>
                        <th>District</th>
                        <th>Crop</th>
                        <th>Variety</th>
                        <th>Seed Class</th>
                        <th>Certified Area</th>
                        <th>Expected Yield</th>
                        <th>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {paginatedImportRegistrations.map((item) => (
                        <tr key={item.id}>
                          <td>{item.cropRegistrationCode}</td>
                          <td>{item.farmerName}</td>
                          <td>{item.village}</td>
                          <td>{item.district}</td>
                          <td>{item.crop}</td>
                          <td>{item.variety}</td>
                          <td>{item.classStage}</td>
                          <td>{formatNumber(item.certifiedAreaHa)} Ha</td>
                          <td>{formatNumber(item.expectedYieldQtl)} QTL</td>
                          <td>
                            <span className={`status ${item.status.toLowerCase()}`}>{item.status}</span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="paginationBar">
                  <button
                    className="secondaryButton"
                    disabled={safeImportPage === 1}
                    onClick={() => setImportPage((current) => Math.max(1, current - 1))}
                    type="button"
                  >
                    Previous
                  </button>
                  <span>
                    Page {safeImportPage} of {importPageCount}
                  </span>
                  <button
                    className="secondaryButton"
                    disabled={safeImportPage === importPageCount}
                    onClick={() => setImportPage((current) => Math.min(importPageCount, current + 1))}
                    type="button"
                  >
                    Next
                  </button>
                </div>
              </section>
            </div>
          )}

          {activeView === "registrations" && (
            <div className="contentStack">
              <section className="panel filtersBar">
                <input
                  onChange={(event) => setRegistrationSearch(event.target.value)}
                  placeholder="Search by registration code, farmer, or village"
                  value={registrationSearch}
                />
                <select onChange={(event) => setStatusFilter(event.target.value)} value={statusFilter}>
                  <option value="ALL">All Statuses</option>
                  <option value="ACTIVE">Active</option>
                  <option value="EXHAUSTED">Exhausted</option>
                </select>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Registration Master</h3>
                  <div className="panelHeaderActions">
                    <span>{sortedRegistrationRows.length} rows</span>
                    <button className="secondaryButton" onClick={exportRegistrationMasterToExcel} type="button">
                      Download Excel
                    </button>
                  </div>
                </div>
                <div className="tableWrap">
                  <table>
                    <thead>
                      <tr>
                        <th>S.No.</th>
                        <th>
                          <button className="sortButton" onClick={() => toggleRegistrationSort("cropRegistrationCode")} type="button">
                            Reg. Code
                          </button>
                        </th>
                        <th>
                          <button className="sortButton" onClick={() => toggleRegistrationSort("farmerName")} type="button">
                            Farmer
                          </button>
                        </th>
                        <th>
                          <button className="sortButton" onClick={() => toggleRegistrationSort("village")} type="button">
                            Village
                          </button>
                        </th>
                        <th>District</th>
                        <th>Seed Class</th>
                        <th>
                          <button className="sortButton" onClick={() => toggleRegistrationSort("expectedYieldQtl")} type="button">
                            Expected
                          </button>
                        </th>
                        <th>
                          <button className="sortButton" onClick={() => toggleRegistrationSort("totalReceivedQtl")} type="button">
                            Received
                          </button>
                        </th>
                        <th>
                          <button className="sortButton" onClick={() => toggleRegistrationSort("balanceQtl")} type="button">
                            Balance
                          </button>
                        </th>
                        <th>
                          <button className="sortButton" onClick={() => toggleRegistrationSort("status")} type="button">
                            Status
                          </button>
                        </th>
                        <th>Organizer</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedRegistrationRows.map((item, index) => (
                        <tr key={item.id}>
                          <td>{index + 1}</td>
                          <td>{item.cropRegistrationCode}</td>
                          <td>{item.farmerName}</td>
                          <td>{item.village}</td>
                          <td>{item.district}</td>
                          <td>{item.classStage}</td>
                          <td>{formatNumber(item.expectedYieldQtl)}</td>
                          <td>{formatNumber(item.totalReceivedQtl)}</td>
                          <td>{formatNumber(item.balanceQtl)}</td>
                          <td>
                            <span className={`status ${item.status.toLowerCase()}`}>{item.status}</span>
                          </td>
                          <td>
                            {item.organizerName ? (
                              <div>
                                <strong>{item.organizerName}</strong>
                                <div>{formatNumber(Number(item.organizerCommissionRatePerQtl ?? 0))}/QTL</div>
                              </div>
                            ) : (
                              <span>-</span>
                            )}
                          </td>
                          <td>
                            {item.status === "ACTIVE" ? (
                              <div className="actionButtons">
                                <button
                                  className="secondaryButton smallButton"
                                  onClick={() => goToIntakeForRegistration(item.id)}
                                  type="button"
                                >
                                  Intake
                                </button>
                                <button
                                  className="secondaryButton smallButton"
                                  onClick={() => setDepositViewRegistrationId(item.id)}
                                  type="button"
                                >
                                  Deposit View
                                </button>
                                <button
                                  className="secondaryButton smallButton"
                                  onClick={() => openOrganizerAssignment(item.id)}
                                  type="button"
                                >
                                  Organizer
                                </button>
                              </div>
                            ) : (
                              <div className="actionButtons">
                                <button
                                  className="secondaryButton smallButton"
                                  onClick={() => setDepositViewRegistrationId(item.id)}
                                  type="button"
                                >
                                  Deposit View
                                </button>
                                <button
                                  className="secondaryButton smallButton"
                                  onClick={() => openOrganizerAssignment(item.id)}
                                  type="button"
                                >
                                  Organizer
                                </button>
                              </div>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            </div>
          )}

          {activeView === "intake" && (
            <div className="contentStack">
              <section className="panel">
                <div className="panelHeader">
                  <h3>Intake Receipt</h3>
                  <span>{editingReceiptNo ? "Edit saved receipt" : "Create new intake receipt"}</span>
                </div>

                <div className="formGrid">
                  <label>
                    <span>Registration</span>
                    <select
                      onChange={(event) => setSelectedRegistrationId(event.target.value)}
                      value={selectedRegistrationId}
                    >
                      <option value="">Select registration</option>
                      {registrations.map((item) => (
                        <option key={item.id} value={item.id}>
                          {item.cropRegistrationCode} - {item.farmerName}
                        </option>
                      ))}
                    </select>
                    </label>
                    <label>
                      <span>Receipt No.</span>
                      <input readOnly value={receiptNo} />
                    </label>
                  <label>
                      <span>Receipt Date</span>
                    <input
                      onChange={(event) => setReceiptDate(event.target.value)}
                      type="date"
                      value={receiptDate ? receiptDate.slice(0, 10) : ""}
                    />
                  </label>
                </div>

                {selectedRegistration && (
                  <div className="selectedCard">
                    <strong>{selectedRegistration.cropRegistrationCode}</strong>
                    <span>{selectedRegistration.farmerName}</span>
                    <span>
                      Balance {formatNumber(availableBalanceForSave)} QTL of{" "}
                      {formatNumber(selectedRegistration.allowedIntakeQtl)} QTL
                    </span>
                    <span>
                      Organizer: {selectedRegistration.organizerName || "Not linked"}{" "}
                      {selectedRegistration.organizerName
                        ? `(${formatNumber(Number(selectedRegistration.organizerCommissionRatePerQtl ?? 0))}/QTL)`
                        : ""}
                    </span>
                    {features.discrepancyWorkflow ? (
                      <span>
                        Over-intake will be saved as discrepancy and the working stack will be marked
                        for later excess shift.
                      </span>
                    ) : null}
                  </div>
                )}
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Intake Receipt</h3>
                  <button
                    className="secondaryButton"
                    disabled={!effectivePermissions?.canEntry}
                    onClick={addDraftLine}
                    type="button"
                  >
                    Add Line
                  </button>
                </div>

                <div className="linesStack">
                  {draftLines.map((line, index) => {
                    return (
                      <article className="lineCard" key={line.id}>
                        <div className="lineCardHeader">
                          <strong>Line {index + 1}</strong>
                          <button
                            className="ghostButton"
                            onClick={() => removeDraftLine(line.id)}
                            type="button"
                          >
                            Remove
                          </button>
                        </div>
                        <div className="formGrid compact">
                          <label>
                            <span>Godown</span>
                            <select
                              onChange={(event) =>
                                updateDraftLine(line.id, "godownId", event.target.value)
                              }
                              value={line.godownId}
                            >
                              {godowns.map((item) => (
                                <option key={item.id} value={item.id}>
                                  {item.name}
                                </option>
                              ))}
                            </select>
                          </label>
                          <label>
                            <span>Stack No.</span>
                            <input
                              onChange={(event) =>
                                updateDraftLine(line.id, "stackNo", event.target.value)
                              }
                              value={line.stackNo}
                            />
                          </label>
                          <label>
                            <span>Gross Weight (QTL)</span>
                            <input
                              onChange={(event) =>
                                updateDraftLine(line.id, "grossWeightQtl", event.target.value)
                              }
                              type="number"
                              value={line.grossWeightQtl || ""}
                            />
                          </label>
                          <label>
                            <span>Bags</span>
                            <input
                              onChange={(event) =>
                                updateDraftLine(line.id, "noOfBags", event.target.value)
                              }
                              type="number"
                              value={line.noOfBags || ""}
                            />
                          </label>
                          <label>
                            <span>Net Weight (QTL)</span>
                            <input readOnly type="number" value={line.netWeightQtl || ""} />
                          </label>
                          <label>
                            <span>Weight per Bag (KGS)</span>
                            <input readOnly type="number" value={line.weightPerBagKg || ""} />
                          </label>
                          <label>
                            <span>Moisture (%)</span>
                            <input
                              onChange={(event) =>
                                updateDraftLine(line.id, "moisturePercent", event.target.value)
                              }
                              type="number"
                              value={line.moisturePercent || ""}
                            />
                          </label>
                          <label>
                            <span>Vehicle No.</span>
                            <input
                              onChange={(event) =>
                                updateDraftLine(line.id, "vehicleNo", event.target.value)
                              }
                              value={line.vehicleNo}
                            />
                          </label>
                          <label className="spanTwo">
                            <span>Remarks</span>
                            <input
                              onChange={(event) =>
                                updateDraftLine(line.id, "remarks", event.target.value)
                              }
                              value={line.remarks}
                            />
                          </label>
                        </div>
                      </article>
                    );
                  })}
                </div>

                <div className="actionsFooter">
                  <button
                    className="primaryButton"
                    disabled={
                      isSavingReceipt ||
                      (editingReceiptNo ? !effectivePermissions?.canEdit : !effectivePermissions?.canEntry)
                    }
                    onClick={saveReceipt}
                    type="button"
                  >
                    {isSavingReceipt
                      ? "Saving..."
                      : editingReceiptNo
                        ? "Update Intake Receipt"
                        : "Save Intake Receipt"}
                  </button>
                </div>
              </section>
            </div>
          )}

          {activeView === "intakeEdit" && (
            <div className="contentStack">
              <section className="panel">
                <div className="panelHeader">
                  <h3>Intake Entry Edit</h3>
                  <span>{receipts.length} saved receipts</span>
                </div>
                <div className="tableWrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Receipt No.</th>
                        <th>Date</th>
                        <th>Reg. Code</th>
                        <th>Farmer</th>
                        <th>Remarks</th>
                        <th>Total Net Weight</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {receipts.map((receipt) => (
                        <tr key={receipt.id}>
                          <td>{receipt.receiptNo}</td>
                          <td>{receipt.receiptDate}</td>
                          <td>{receipt.cropRegistrationCode}</td>
                          <td>{receipt.farmerName}</td>
                          <td>{getReceiptRemarks(receipt)}</td>
                          <td>{formatNumber(receipt.lines.reduce((sum, line) => sum + line.qtyQtl, 0))} QTL</td>
                          <td>
                            <div className="actionButtons">
                              <button
                                className="secondaryButton"
                                disabled={!effectivePermissions?.canEdit}
                                onClick={() => startEditReceipt(receipt)}
                                type="button"
                              >
                                Edit
                              </button>
                              <button
                                className="ghostButton"
                                disabled={!effectivePermissions?.canDelete}
                                onClick={() => deleteReceipt(receipt.id, receipt.receiptNo)}
                                type="button"
                              >
                                Delete
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            </div>
          )}

          {activeView === "reports" && (
            <div className="contentStack">
              <section className="panel">
                <div className="panelHeader">
                  <h3>Report Module</h3>
                  <span>Preview and export all intake reports from one place</span>
                </div>
                {reportType === "DAILY_INTAKE_REGISTER" && (
                  <p className="mastersIntro">
                    Daily Intake Register is generated receipt-line wise, with one row for each saved intake
                    line.
                  </p>
                )}
                {isPaymentRegisterReport && (
                  <p className="mastersIntro">
                    Custom Date Payment Register shows all recorded voucher payments in the selected date range,
                    always sorted by payment date.
                  </p>
                )}
                {isOrganizerFarmerPaymentReport && (
                  <p className="mastersIntro">
                    Organizer Wise Farmer Payment report groups farmer voucher position under the linked organizer.
                    Farmers without organizer linkage are shown under `Direct Farmer`.
                  </p>
                )}
                {isOrganizerPaymentTransactionReport && (
                  <p className="mastersIntro">
                    Organizer Payment Transaction Report shows payment entries organiser-wise with reg code, farmer,
                    village, district, gross amount, deduction, final amount, payment, transaction number, and remark.
                  </p>
                )}
                {isOverpaidFarmerReport && (
                  <p className="mastersIntro">
                    Overpaid Farmer Report highlights vouchers where deduction or payment has pushed the farmer
                    balance below zero.
                  </p>
                )}
                {isReceiptVoucherTraceabilityReport && (
                  <p className="mastersIntro">
                    Receipt to Voucher Traceability shows which saved receipt is linked to which farmer voucher,
                    along with current payment position.
                  </p>
                )}
                {isOrganizerIntakePaymentCommissionReport && (
                  <p className="mastersIntro">
                    Organizer Intake vs Payment vs Commission compares organizer-linked expected yield, deposited
                    intake, farmer payment position, and organizer commission settlement together.
                  </p>
                )}
                {isStackCardRegisterReport && (
                  <p className="mastersIntro">
                    Stack Card Register downloads every stack in one file. Non-adjusted stacks show original
                    sequence, while adjusted stacks show only final adjusted position with `*` marked farmers.
                  </p>
                )}
                {isAdjustedLotFormationReport && (
                  <p className="mastersIntro">
                    Adjusted Lot Formation shows final lot-wise farmer position after stack accommodations. Rows
                    marked with `*` are adjusted reporting rows; original intake, lot and voucher data is not changed.
                  </p>
                )}
                {isAdjustedLotLedgerFarmerWiseReport && (
                  <p className="mastersIntro">
                    Adjusted Lot Ledger Farmer Wise groups final adjusted lot positions by farmer/reg no. Common
                    farmer, warehouse, stack and lot cells are merged in Excel/PDF for cleaner reading.
                  </p>
                )}
                <div className="formGrid">
                  <label>
                    <span>Report Type</span>
                    <select value={reportType} onChange={(event) => setReportType(event.target.value as ReportType)}>
                      {reportTypeOptions.map((item) => (
                        <option key={item.value} value={item.value}>
                          {item.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  {!isPaymentRegisterReport &&
                    !isOrganizerFarmerPaymentReport &&
                    !isOrganizerPaymentTransactionReport &&
                    !isOverpaidFarmerReport &&
                    !isReceiptVoucherTraceabilityReport &&
                    !isOrganizerIntakePaymentCommissionReport && (
                    <label>
                      <span>Season</span>
                      <input value={reportSeasonLabel} onChange={(event) => setReportSeasonLabel(event.target.value)} />
                    </label>
                  )}
                  <label>
                    <span>From Date</span>
                    <input type="date" value={reportFromDate} onChange={(event) => setReportFromDate(event.target.value)} />
                  </label>
                  <label>
                    <span>To Date</span>
                    <input type="date" value={reportToDate} onChange={(event) => setReportToDate(event.target.value)} />
                  </label>
                  <label>
                    <span>District</span>
                    {isPaymentRegisterReport ||
                    isOrganizerFarmerPaymentReport ||
                    isOrganizerPaymentTransactionReport ||
                    isOverpaidFarmerReport ||
                    isReceiptVoucherTraceabilityReport ||
                    isOrganizerIntakePaymentCommissionReport ? (
                      <select
                        value={reportDistrict}
                        onChange={(event) => {
                          setReportDistrict(event.target.value);
                          setReportVillage("");
                          setReportFarmerName("");
                          if (
                            isOrganizerFarmerPaymentReport ||
                            isOrganizerPaymentTransactionReport ||
                            isOrganizerIntakePaymentCommissionReport
                          ) {
                            setReportOrganizerName("");
                          }
                        }}
                      >
                        <option value="">All Districts</option>
                        {paymentRegisterDistrictOptions.map((item) => (
                          <option key={item} value={item}>
                            {item}
                          </option>
                        ))}
                      </select>
                    ) : (
                      <input value={reportDistrict} onChange={(event) => setReportDistrict(event.target.value)} />
                    )}
                  </label>
                  {isOrganizerFarmerPaymentReport ||
                  isOrganizerPaymentTransactionReport ||
                  isOrganizerIntakePaymentCommissionReport ? (
                    <label>
                      <span>Organizer</span>
                      <select value={reportOrganizerName} onChange={(event) => setReportOrganizerName(event.target.value)}>
                        <option value="">All Organizers</option>
                        {organizerReportOptions.map((item) => (
                          <option key={item} value={item}>
                            {item}
                          </option>
                        ))}
                      </select>
                    </label>
                  ) : (
                    <label>
                      <span>Farmer Name</span>
                      {isPaymentRegisterReport || isOverpaidFarmerReport ? (
                      <select value={reportFarmerName} onChange={(event) => setReportFarmerName(event.target.value)}>
                        <option value="">All Farmers</option>
                        {paymentRegisterFarmerOptions.map((item) => (
                          <option key={item} value={item}>
                            {item}
                          </option>
                        ))}
                      </select>
                      ) : (
                        <input value={reportFarmerName} onChange={(event) => setReportFarmerName(event.target.value)} />
                      )}
                    </label>
                  )}
                  {isPaymentRegisterReport ? (
                    <>
                      <label>
                        <span>Village</span>
                        <select
                          value={reportVillage}
                          onChange={(event) => {
                            setReportVillage(event.target.value);
                            setReportFarmerName("");
                          }}
                        >
                          <option value="">All Villages</option>
                          {paymentRegisterVillageOptions.map((item) => (
                            <option key={item} value={item}>
                              {item}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label>
                        <span>Status</span>
                        <select value={reportPaymentStatus} onChange={(event) => setReportPaymentStatus(event.target.value)}>
                          <option value="">All Status</option>
                          {paymentRegisterStatusOptions.map((item) => (
                            <option key={item} value={item}>
                              {item}
                            </option>
                          ))}
                        </select>
                      </label>
                    </>
                  ) : isOrganizerFarmerPaymentReport ||
                    isOrganizerPaymentTransactionReport ||
                    isOverpaidFarmerReport ||
                    isReceiptVoucherTraceabilityReport ||
                    isOrganizerIntakePaymentCommissionReport ? (
                    <>
                      <label>
                        <span>Season</span>
                        <input value={reportSeasonLabel} onChange={(event) => setReportSeasonLabel(event.target.value)} />
                      </label>
                      {(isOverpaidFarmerReport || isReceiptVoucherTraceabilityReport) && (
                        <label>
                          <span>Reg. Code</span>
                          <input
                            value={reportRegistrationCode}
                            onChange={(event) => setReportRegistrationCode(event.target.value)}
                          />
                        </label>
                      )}
                    </>
                  ) : (
                    <>
                      <label>
                        <span>Crop</span>
                        <input value={reportCrop} onChange={(event) => setReportCrop(event.target.value)} />
                      </label>
                      <label>
                        <span>Variety</span>
                        <input value={reportVariety} onChange={(event) => setReportVariety(event.target.value)} />
                      </label>
                      <label>
                        <span>Seed Class</span>
                        <input value={reportClassStage} onChange={(event) => setReportClassStage(event.target.value)} />
                      </label>
                      <label>
                        <span>Godown</span>
                        <select value={reportGodownId} onChange={(event) => setReportGodownId(event.target.value)}>
                          <option value="">All Godowns</option>
                          {godowns.map((item) => (
                            <option key={item.id} value={item.id}>
                              {item.name}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label>
                        <span>Stack No.</span>
                        <input value={reportStackNo} onChange={(event) => setReportStackNo(event.target.value)} />
                      </label>
                      <label>
                        <span>Reg. Code</span>
                        <input value={reportRegistrationCode} onChange={(event) => setReportRegistrationCode(event.target.value)} />
                      </label>
                      <label>
                        <span>Report Mode</span>
                        <select value={reportMode} onChange={(event) => setReportMode(event.target.value as ReportMode)}>
                          <option value="ALL">Accepted + Discrepancy</option>
                          <option value="ACCEPTED_ONLY">Accepted Only</option>
                          <option value="DISCREPANCY_ONLY">Discrepancy Only</option>
                        </select>
                      </label>
                    </>
                  )}
                </div>
                <div className="actionsFooter">
                  <button className="primaryButton" onClick={previewReportModule} type="button">
                    Preview
                  </button>
                  <button className="secondaryButton" onClick={downloadReportWorkbook} type="button">
                    Download Excel
                  </button>
                  <button className="secondaryButton" onClick={downloadReportPdf} type="button">
                    Download PDF
                  </button>
                  <button className="ghostButton" onClick={resetReportFilters} type="button">
                    Reset
                  </button>
                </div>
              </section>

              <section className="panel buttonGrid">
                {reportTypeOptions.map((item) => (
                  <button
                    key={item.value}
                    className={item.value === reportType ? "primaryButton" : "secondaryButton"}
                    onClick={() => setReportType(item.value)}
                    type="button"
                  >
                    {item.label}
                  </button>
                ))}
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>{reportPreview?.title ?? "Report Preview"}</h3>
                  <span>{reportPreview ? `${reportPreview.rows.length} row(s)` : "No preview loaded yet"}</span>
                </div>
                {reportPreview ? (
                  <>
                    <div className="detailGrid">
                      {Object.entries(reportPreview.totals).map(([key, value]) => (
                        <div className="metricBox" key={key}>
                          <span>{key}</span>
                          <strong>{String(value)}</strong>
                        </div>
                      ))}
                    </div>
                    <div className="tableWrap">
                      <table>
                        <thead>
                          <tr>
                            {reportPreview.columns.map((column) => (
                              <th key={column}>{column}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {reportPreview.rows.map((row, index) => (
                            <tr key={`${reportPreview.reportType}-${index}`}>
                              {reportPreview.columns.map((column) => (
                                <td key={`${column}-${index}`}>{String(row[column] ?? "")}</td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </>
                ) : (
                  <p>Choose a report, apply filters if needed, and click preview.</p>
                )}
              </section>
            </div>
          )}

          {activeView === "finance" && (
            <div className="contentStack">
              <section className="panel">
                <div className="panelHeader">
                  <h3>Financial Voucher</h3>
                  <span>Generate overall intake cum payment vouchers with dual-rate discrepancy support</span>
                </div>
                <div className="filtersBar">
                  <input
                    placeholder="Search by reg. code, farmer, village, or district"
                    value={voucherGenerationSearch}
                    onChange={(event) => setVoucherGenerationSearch(event.target.value)}
                  />
                  <select
                    value={voucherGenerationDistrictFilter}
                    onChange={(event) => setVoucherGenerationDistrictFilter(event.target.value)}
                  >
                    <option value="">All Districts</option>
                    {voucherDistrictOptions.map((item) => (
                      <option key={item} value={item}>
                        {item}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="tableWrap">
                  <table className="registrationTable compactTable">
                    <thead>
                      <tr>
                        <th>S.No.</th>
                        <th>Reg. Code</th>
                        <th>Farmer</th>
                        <th>Village</th>
                        <th>District</th>
                        <th>Total Bags</th>
                        <th>Total Net</th>
                        <th>Discrepancy</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {voucherRows.length ? (
                        voucherRows.map((item, index) => {
                          const discrepancyInfo =
                            registrationDiscrepancyMap.get(item.id) ?? { qtyQtl: 0, bags: 0 };
                          const totalBags = registrationBagMap.get(item.id) ?? 0;
                          const existingVoucher = voucherByRegistrationId.get(item.id);
                          return (
                            <tr key={item.id}>
                              <td>{index + 1}</td>
                              <td>{item.cropRegistrationCode}</td>
                              <td>{item.farmerName}</td>
                              <td>{item.village}</td>
                              <td>{item.district}</td>
                              <td>{totalBags}</td>
                              <td>{formatNumber(item.totalReceivedQtl)} QTL</td>
                              <td>{formatNumber(discrepancyInfo.qtyQtl)} QTL</td>
                              <td>
                                {existingVoucher ? (
                                  <div className="inlineActionRow">
                                    <button
                                      className="smallButton"
                                      type="button"
                                      onClick={() => openVoucherModal(item.id)}
                                    >
                                      Edit
                                    </button>
                                    <button
                                      className="smallButton"
                                      type="button"
                                      onClick={() => downloadVoucherPdf(existingVoucher)}
                                    >
                                      Download PDF
                                    </button>
                                  </div>
                                ) : (
                                  <button
                                    className="smallButton"
                                    type="button"
                                    onClick={() => openVoucherModal(item.id)}
                                  >
                                    Generate Voucher
                                  </button>
                                )}
                              </td>
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={9} className="emptyStateCell">
                            No registrations match the current financial-voucher filters.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Voucher Register</h3>
                  <span>{filteredVoucherRegisterRows.length} voucher(s)</span>
                </div>
                <div className="filtersBar">
                  <input
                    placeholder="Search voucher no., reg. code, farmer, village, or district"
                    value={voucherRegisterSearch}
                    onChange={(event) => setVoucherRegisterSearch(event.target.value)}
                  />
                  <select
                    value={voucherRegisterDistrictFilter}
                    onChange={(event) => setVoucherRegisterDistrictFilter(event.target.value)}
                  >
                    <option value="">All Districts</option>
                    {voucherDistrictOptions.map((item) => (
                      <option key={item} value={item}>
                        {item}
                      </option>
                    ))}
                  </select>
                  <select
                    value={voucherRegisterStatusFilter}
                    onChange={(event) => setVoucherRegisterStatusFilter(event.target.value)}
                  >
                    <option value="">All Status</option>
                    {voucherRegisterStatusOptions.map((item) => (
                      <option key={item} value={item}>
                        {item}
                      </option>
                    ))}
                  </select>
                  <select
                    value={voucherRegisterOrganizerFilter}
                    onChange={(event) => setVoucherRegisterOrganizerFilter(event.target.value)}
                  >
                    <option value="">All Organizers</option>
                    {voucherRegisterOrganizerOptions.map((item) => (
                      <option key={item} value={item}>
                        {item}
                      </option>
                    ))}
                  </select>
                  <button
                    className="secondaryButton"
                    type="button"
                    onClick={() => downloadBulkVoucherPdf(filteredVoucherRegisterRows)}
                  >
                    Bulk Download PDF
                  </button>
                  <button
                    className="secondaryButton"
                    type="button"
                    onClick={() => downloadBulkPaymentLedgerPdf(filteredVoucherRegisterRows)}
                  >
                    Bulk Ledger PDF
                  </button>
                  <button
                    className="secondaryButton"
                    type="button"
                    onClick={() => downloadOrganizerWiseBulkPaymentLedgerPdf(filteredVoucherRegisterRows)}
                  >
                    Organizer-wise Ledger PDF
                  </button>
                </div>
                <div className="tableWrap">
                  <table className="registrationTable compactTable">
                    <thead>
                      <tr>
                        <th>Voucher No.</th>
                        <th>Date</th>
                        <th>Reg. Code</th>
                        <th>Farmer</th>
                        <th>Total Bags</th>
                        <th>Net Qty</th>
                        <th>Final Payable</th>
                        <th>Net Paid</th>
                        <th>Balance</th>
                        <th>Discrepancy Qty</th>
                        <th>Status</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredVoucherRegisterRows.map((voucher) => (
                        <tr key={voucher.id}>
                          <td>{voucher.voucherNo}</td>
                          <td>{voucher.voucherDate}</td>
                          <td>{voucher.cropRegistrationCode}</td>
                          <td>{voucher.farmerName}</td>
                          <td>{voucher.totalBags}</td>
                          <td>{formatNumber(voucher.totalNetQtyQtl)} QTL</td>
                          <td>{formatNumber(getVoucherFinalPayable(voucher))}</td>
                          <td>{formatNumber(getVoucherTotalPaid(voucher))}</td>
                          <td>{formatNumber(getVoucherBalance(voucher))}</td>
                          <td>{formatNumber(voucher.discrepancyQtyQtl)} QTL</td>
                          <td>
                            <span
                              className={`status ${
                                voucher.status === "PAID"
                                  ? "full"
                                  : voucher.status === "OVERPAID"
                                    ? "blocked"
                                    : "active"
                              }`}
                            >
                              {voucher.status}
                            </span>
                          </td>
                          <td>
                            <div className="inlineActionRow">
                              <button
                                className="smallButton"
                                type="button"
                                onClick={() => openVoucherModal(voucher.cropRegistrationId)}
                              >
                                Edit
                              </button>
                              <button
                                className="smallButton"
                                type="button"
                                onClick={() => {
                                  void deleteVoucher(voucher).catch((error) => {
                                    notifyUser(
                                      error instanceof Error
                                        ? error.message
                                        : "Unable to delete financial voucher."
                                    );
                                  });
                                }}
                              >
                                Delete
                              </button>
                              <button
                                className="smallButton"
                                type="button"
                                onClick={() => openPaymentLedger(voucher)}
                              >
                                Ledger View
                              </button>
                              <button
                                className="smallButton"
                                type="button"
                                onClick={() => downloadVoucherPdf(voucher)}
                              >
                                Download PDF
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            </div>
          )}

          {activeView === "commission" && (
            <div className="contentStack">
              <section className="panel twoColumn">
                <article className="infoCard">
                  <div className="panelHeader">
                    <h3>{editingOrganizerId ? "Edit Organizer" : "Organizer Master"}</h3>
                    <span>One farmer can be linked with one organizer only</span>
                  </div>
                  <div className="formGrid">
                    <label>
                      <span>Organizer Name</span>
                      <input
                        value={newOrganizerName}
                        onChange={(event) => setNewOrganizerName(event.target.value)}
                        placeholder="Organizer name"
                      />
                    </label>
                    <label>
                      <span>Mobile</span>
                      <input
                        value={newOrganizerMobile}
                        onChange={(event) => setNewOrganizerMobile(event.target.value)}
                        placeholder="Mobile no."
                      />
                    </label>
                    <label>
                      <span>Village</span>
                      <input
                        value={newOrganizerVillage}
                        onChange={(event) => setNewOrganizerVillage(event.target.value)}
                        placeholder="Village"
                      />
                    </label>
                    <label>
                      <span>District</span>
                      <input
                        value={newOrganizerDistrict}
                        onChange={(event) => setNewOrganizerDistrict(event.target.value)}
                        placeholder="District"
                      />
                    </label>
                    <label>
                      <span>Commission Rate / QTL</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        value={newOrganizerRate}
                        onChange={(event) => setNewOrganizerRate(event.target.value)}
                        placeholder="0.00"
                      />
                    </label>
                    <label>
                      <span>Deduction Amount</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        value={newOrganizerDeduction}
                        onChange={(event) => setNewOrganizerDeduction(event.target.value)}
                        placeholder="0.00"
                      />
                    </label>
                    <label className="checkboxRow">
                      <input
                        type="checkbox"
                        checked={newOrganizerActive}
                        onChange={(event) => setNewOrganizerActive(event.target.checked)}
                      />
                      <span>Active Organizer</span>
                    </label>
                  </div>
                  <div className="actionsFooter">
                    <button className="primaryButton" onClick={saveOrganizer} type="button">
                      {editingOrganizerId ? "Update Organizer" : "Create Organizer"}
                    </button>
                    <button className="ghostButton" onClick={resetOrganizerForm} type="button">
                      Reset
                    </button>
                  </div>
                </article>

                <article className="infoCard">
                  <div className="panelHeader">
                    <h3>Assignment Control</h3>
                    <span>Registration-wise organizer mapping</span>
                  </div>
                  {organizerAssignmentRegistration ? (
                    <>
                      <div className="selectedCard">
                        <strong>{organizerAssignmentRegistration.cropRegistrationCode}</strong>
                        <span>{organizerAssignmentRegistration.farmerName}</span>
                        <span>
                          Current Organizer: {organizerAssignmentRegistration.organizerName || "Not linked"}
                        </span>
                      </div>
                      <div className="formGrid">
                        <label className="spanTwo">
                          <span>Select Organizer</span>
                          <select
                            value={organizerAssignmentOrganizerId}
                            onChange={(event) => setOrganizerAssignmentOrganizerId(event.target.value)}
                          >
                            <option value="">No Organizer Linked</option>
                            {organizers
                              .filter((item) => item.isActive)
                              .map((item) => (
                                <option key={item.id} value={item.id}>
                                  {item.name} ({formatNumber(item.commissionRatePerQtl)}/QTL)
                                </option>
                              ))}
                          </select>
                        </label>
                      </div>
                      <div className="actionsFooter">
                        <button className="primaryButton" onClick={saveOrganizerAssignment} type="button">
                          Save Mapping
                        </button>
                        <button className="ghostButton" onClick={closeOrganizerAssignment} type="button">
                          Close
                        </button>
                      </div>
                    </>
                  ) : (
                    <p>Select any farmer from Registration Master and click `Organizer` to link the farmer.</p>
                  )}
                </article>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Organizer Commission Register</h3>
                  <span>{organizerCommissionRows.length} organizer(s)</span>
                </div>
                <div className="filtersBar">
                  <input
                    placeholder="Search organizer, mobile, village, or district"
                    value={organizerSearch}
                    onChange={(event) => setOrganizerSearch(event.target.value)}
                  />
                  <select
                    value={organizerDistrictFilter}
                    onChange={(event) => setOrganizerDistrictFilter(event.target.value)}
                  >
                    <option value="">All Districts</option>
                    {organizerDistrictOptions.map((item) => (
                      <option key={item} value={item}>
                        {item}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="tableWrap">
                  <table className="registrationTable compactTable">
                    <thead>
                      <tr>
                        <th>Organizer</th>
                        <th>District</th>
                        <th>Farmers</th>
                        <th>Total Intake</th>
                        <th>Rate / QTL</th>
                        <th>Gross Commission</th>
                        <th>Deduction</th>
                        <th>Net Payable</th>
                        <th>Paid</th>
                        <th>Balance</th>
                        <th>Status</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {organizerCommissionRows.length ? (
                        organizerCommissionRows.map((row) => (
                          <tr key={row.organizer.id}>
                            <td>
                              <strong>{row.organizer.name}</strong>
                              <div>{row.organizer.mobile || "-"}</div>
                            </td>
                            <td>{row.organizer.district || "-"}</td>
                            <td>{row.farmerCount}</td>
                            <td>{formatNumber(row.totalIntakeQtl)} QTL</td>
                            <td>{formatNumber(row.ratePerQtl)}</td>
                            <td>{formatNumber(row.grossCommissionAmount)}</td>
                            <td>{formatNumber(row.deductionAmount)}</td>
                            <td>{formatNumber(row.netPayableAmount)}</td>
                            <td>{formatNumber(row.paidAmount)}</td>
                            <td>{formatNumber(row.balanceAmount)}</td>
                            <td>
                              <span className={`status ${row.balanceAmount <= 0 ? "full" : "active"}`}>
                                {row.balanceAmount <= 0 ? "SETTLED" : row.paidAmount > 0 ? "PART PAID" : "PENDING"}
                              </span>
                            </td>
                            <td>
                              <div className="inlineActionRow">
                                <button
                                  className="smallButton"
                                  type="button"
                                  onClick={() => openOrganizerLedger(row.organizer.id)}
                                >
                                  Ledger View
                                </button>
                                <button
                                  className="smallButton"
                                  type="button"
                                  onClick={() => downloadOrganizerCommissionVoucherPdf(row)}
                                >
                                  Download Voucher
                                </button>
                                <button
                                  className="smallButton"
                                  type="button"
                                  onClick={() => openOrganizerForEdit(row.organizer)}
                                >
                                  Edit
                                </button>
                                <button
                                  className="smallButton"
                                  type="button"
                                  onClick={() => deleteOrganizer(row.organizer)}
                                >
                                  Delete
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={12} className="emptyStateCell">
                            No organizer commission records available.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Organizer Master Register</h3>
                  <span>{filteredOrganizerMasterRows.length} record(s)</span>
                </div>
                <div className="tableWrap">
                  <table className="registrationTable compactTable">
                    <thead>
                      <tr>
                        <th>Name</th>
                        <th>Mobile</th>
                        <th>Village</th>
                        <th>District</th>
                        <th>Rate / QTL</th>
                        <th>Deduction</th>
                        <th>Status</th>
                        <th>Linked Farmers</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredOrganizerMasterRows.map((organizer) => {
                        const linkedCount = registrations.filter((item) => item.organizerId === organizer.id).length;
                        return (
                          <tr key={organizer.id}>
                            <td>{organizer.name}</td>
                            <td>{organizer.mobile || "-"}</td>
                            <td>{organizer.village || "-"}</td>
                            <td>{organizer.district || "-"}</td>
                            <td>{formatNumber(organizer.commissionRatePerQtl)}</td>
                            <td>{formatNumber(organizer.deductionAmount ?? 0)}</td>
                            <td>
                              <span className={`status ${organizer.isActive ? "active" : "blocked"}`}>
                                {organizer.isActive ? "ACTIVE" : "INACTIVE"}
                              </span>
                            </td>
                            <td>{linkedCount}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </section>
            </div>
          )}

          {activeView === "slips" && (
            <div className="contentStack">
              <section className="panel">
                <div className="panelHeader">
                  <h3>Slip Print Center</h3>
                  <span>Search, filter, and open the required slip format</span>
                </div>
                <div className="formGrid">
                  <label className="spanTwo">
                    <span>Search</span>
                    <input
                      placeholder="Search by reg. code, farmer, village, district, or class"
                      value={slipSearch}
                      onChange={(event) => setSlipSearch(event.target.value)}
                    />
                  </label>
                  <label>
                    <span>District</span>
                    <select value={slipDistrictFilter} onChange={(event) => setSlipDistrictFilter(event.target.value)}>
                      <option value="">All Districts</option>
                      {slipDistrictOptions.map((item) => (
                        <option key={item} value={item}>
                          {item}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    <span>Village</span>
                    <select value={slipVillageFilter} onChange={(event) => setSlipVillageFilter(event.target.value)}>
                      <option value="">All Villages</option>
                      {slipVillageOptions.map((item) => (
                        <option key={item} value={item}>
                          {item}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    <span>Seed Class</span>
                    <select value={slipClassFilter} onChange={(event) => setSlipClassFilter(event.target.value)}>
                      <option value="">All Classes</option>
                      {slipClassOptions.map((item) => (
                        <option key={item} value={item}>
                          {item}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    <span>Crop</span>
                    <select value={slipCropFilter} onChange={(event) => setSlipCropFilter(event.target.value)}>
                      <option value="">All Crops</option>
                      {slipCropOptions.map((item) => (
                        <option key={item} value={item}>
                          {item}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>
                <div className="filtersBar">
                  <label className="checkboxRow">
                    <input
                      type="checkbox"
                      checked={slipOnlyWithIntake}
                      onChange={(event) => setSlipOnlyWithIntake(event.target.checked)}
                    />
                    <span>Only registrations with intake</span>
                  </label>
                  <button className="secondaryButton" onClick={resetSlipFilters} type="button">
                    Reset Filters
                  </button>
                  <button className="primaryButton" onClick={downloadAllOverallSlips} type="button">
                    Download All Overall Slips
                  </button>
                </div>
                <div className="tableHint">
                  Recommended use: filter by district, village, seed class, crop, and keep `Only registrations with intake` enabled before bulk download.
                </div>
                <div className="filtersBar">
                  <input
                    value={`${slipRegistrationRows.length} registration row(s) after filters`}
                    readOnly
                  />
                </div>
                <div className="tableWrap">
                  <table className="registrationTable compactTable">
                    <thead>
                      <tr>
                        <th>S.No.</th>
                        <th>Reg. Code</th>
                        <th>Farmer</th>
                        <th>Village</th>
                        <th>District</th>
                        <th>Seed Class</th>
                        <th>Expected</th>
                        <th>Received</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {slipRegistrationRows.map((item, index) => (
                        <tr key={item.id}>
                          <td>{index + 1}</td>
                          <td>{item.cropRegistrationCode}</td>
                          <td>{item.farmerName}</td>
                          <td>{item.village}</td>
                          <td>{item.district}</td>
                          <td>{item.classStage}</td>
                          <td>{formatNumber(item.expectedYieldQtl)}</td>
                          <td>{formatNumber(item.totalReceivedQtl)}</td>
                          <td>
                            <div className="actionButtons">
                              <button
                                className="smallButton"
                                type="button"
                                onClick={() => openSlipModal("FARMER_SINGLE_RECEIPT", item.id)}
                              >
                                Single
                              </button>
                              <button
                                className="smallButton"
                                type="button"
                                onClick={() => openSlipModal("FARMER_OVERALL", item.id)}
                              >
                                Overall
                              </button>
                              <button
                                className="smallButton"
                                type="button"
                                onClick={() => openSlipModal("DAILY_CONSOLIDATED", item.id)}
                              >
                                Day Wise
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            </div>
          )}

          {activeView === "discrepancies" && features.discrepancyWorkflow && (
            <div className="contentStack">
              <section className="panel metricsPanel">
                <div className="metricBox">
                  <span>Open Discrepancies</span>
                  <strong>{openDiscrepancies.length}</strong>
                </div>
                <div className="metricBox">
                  <span>Excess Qty Pending</span>
                  <strong>
                    {formatNumber(
                      openDiscrepancies.reduce((sum, item) => sum + item.excessQtyQtl, 0)
                    )}{" "}
                    QTL
                  </strong>
                </div>
                <div className="metricBox">
                  <span>Affected Stacks</span>
                  <strong>
                    {new Set(openDiscrepancies.map((item) => `${item.godownName}-${item.stackNo}`)).size}
                  </strong>
                </div>
                <div className="metricBox">
                  <span>Estimated Excess Bags</span>
                  <strong>{openDiscrepancies.reduce((sum, item) => sum + item.estimatedExcessBags, 0)}</strong>
                </div>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Discrepancy Register</h3>
                  <span>{discrepancies.length} total discrepancy records</span>
                </div>
                {selectedDiscrepancy && (
                  <div className="contentStack">
                    <div className="workflowSwitchRow">
                      <button
                        className={discrepancyWorkflowMode === "accommodation" ? "primaryButton" : "secondaryButton"}
                        onClick={() => setDiscrepancyWorkflowMode("accommodation")}
                        type="button"
                      >
                        In-Stack Accommodation
                      </button>
                      <button
                        className={discrepancyWorkflowMode === "shift" ? "primaryButton" : "secondaryButton"}
                        onClick={() => setDiscrepancyWorkflowMode("shift")}
                        type="button"
                      >
                        Physical Shift
                      </button>
                    </div>
                    {discrepancyWorkflowMode === "shift" ? (
                    <div className="shiftPanel">
                      <div className="panelHeader">
                        <h3>Shift Excess Entry</h3>
                        <span>{selectedDiscrepancy.discrepancyNo}</span>
                      </div>
                      <div className="selectedCard">
                        <strong>{selectedDiscrepancy.cropRegistrationCode}</strong>
                        <span>{selectedDiscrepancy.farmerName}</span>
                        <span>
                          Pending excess {formatNumber(selectedDiscrepancy.excessQtyQtl)} QTL /{" "}
                          {selectedDiscrepancy.estimatedExcessBags} bags
                        </span>
                        <span>
                          Source {selectedDiscrepancy.godownName} / Stack {selectedDiscrepancy.stackNo}
                        </span>
                      </div>
                      <div className="formGrid compact">
                        <label>
                          <span>Target Godown</span>
                          <select
                            onChange={(event) => setShiftTargetGodownId(event.target.value)}
                            value={shiftTargetGodownId}
                          >
                            {godowns.map((item) => (
                              <option key={item.id} value={item.id}>
                                {item.name}
                              </option>
                            ))}
                          </select>
                        </label>
                        <label>
                          <span>Target Non-Cert Stack</span>
                          <input
                            onChange={(event) => setShiftTargetStackNo(event.target.value)}
                            value={shiftTargetStackNo}
                          />
                        </label>
                        <label>
                          <span>Shift Qty (QTL)</span>
                          <input
                            onChange={(event) => setShiftQtyQtl(event.target.value)}
                            type="number"
                            value={shiftQtyQtl}
                          />
                        </label>
                        <label>
                          <span>Shift Bags</span>
                          <input
                            onChange={(event) => setShiftBags(event.target.value)}
                            type="number"
                            value={shiftBags}
                          />
                        </label>
                        <label>
                          <span>Shift Date</span>
                          <input
                            onChange={(event) => setShiftDate(event.target.value)}
                            type="date"
                            value={shiftDate}
                          />
                        </label>
                        <label>
                          <span>Approved By</span>
                          <input
                            onChange={(event) => setShiftApprovedBy(event.target.value)}
                            value={shiftApprovedBy}
                          />
                        </label>
                        <label className="spanTwo">
                          <span>Shift Remarks</span>
                          <input
                            onChange={(event) => setShiftRemarks(event.target.value)}
                            value={shiftRemarks}
                          />
                        </label>
                      </div>
                      <div className="actionsFooter">
                        <button className="primaryButton" onClick={saveDiscrepancyShift} type="button">
                          Save Shift Entry
                        </button>
                      </div>
                    </div>
                    ) : null}
                    {discrepancyWorkflowMode === "accommodation" ? (
                    <div className="shiftPanel">
                      <div className="panelHeader">
                        <h3>In-Stack Accommodation Register</h3>
                        <span>For this module only</span>
                      </div>
                      <div className="selectedCard">
                        <span>
                          Source: {selectedDiscrepancy.cropRegistrationCode} / {selectedDiscrepancy.farmerName}
                        </span>
                        <span>
                          Mapped in module: {formatNumber(selectedDiscrepancyMappedQty)} QTL /{" "}
                          {selectedDiscrepancyMappedBags} bags
                        </span>
                        <span>
                          Remaining for module view: {formatNumber(selectedDiscrepancyRemainingForAccommodation)} QTL /{" "}
                          {selectedDiscrepancyRemainingBagsForAccommodation} bags
                        </span>
                        <span>Same stack only: {selectedDiscrepancy.godownName} / {selectedDiscrepancy.stackNo}</span>
                        <span>This does not change finance, intake, lots, or main stock position.</span>
                      </div>
                      <div className="formGrid compact">
                        <label>
                          <span>Target Farmer In Same Stack</span>
                          <select
                            onChange={(event) => setAccommodationTargetRegistrationId(event.target.value)}
                            value={accommodationTargetRegistrationId}
                          >
                            <option value="">Select target farmer</option>
                            {eligibleAccommodationTargets.map((item) => (
                              <option key={item.registration.id} value={item.registration.id}>
                                {item.registration.cropRegistrationCode} - {item.registration.farmerName}
                              </option>
                            ))}
                          </select>
                        </label>
                        <label>
                          <span>Accommodation Qty (QTL)</span>
                          <input
                            onChange={(event) => setAccommodationQtyQtl(event.target.value)}
                            type="number"
                            value={accommodationQtyQtl}
                          />
                        </label>
                        <label>
                          <span>Accommodation Bags</span>
                          <input
                            onChange={(event) => setAccommodationBags(event.target.value)}
                            type="number"
                            value={accommodationBags}
                          />
                        </label>
                        <label>
                          <span>Accommodation Date</span>
                          <input
                            onChange={(event) => setAccommodationDate(event.target.value)}
                            type="date"
                            value={accommodationDate}
                          />
                        </label>
                        <label className="spanTwo">
                          <span>Accommodation Remarks</span>
                          <input
                            onChange={(event) => setAccommodationRemarks(event.target.value)}
                            value={accommodationRemarks}
                          />
                        </label>
                      </div>
                      <div className="actionsFooter">
                        <button className="primaryButton" onClick={saveStackAccommodation} type="button">
                          {editingAccommodationId ? "Update Accommodation Mapping" : "Save Accommodation Mapping"}
                        </button>
                        {editingAccommodationId ? (
                          <button
                            className="secondaryButton"
                            onClick={() => {
                              setEditingAccommodationId("");
                              setAccommodationTargetRegistrationId("");
                              setAccommodationQtyQtl(String(selectedDiscrepancyRemainingForAccommodation || ""));
                              setAccommodationBags(
                                String(selectedDiscrepancyRemainingBagsForAccommodation || 0)
                              );
                              setAccommodationDate(new Date().toISOString().slice(0, 10));
                              setAccommodationRemarks("");
                            }}
                            type="button"
                          >
                            Cancel Edit
                          </button>
                        ) : null}
                      </div>
                      <div className="tableWrap">
                        <table className="registrationTable compactTable">
                          <thead>
                            <tr>
                              <th>Target Reg. Code</th>
                              <th>Target Farmer</th>
                              <th>Village</th>
                              <th>District</th>
                              <th>Stack Qty</th>
                              <th>Stack Bags</th>
                              <th>Farmer Pending</th>
                            </tr>
                          </thead>
                          <tbody>
                            {eligibleAccommodationTargets.length ? (
                              eligibleAccommodationTargets.map((item) => (
                                <tr
                                  key={item.registration.id}
                                  className={
                                    item.registration.id === accommodationTargetRegistrationId ? "rowActive" : ""
                                  }
                                >
                                  <td>{item.registration.cropRegistrationCode}</td>
                                  <td>{item.registration.farmerName}</td>
                                  <td>{item.registration.village || "-"}</td>
                                  <td>{item.registration.district || "-"}</td>
                                  <td>{formatNumber(item.stackQtyQtl)} QTL</td>
                                  <td>{item.stackBags}</td>
                                  <td>{formatNumber(item.registration.balanceQtl)} QTL</td>
                                </tr>
                              ))
                            ) : (
                              <tr>
                                <td colSpan={7} className="emptyStateCell">
                                  No eligible target farmer found in this same stack.
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                      <div className="tableWrap">
                        <table className="registrationTable compactTable">
                          <thead>
                            <tr>
                              <th>Date</th>
                              <th>Source Reg. Code</th>
                              <th>Source Farmer</th>
                              <th>Target Reg. Code</th>
                              <th>Target Farmer</th>
                              <th>Adjusted Qty</th>
                              <th>Adjusted Bags</th>
                              <th>Remaining After Map</th>
                              <th>Remarks</th>
                              <th>Action</th>
                            </tr>
                          </thead>
                          <tbody>
                            {selectedDiscrepancyAccommodations.length ? (
                              selectedDiscrepancyAccommodations.map((item, index) => {
                                const qtyAfterThisMap = roundQtl(
                                  Number(selectedDiscrepancy.excessQtyQtl ?? 0) -
                                    selectedDiscrepancyAccommodations
                                      .slice(0, index + 1)
                                      .reduce((sum, current) => sum + Number(current.adjustedQtyQtl ?? 0), 0)
                                );
                                return (
                                <tr key={item.id}>
                                  <td>{item.adjustmentDate}</td>
                                  <td>{item.sourceRegistrationCode}</td>
                                  <td>{item.sourceFarmerName}</td>
                                  <td>{item.targetRegistrationCode}</td>
                                  <td>{item.targetFarmerName}</td>
                                  <td>{formatNumber(item.adjustedQtyQtl)} QTL</td>
                                  <td>{item.adjustedBags}</td>
                                  <td>{formatNumber(Math.max(qtyAfterThisMap, 0))} QTL</td>
                                  <td>{item.remarks || "-"}</td>
                                  <td>
                                    <div className="actionButtons">
                                      <button
                                        className="secondaryButton smallButton"
                                        onClick={() => beginEditStackAccommodation(item)}
                                        type="button"
                                      >
                                        Edit
                                      </button>
                                      <button
                                        className="secondaryButton smallButton"
                                        onClick={() => deleteStackAccommodation(item)}
                                        type="button"
                                      >
                                        Delete
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                                );
                              })
                            ) : (
                              <tr>
                                <td colSpan={10} className="emptyStateCell">
                                  No in-stack accommodation mapping saved for this discrepancy yet.
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                      {adjustedStackCardPreview ? (
                        <section className="panel">
                          <div className="panelHeader">
                            <h3>Adjusted Stack Card Preview</h3>
                            <span>
                              {adjustedStackCardPreview.godownName} / {adjustedStackCardPreview.stackNo}
                            </span>
                          </div>
                          <div className="selectedCard">
                            <span>Godown: {adjustedStackCardPreview.godownName}</span>
                            <span>Stack No.: {adjustedStackCardPreview.stackNo}</span>
                            <span>Original Farmer Rows: {adjustedStackCardPreview.originalRows.length}</span>
                            <span>Changed Farmers: {adjustedStackCardPreview.changedFarmerCount}</span>
                            <span>For reference only. Does not alter main stock records.</span>
                          </div>
                          <div className="actionsFooter">
                            <button
                              className="secondaryButton"
                              onClick={downloadAdjustedStackCardExcel}
                              type="button"
                            >
                              Download Excel
                            </button>
                            <button
                              className="secondaryButton"
                              onClick={downloadAdjustedStackCardPdf}
                              type="button"
                            >
                              Download PDF
                            </button>
                          </div>
                          <div className="tableWrap">
                            <table className="registrationTable compactTable">
                              <thead>
                                <tr>
                                  <th colSpan={6}>Original Stack Position</th>
                                </tr>
                                <tr>
                                  <th>S. No.</th>
                                  <th>Reg. Code</th>
                                  <th>Farmer Name</th>
                                  <th>Village</th>
                                  <th>Qty</th>
                                  <th>Bags</th>
                                </tr>
                              </thead>
                              <tbody>
                                {adjustedStackCardPreview.originalRows.map((row, index) => (
                                  <tr key={`original-${row.regCode}`}>
                                    <td>{index + 1}</td>
                                    <td>{row.regCode}</td>
                                    <td>{row.farmerName}</td>
                                    <td>{row.village}</td>
                                    <td>{formatNumber(row.qtyQtl)} QTL</td>
                                    <td>{row.bags}</td>
                                  </tr>
                                ))}
                                <tr className="rowActive">
                                  <td />
                                  <td />
                                  <td />
                                  <td>TOTAL</td>
                                  <td>
                                    {formatNumber(
                                      adjustedStackCardPreview.originalRows.reduce(
                                        (sum, row) => sum + Number(row.qtyQtl ?? 0),
                                        0
                                      )
                                    )}{" "}
                                    QTL
                                  </td>
                                  <td>
                                    {adjustedStackCardPreview.originalRows.reduce(
                                      (sum, row) => sum + Number(row.bags ?? 0),
                                      0
                                    )}
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                          </div>
                          <div className="tableWrap">
                            <table className="registrationTable compactTable">
                              <thead>
                                <tr>
                                  <th colSpan={7}>Final Stack Position After Accommodation</th>
                                </tr>
                                <tr>
                                  <th>S. No.</th>
                                  <th>Reg. Code</th>
                                  <th>Farmer Name</th>
                                  <th>Village</th>
                                  <th>Qty</th>
                                  <th>Bags</th>
                                  <th>Mark</th>
                                </tr>
                              </thead>
                              <tbody>
                                {adjustedStackCardPreview.adjustedRows.map((row, index) => (
                                  <tr
                                    key={`adjusted-${row.regCode}`}
                                    className={row.changed ? "rowActive" : ""}
                                  >
                                    <td>{index + 1}</td>
                                    <td>{row.regCode}</td>
                                    <td>{row.farmerName}</td>
                                    <td>{row.village}</td>
                                    <td>{formatNumber(row.finalQtyQtl)} QTL</td>
                                    <td>{row.finalBags}</td>
                                    <td>{row.changed ? "*" : ""}</td>
                                  </tr>
                                ))}
                                <tr className="rowActive">
                                  <td />
                                  <td />
                                  <td />
                                  <td>TOTAL</td>
                                  <td>
                                    {formatNumber(
                                      adjustedStackCardPreview.adjustedRows.reduce(
                                        (sum, row) => sum + Number(row.finalQtyQtl ?? 0),
                                        0
                                      )
                                    )}{" "}
                                    QTL
                                  </td>
                                  <td>
                                    {adjustedStackCardPreview.adjustedRows.reduce(
                                      (sum, row) => sum + Number(row.finalBags ?? 0),
                                      0
                                    )}
                                  </td>
                                  <td />
                                </tr>
                              </tbody>
                            </table>
                          </div>
                          <div className="selectedCard">
                            <span>Changed Farmers: {adjustedStackCardPreview.changedFarmerCount}</span>
                            <span>
                              Accommodated Qty: {formatNumber(adjustedStackCardPreview.totalAccommodatedQtyQtl)} QTL
                            </span>
                            <span>Accommodated Bags: {adjustedStackCardPreview.totalAccommodatedBags}</span>
                            <span>* Marked rows indicate farmers affected in accommodation.</span>
                          </div>
                        </section>
                      ) : null}
                    </div>
                    ) : null}
                  </div>
                )}
                <div className="tableWrap">
                  <table className="registrationTable compactTable">
                    <thead>
                      <tr>
                        <th>S.No.</th>
                        <th>Godown</th>
                        <th>Stack</th>
                        <th>Reg. Code</th>
                        <th>Farmer</th>
                        <th>Receipt</th>
                        <th>Date</th>
                        <th>Expected</th>
                        <th>Receipt Net</th>
                        <th>Excess</th>
                        <th>Excess Bags</th>
                        <th>Status</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {discrepancies
                        .slice()
                        .sort((left, right) => {
                          if (left.status !== right.status) {
                            return left.status === "RESOLVED" ? 1 : -1;
                          }
                          if (left.receiptDate !== right.receiptDate) {
                            return left.receiptDate.localeCompare(right.receiptDate);
                          }
                          return left.discrepancyNo.localeCompare(right.discrepancyNo);
                        })
                        .map((item, index) => (
                          <tr
                            key={item.id}
                            className={item.id === selectedDiscrepancyId ? "rowActive" : ""}
                            onClick={() => {
                              setSelectedDiscrepancyId(item.id);
                              setDiscrepancyWorkflowMode("accommodation");
                            }}
                          >
                            <td>{index + 1}</td>
                            <td>{item.godownName}</td>
                            <td>{item.stackNo}</td>
                            <td>{item.cropRegistrationCode}</td>
                            <td>{item.farmerName}</td>
                            <td>{item.receiptNo}</td>
                            <td>{item.receiptDate}</td>
                            <td>{formatNumber(item.expectedQtyQtl)} QTL</td>
                            <td>{formatNumber(item.receiptNetQtyQtl)} QTL</td>
                            <td>{formatNumber(item.excessQtyQtl)} QTL</td>
                            <td>{item.estimatedExcessBags}</td>
                            <td>
                              <span className={`status ${item.status === "RESOLVED" ? "active" : "blocked"}`}>
                                {item.status}
                              </span>
                            </td>
                            <td>
                              <div className="actionButtons">
                                <button
                                  className="secondaryButton smallButton"
                                  onClick={() => {
                                    setSelectedDiscrepancyId(item.id);
                                    setDiscrepancyWorkflowMode("accommodation");
                                  }}
                                  type="button"
                                >
                                  Select
                                </button>
                                {item.status !== "RESOLVED" ? (
                                  <button
                                    className="secondaryButton smallButton"
                                    onClick={() => startShiftEntry(item)}
                                    type="button"
                                  >
                                    Shift Excess
                                  </button>
                                ) : null}
                              </div>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </section>
              {discrepancyWorkflowMode === "shift" ? (
                <section className="panel">
                  <div className="panelHeader">
                    <h3>Shift History</h3>
                    <span>{discrepancyShifts.length} saved shift entries</span>
                  </div>
                  <div className="tableWrap">
                    <table>
                      <thead>
                        <tr>
                          <th>Discrepancy No.</th>
                          <th>Reg. Code</th>
                          <th>From Stack</th>
                          <th>To Stack</th>
                          <th>Shift Qty</th>
                          <th>Shift Bags</th>
                          <th>Shift Date</th>
                          <th>Approved By</th>
                          <th>Remarks</th>
                        </tr>
                      </thead>
                      <tbody>
                        {discrepancyShifts.map((item) => (
                          <tr key={item.id}>
                            <td>{item.discrepancyNo}</td>
                            <td>{item.cropRegistrationCode}</td>
                            <td>
                              {item.fromGodownName} / {item.fromStackNo}
                            </td>
                            <td>
                              {item.toGodownName} / {item.toStackNo}
                            </td>
                            <td>{formatNumber(item.shiftedQtyQtl)} QTL</td>
                            <td>{item.shiftedBags}</td>
                            <td>{item.shiftDate}</td>
                            <td>{item.approvedBy || "-"}</td>
                            <td>{item.remarks || "-"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              ) : (
                <section className="panel">
                  <div className="panelHeader">
                    <h3>Pending Accommodation Register</h3>
                    <span>{pendingAccommodationRows.length} discrepancy row(s) still pending adjustment</span>
                  </div>
                  <p className="mastersIntro">
                    This view shows only the discrepancy rows where accommodation is still pending. Completed
                    accommodation mappings stay inside the selected discrepancy card only.
                  </p>
                  <div className="tableWrap">
                    <table className="registrationTable compactTable">
                      <thead>
                        <tr>
                          <th>Discrepancy No.</th>
                          <th>Godown / Stack</th>
                          <th>Reg. Code</th>
                          <th>Farmer</th>
                          <th>Excess Qty</th>
                          <th>Mapped Qty</th>
                          <th>Pending Qty</th>
                          <th>Pending Bags</th>
                          <th>Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {pendingAccommodationRows.length ? (
                          pendingAccommodationRows.map((item) => (
                            <tr
                              key={item.discrepancy.id}
                              className={item.discrepancy.id === selectedDiscrepancyId ? "rowActive" : ""}
                            >
                              <td>{item.discrepancy.discrepancyNo}</td>
                              <td>
                                {item.discrepancy.godownName} / {item.discrepancy.stackNo}
                              </td>
                              <td>{item.discrepancy.cropRegistrationCode}</td>
                              <td>{item.discrepancy.farmerName}</td>
                              <td>{formatNumber(item.discrepancy.excessQtyQtl)} QTL</td>
                              <td>{formatNumber(item.mappedQtyQtl)} QTL</td>
                              <td>{formatNumber(item.remainingQtyQtl)} QTL</td>
                              <td>{item.remainingBags}</td>
                              <td>
                                <button
                                  className="secondaryButton smallButton"
                                  onClick={() => {
                                    setSelectedDiscrepancyId(item.discrepancy.id);
                                    setDiscrepancyWorkflowMode("accommodation");
                                  }}
                                  type="button"
                                >
                                  Open Adjustment
                                </button>
                              </td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={9} className="emptyStateCell">
                              No pending accommodation rows remain.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </section>
              )}
            </div>
          )}

          {activeView === "validations" && (
            <div className="contentStack">
              <section className="panel twoColumn">
                <article className="infoCard">
                  <h3>Discrepancy Auto Validation</h3>
                  <ul>
                    <li>Rechecks discrepancies against actual receipt totals for every registration</li>
                    <li>Removes invalid discrepancy records automatically</li>
                    <li>Resolves duplicate discrepancy rows for the same receipt</li>
                  </ul>
                  <button className="primaryButton" onClick={runDiscrepancyValidation} type="button">
                    Run Discrepancy Auto Validation
                  </button>
                </article>
                <article className="infoCard">
                  <h3>Lot Auto Validation</h3>
                  <ul>
                    <li>Rechecks lots against actual receipt allocations</li>
                    <li>Updates lot quantity and status when stored value differs from allocation sum</li>
                    <li>Flags lots whose allocation total still exceeds 200 QTL</li>
                  </ul>
                  <button className="primaryButton" onClick={runLotValidation} type="button">
                    Run Lot Auto Validation
                  </button>
                </article>
                <article className="infoCard">
                  <h3>Lot Reindex</h3>
                  <ul>
                    <li>Renumbers existing lots registration-wise into gap-free sequence</li>
                    <li>Updates receipt allocation lot codes to match the new lot numbers</li>
                    <li>Keeps the same lot records and traceability ids intact</li>
                  </ul>
                  <button className="primaryButton" onClick={runLotReindex} type="button">
                    Run Lot Reindex
                  </button>
                </article>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Last Validation Result</h3>
                  <span>{validationSummary?.type ?? "No validation run yet"}</span>
                </div>
                {validationSummary ? (
                  <div className="detailGrid">
                    {validationSummary.type === "DISCREPANCY_AUTO_VALIDATION" ? (
                      <>
                        <div className="metricBox">
                          <span>Registrations Reconciled</span>
                          <strong>{validationSummary.reconciledRegistrations ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Open Discrepancies</span>
                          <strong>{validationSummary.openDiscrepancies ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Resolved Discrepancies</span>
                          <strong>{validationSummary.resolvedDiscrepancies ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Pending Excess Qty</span>
                          <strong>{formatNumber(validationSummary.totalPendingExcessQtyQtl ?? 0)} QTL</strong>
                        </div>
                      </>
                    ) : validationSummary.type === "LOT_REINDEX" ? (
                      <>
                        <div className="metricBox">
                          <span>Total Lots Checked</span>
                          <strong>{validationSummary.totalLotsChecked ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Reindexed Lots</span>
                          <strong>{validationSummary.reindexedLots ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Updated Receipt Allocations</span>
                          <strong>{validationSummary.updatedReceiptAllocations ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Affected Registrations</span>
                          <strong>{validationSummary.affectedRegistrations ?? 0}</strong>
                        </div>
                      </>
                    ) : (
                      <>
                        <div className="metricBox">
                          <span>Total Lots Checked</span>
                          <strong>{validationSummary.totalLotsChecked ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Updated Lots</span>
                          <strong>{validationSummary.updatedLots ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Orphan Lots Removed</span>
                          <strong>{validationSummary.orphanLotsRemoved ?? 0}</strong>
                        </div>
                        <div className="metricBox">
                          <span>Over-Cap Lots</span>
                          <strong>{validationSummary.overCapLots ?? 0}</strong>
                        </div>
                      </>
                    )}
                  </div>
                ) : (
                  <p>No validation has been run from this screen yet.</p>
                )}
              </section>
            </div>
          )}

          {activeView === "lots" && (
            <div className="contentStack">
              <section className="panel">
                <div className="panelHeader">
                  <div>
                    <h3>Certification Lots</h3>
                    <span>{lots.length} total lots</span>
                  </div>
                  <div className="actionButtons">
                    <button
                      className="secondaryButton"
                      type="button"
                      onClick={() => {
                        setReportType("ADJUSTED_LOT_FORMATION_REGISTER");
                        setActiveView("reports");
                      }}
                    >
                      Adjusted Lot Formation Report
                    </button>
                    <button
                      className="secondaryButton"
                      type="button"
                      onClick={() => {
                        setReportType("ADJUSTED_LOT_LEDGER_FARMER_WISE");
                        setActiveView("reports");
                      }}
                    >
                      Farmer-wise Adjusted Lot Ledger
                    </button>
                  </div>
                </div>
                <div className="tableWrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Lot Code</th>
                        <th>Reg. Code</th>
                        <th>Godown</th>
                        <th>Stack</th>
                        <th>Bags</th>
                        <th>Qty</th>
                        <th>Balance</th>
                        <th>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {lotLedgerRows.map((lot) => (
                          <tr key={lot.id}>
                            <td>{lot.lotCode}</td>
                            <td>{lot.cropRegistrationCode}</td>
                            <td>{lot.godownName}</td>
                            <td>{lot.stackNo}</td>
                            <td>{lot.bags}</td>
                            <td>{formatNumber(lot.displayQtyQtl)} QTL</td>
                            <td>{formatNumber(lot.maxAllowedQtyQtl - lot.currentQtyQtl)} QTL</td>
                            <td>
                              <span className={`status ${lot.status.toLowerCase()}`}>{lot.status}</span>
                              {openDiscrepancyKeySet.has(`${lot.cropRegistrationCode}::${lot.stackNo}`) ? (
                                <span className="status blocked">DISCREPANCY</span>
                              ) : null}
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Recent Receipts</h3>
                  <span>{receipts.length} saved</span>
                </div>
                <div className="receiptsList">
                  {receipts.map((receipt) => (
                    <article className="receiptCard" key={receipt.id}>
                      <div>
                        <strong>{receipt.receiptNo}</strong>
                        <span>{receipt.receiptDate}</span>
                      </div>
                      <div>
                        <strong>{receipt.cropRegistrationCode}</strong>
                        <span>{receipt.farmerName}</span>
                      </div>
                      <div>
                        <strong>{receipt.lines.length} line(s)</strong>
                        <span>
                          {formatNumber(
                            receipt.lines.reduce((sum, line) => sum + line.qtyQtl, 0)
                          )}{" "}
                          QTL
                        </span>
                      </div>
                    </article>
                  ))}
                </div>
              </section>
            </div>
          )}

          {activeView === "backup" && (
            <div className="contentStack">
              <section className="panel">
                <div className="panelHeader">
                  <h3>Database Backup</h3>
                  <span>Admin only</span>
                </div>
                <div className="formGrid">
                  <label className="spanTwo">
                    <span>Backup Directory</span>
                    <input
                      value={backupDirectory}
                      onChange={(event) => setBackupDirectory(event.target.value)}
                      disabled={!isAdminUser}
                    />
                  </label>
                </div>
                <div className="actionsFooter">
                  <button
                    className="primaryButton"
                    disabled={!isAdminUser}
                    onClick={runDatabaseBackup}
                    type="button"
                  >
                    Run Database Backup
                  </button>
                </div>
                {maintenanceMessage ? <p className="tableHint">{maintenanceMessage}</p> : null}
              </section>
            </div>
          )}

          {activeView === "restore" && (
            <div className="contentStack">
              <section className="panel">
                <div className="panelHeader">
                  <h3>Database Restore</h3>
                  <span>Admin only</span>
                </div>
                <div className="formGrid">
                  <label className="spanTwo">
                    <span>Restore Folder / Directory</span>
                    <input
                      value={restoreDirectory}
                      onChange={(event) => setRestoreDirectory(event.target.value)}
                      disabled={!isAdminUser}
                    />
                  </label>
                </div>
                <div className="actionsFooter">
                  <button
                    className="primaryButton"
                    disabled={!isAdminUser}
                    onClick={runDatabaseRestore}
                    type="button"
                  >
                    Run Database Restore
                  </button>
                </div>
                {maintenanceMessage ? <p className="tableHint">{maintenanceMessage}</p> : null}
              </section>
            </div>
          )}
        </section>
      </div>

      {voucherModalOpen ? (
        <div className="modalOverlay" onClick={closeVoucherModal} role="presentation">
          <div
            className="modalCard"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
          >
            <div className="panelHeader">
              <h3>{activeVoucher ? "Edit Financial Voucher" : "Generate Financial Voucher"}</h3>
              <button className="ghostButton" onClick={closeVoucherModal} type="button">
                Close
              </button>
            </div>

            <div className="selectedCard">
              <strong>{voucherRegistration?.cropRegistrationCode ?? "-"}</strong>
              <span>{voucherRegistration?.farmerName ?? "-"}</span>
              <span>
                {voucherRegistration?.village ?? "-"}, {voucherRegistration?.district ?? "-"}
              </span>
              <span>
                Net Received {formatNumber(voucherRegistration?.totalReceivedQtl ?? 0)} QTL | Discrepancy{" "}
                {formatNumber(registrationDiscrepancyMap.get(voucherRegistrationId)?.qtyQtl ?? 0)} QTL
              </span>
            </div>

            <div className="formGrid">
              <label>
                <span>Voucher Date</span>
                <input
                  type="date"
                  value={voucherDate}
                  onChange={(event) => setVoucherDate(event.target.value)}
                />
              </label>
              <label>
                <span>Certified Rate / QTL</span>
                <input
                  type="number"
                  value={certifiedRate}
                  onChange={(event) => setCertifiedRate(event.target.value)}
                />
              </label>
              <label>
                <span>Discrepancy Rate / QTL</span>
                <input
                  type="number"
                  value={discrepancyRate}
                  onChange={(event) => setDiscrepancyRate(event.target.value)}
                />
              </label>
              <label>
                <span>Seed Payment</span>
                <input
                  type="number"
                  value={voucherDeduction}
                  onChange={(event) => setVoucherDeduction(event.target.value)}
                />
              </label>
              <label className="spanTwo">
                <span>Remarks</span>
                <input
                  value={voucherRemarks}
                  onChange={(event) => setVoucherRemarks(event.target.value)}
                />
              </label>
            </div>

            <div className="actionsFooter">
              <button
                className="primaryButton"
                onClick={() => {
                  void generateVoucher();
                }}
                type="button"
              >
                {activeVoucher ? "Update Voucher" : "Generate Voucher"}
              </button>
              <button
                className="secondaryButton"
                onClick={() => voucherPreview && downloadVoucherPdf(voucherPreview.voucher)}
                type="button"
              >
                Download PDF
              </button>
            </div>

            <section className="panel">
              <div className="panelHeader">
                <h3>{voucherPreview?.voucher.voucherNo ?? "Voucher Preview"}</h3>
                <span>A5 portrait PDF only</span>
              </div>
              {voucherPreview ? (
                <div className="slipClassic voucherClassic">
                  <div className="slipClassicRule" />
                  <div className="slipClassicTitle">
                    <img className="slipClassicLogo" src={BRAND_LOGO_SRC} alt="Krishiv Seeds" />
                    <div>{COMPANY_NAME}</div>
                    <div>FARMER OVERALL INTAKE CUM PAYMENT VOUCHER</div>
                  </div>
                  <div className="slipClassicRule" />
                  <div className="slipClassicMeta">
                    <div><strong>Voucher No.</strong> : {voucherPreview.voucher.voucherNo}</div>
                    <div><strong>Voucher Date</strong> : {formatDateDisplay(voucherPreview.voucher.voucherDate)}</div>
                    <div><strong>Season</strong> : {voucherPreview.voucher.season} {voucherPreview.voucher.year}</div>
                    <div><strong>Reg. Code</strong> : {voucherPreview.voucher.cropRegistrationCode}</div>
                  </div>
                  <div className="slipClassicDetails">
                    {[
                      ["Farmer Name", voucherPreview.voucher.farmerName],
                      ["F/H Name", voucherPreview.voucher.fatherName || "-"],
                      ["Village", voucherPreview.voucher.village || "-"],
                      ["Block", voucherPreview.voucher.block || "-"],
                      ["District", voucherPreview.voucher.district || "-"],
                      ["Crop", voucherPreview.voucher.crop],
                      ["Variety", voucherPreview.voucher.variety],
                      ["Class", voucherPreview.voucher.classStage],
                      ["Exp. Yield", `${formatNumber(voucherPreview.voucher.expectedYieldQtl)} QTL`]
                    ].map(([label, value]) => (
                      <div className="slipClassicLine" key={label}>
                        <span>{label}</span>
                        <strong>{value}</strong>
                      </div>
                    ))}
                  </div>
                  <div className="slipClassicRule" />
                  <div className="slipClassicSection">Overall Intake Detail</div>
                  <div className="slipClassicRule" />
                  <div className="tableWrap">
                    <table className="slipClassicTable">
                      <thead>
                        <tr>
                          <th>Date</th>
                          <th>Rcpt No</th>
                          <th>Vehicle No</th>
                          <th>Stack</th>
                          <th>Bags</th>
                          <th>Gross</th>
                          <th>Net</th>
                        </tr>
                      </thead>
                      <tbody>
                        {voucherPreview.voucher.lines.map((line) => (
                          <tr key={line.receiptId}>
                            <td>{formatDateDisplay(line.receiptDate)}</td>
                            <td>{line.receiptNo}</td>
                            <td>{line.vehicleNo}</td>
                            <td>{line.stackNo}</td>
                            <td>{line.bags}</td>
                            <td>{formatNumber(line.grossQtyQtl)}</td>
                            <td>{formatNumber(line.netQtyQtl)}</td>
                          </tr>
                        ))}
                        <tr className="slipClassicTotalRow">
                          <td colSpan={4}>Total</td>
                          <td>{voucherPreview.voucher.totalBags}</td>
                          <td>{formatNumber(voucherPreview.voucher.totalGrossQtyQtl)}</td>
                          <td>{formatNumber(voucherPreview.voucher.totalNetQtyQtl)}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                  {voucherPreview.hasDiscrepancy ? (
                    <>
                      <div className="slipClassicRule" />
                      <div className="slipClassicSection">Discrepancy Summary</div>
                      <div className="slipClassicDetails">
                        <div className="slipClassicLine">
                          <span>Excess Qty</span>
                          <strong>{formatNumber(voucherPreview.voucher.discrepancyQtyQtl)} QTL</strong>
                        </div>
                        <div className="slipClassicLine">
                          <span>Excess Bags</span>
                          <strong>{voucherPreview.voucher.discrepancyBags}</strong>
                        </div>
                      </div>
                    </>
                  ) : null}
                  <div className="slipClassicRule" />
                  <div className="slipClassicSection">Payment Calculation</div>
                  <div className="tableWrap">
                    <table className="slipClassicTable">
                      <thead>
                        <tr>
                          <th>Type</th>
                          <th>Qty (QTL)</th>
                          <th>Rate/QTL</th>
                          <th>Amount</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td>Certified Qty</td>
                          <td>{formatNumber(voucherPreview.voucher.certifiedQtyQtl)}</td>
                          <td>{formatNumber(voucherPreview.voucher.certifiedRatePerQtl)}</td>
                          <td>{formatNumber(voucherPreview.voucher.certifiedAmount)}</td>
                        </tr>
                        {voucherPreview.hasDiscrepancy ? (
                          <tr>
                            <td>Discrepancy Qty</td>
                            <td>{formatNumber(voucherPreview.voucher.discrepancyQtyQtl)}</td>
                            <td>{formatNumber(voucherPreview.voucher.discrepancyRatePerQtl)}</td>
                            <td>{formatNumber(voucherPreview.voucher.discrepancyAmount)}</td>
                          </tr>
                        ) : null}
                        <tr className="slipClassicTotalRow">
                          <td colSpan={3}>Gross Payable Amount</td>
                          <td>{formatNumber(voucherPreview.voucher.grossPayableAmount)}</td>
                        </tr>
                        <tr>
                          <td colSpan={3}>Seed Payment</td>
                          <td>{formatNumber(voucherPreview.voucher.deductionAmount)}</td>
                        </tr>
                        <tr className="slipClassicTotalRow">
                          <td colSpan={3}>Net Payable Amount</td>
                          <td>{formatNumber(voucherPreview.voucher.netPayableAmount)}</td>
                        </tr>
                        <tr>
                          <td colSpan={3}>Rounded Off</td>
                          <td>{formatNumber(voucherPreview.voucher.roundedOffAmount ?? 0)}</td>
                        </tr>
                        <tr className="slipClassicTotalRow">
                          <td colSpan={3}>Final Payable Amount</td>
                          <td>{formatNumber(getVoucherFinalPayable(voucherPreview.voucher))}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                  <div className="slipClassicRule" />
                  <div className="slipClassicSection">Approval Status</div>
                  <div className="slipClassicSignatures">
                    <span>Prepared By        : __________________</span>
                    <span>Checked By         : __________________</span>
                    <span>Approved By        : __________________</span>
                  </div>
                  <div className="slipClassicSignatures">
                    <span>Farmer Signature   : __________________</span>
                  </div>
                  <div className="slipClassicRule" />
                  <div className="slipClassicRule" />
                </div>
              ) : (
                <p>Enter rates and generate the voucher to preview it here.</p>
              )}
            </section>
          </div>
        </div>
      ) : null}

      {paymentLedgerVoucher ? (
        <div className="modalOverlay" onClick={closePaymentLedger} role="presentation">
          <div
            className="modalCard"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
          >
            <div className="panelHeader">
              <h3>Ledger View</h3>
              <button className="ghostButton" onClick={closePaymentLedger} type="button">
                Close
              </button>
            </div>

            <section className="panel">
              <div className="panelHeader">
                <h3>Voucher Ledger Summary</h3>
                <span>{paymentLedgerVoucher.status}</span>
              </div>
              <div className="selectedCard">
                <strong>{paymentLedgerVoucher.voucherNo}</strong>
                <span>{paymentLedgerVoucher.cropRegistrationCode} | {paymentLedgerVoucher.farmerName}</span>
                <span>
                  Final Payable {formatNumber(getVoucherFinalPayable(paymentLedgerVoucher))} | Net Paid{" "}
                  {formatNumber(getVoucherTotalPaid(paymentLedgerVoucher))} | Balance{" "}
                  {formatNumber(getVoucherBalance(paymentLedgerVoucher))}
                </span>
              </div>
            </section>

            <section className="panel">
              <div className="panelHeader">
                <h3>{editingVoucherPaymentId ? "Edit Payment Entry" : "Record Payment"}</h3>
                <span>{editingVoucherPaymentId ? "Update existing RTGS / NEFT entry" : "RTGS / NEFT entry"}</span>
              </div>
              <div className="formGrid">
                <label>
                  <span>Payment Date</span>
                  <input
                    type="date"
                    value={paymentDate}
                    onChange={(event) => setPaymentDate(event.target.value)}
                  />
                </label>
                <label>
                  <span>Amount</span>
                  <input
                    type="number"
                    value={paymentAmount}
                    onChange={(event) => setPaymentAmount(event.target.value)}
                  />
                </label>
                <label>
                  <span>Mode</span>
                  <input value="RTGS/NEFT" readOnly />
                </label>
                <label>
                  <span>Transaction No.</span>
                  <input
                    value={paymentTransactionNo}
                    onChange={(event) => setPaymentTransactionNo(event.target.value)}
                  />
                </label>
                <label className="spanTwo">
                  <span>Remarks</span>
                  <input
                    value={paymentRemarks}
                    onChange={(event) => setPaymentRemarks(event.target.value)}
                  />
                </label>
                {paymentLedgerVoucher && isVoucherLockedStatus(paymentLedgerVoucher.status) ? (
                  <label className="spanTwo">
                    <span>Admin Password</span>
                    <input
                      type="password"
                      value={paymentAdminPassword}
                      onChange={(event) => setPaymentAdminPassword(event.target.value)}
                    />
                  </label>
                ) : null}
              </div>

              <div className="actionsFooter">
                <button
                  className="primaryButton"
                  disabled={isAddingVoucherPayment}
                  onClick={() => {
                    void addVoucherPayment().catch((error) => {
                      notifyUser(error instanceof Error ? error.message : "Unable to save voucher payment.");
                    });
                  }}
                  type="button"
                >
                  {isAddingVoucherPayment
                    ? "Saving..."
                    : editingVoucherPaymentId
                      ? "Update Payment"
                      : "Add Payment"}
                </button>
                {editingVoucherPaymentId ? (
                  <button
                    className="secondaryButton"
                    onClick={() => {
                      setEditingVoucherPaymentId("");
                      setPaymentDate(new Date().toISOString().slice(0, 10));
                      setPaymentAmount("");
                      setPaymentTransactionNo("");
                      setPaymentRemarks("");
                      setPaymentAdminPassword("");
                    }}
                    type="button"
                  >
                    Cancel Edit
                  </button>
                ) : null}
                <button
                  className="secondaryButton"
                  onClick={() => downloadPaymentLedgerPdf(paymentLedgerVoucher)}
                  type="button"
                >
                  Download Ledger
                </button>
              </div>
            </section>

            <section className="panel">
              <div className="panelHeader">
                <h3>Payment History</h3>
                <span>{paymentLedgerVoucher.payments?.length ?? 0} payment(s)</span>
              </div>
              <div className="tableWrap">
                <table className="registrationTable compactTable">
                  <thead>
                    <tr>
                      <th>S.No.</th>
                      <th>Date</th>
                      <th>Mode</th>
                      <th>Transaction No.</th>
                      <th>Amount</th>
                      <th>Remarks</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(paymentLedgerVoucher.payments ?? []).length > 0 ? (
                      paymentLedgerVoucher.payments.map((payment: FinancialVoucherPayment, index) => (
                        <tr key={payment.id}>
                          <td>{index + 1}</td>
                          <td>{formatDateDisplay(payment.paymentDate)}</td>
                          <td>{payment.mode || "RTGS/NEFT"}</td>
                          <td>{payment.transactionNo}</td>
                          <td>{formatNumber(Number(payment.amount ?? 0))}</td>
                          <td>{payment.remarks || "-"}</td>
                          <td>
                            <button
                              className="smallButton"
                              type="button"
                              onClick={() => startEditVoucherPayment(paymentLedgerVoucher, payment)}
                            >
                              Edit
                            </button>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={7}>No payments recorded yet.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </div>
        </div>
      ) : null}

      {organizerAssignmentRegistration ? (
        <div className="modalOverlay" onClick={closeOrganizerAssignment} role="presentation">
          <div
            className="modalCard"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
          >
            <div className="panelHeader">
              <h3>Organizer Assignment</h3>
              <button className="ghostButton" onClick={closeOrganizerAssignment} type="button">
                Close
              </button>
            </div>

            <section className="panel">
              <div className="panelHeader">
                <h3>Farmer Registration</h3>
                <span>One farmer can be linked with one organizer only</span>
              </div>
              <div className="selectedCard">
                <strong>{organizerAssignmentRegistration.cropRegistrationCode}</strong>
                <span>{organizerAssignmentRegistration.farmerName}</span>
                <span>
                  {organizerAssignmentRegistration.village || "-"},{" "}
                  {organizerAssignmentRegistration.district || "-"}
                </span>
                <span>
                  Current Organizer: {organizerAssignmentRegistration.organizerName || "Not linked"}
                </span>
                <span>
                  Intake Received: {formatNumber(Number(organizerAssignmentRegistration.totalReceivedQtl ?? 0))} QTL
                </span>
              </div>
            </section>

            <section className="panel">
              <div className="panelHeader">
                <h3>Link Organizer</h3>
                <span>Commission will apply at fixed organizer rate per QTL</span>
              </div>
              <div className="formGrid">
                <label className="spanTwo">
                  <span>Select Organizer</span>
                  <select
                    value={organizerAssignmentOrganizerId}
                    onChange={(event) => setOrganizerAssignmentOrganizerId(event.target.value)}
                  >
                    <option value="">No Organizer Linked</option>
                    {organizers
                      .filter((item) => item.isActive)
                      .slice()
                      .sort((left, right) => left.name.localeCompare(right.name, "en", { sensitivity: "base" }))
                      .map((item) => (
                        <option key={item.id} value={item.id}>
                          {item.name} | {item.village || "-"}, {item.district || "-"} |{" "}
                          {formatNumber(item.commissionRatePerQtl)}/QTL
                        </option>
                      ))}
                  </select>
                </label>
              </div>
              <div className="actionsFooter">
                <button className="primaryButton" onClick={saveOrganizerAssignment} type="button">
                  Save Mapping
                </button>
                <button className="secondaryButton" onClick={closeOrganizerAssignment} type="button">
                  Cancel
                </button>
              </div>
            </section>
          </div>
        </div>
      ) : null}

      {organizerLedgerSummary ? (
        <div className="modalOverlay" onClick={closeOrganizerLedger} role="presentation">
          <div
            className="modalCard"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
          >
            <div className="panelHeader">
              <h3>Organizer Commission Ledger</h3>
              <button className="ghostButton" onClick={closeOrganizerLedger} type="button">
                Close
              </button>
            </div>

            <section className="panel metricsPanel">
              <div className="metricBox">
                <span>Organizer</span>
                <strong>{organizerLedgerSummary.organizer.name}</strong>
              </div>
              <div className="metricBox">
                <span>Linked Farmers</span>
                <strong>{organizerLedgerSummary.farmerCount}</strong>
              </div>
              <div className="metricBox">
                <span>Total Intake</span>
                <strong>{formatNumber(organizerLedgerSummary.totalIntakeQtl)} QTL</strong>
              </div>
              <div className="metricBox">
                <span>Rate / QTL</span>
                <strong>{formatNumber(organizerLedgerSummary.ratePerQtl)}</strong>
              </div>
              <div className="metricBox">
                <span>Gross Commission</span>
                <strong>{formatNumber(organizerLedgerSummary.grossCommissionAmount)}</strong>
              </div>
              <div className="metricBox">
                <span>Deduction</span>
                <strong>{formatNumber(organizerLedgerSummary.deductionAmount)}</strong>
              </div>
              <div className="metricBox">
                <span>Net Payable</span>
                <strong>{formatNumber(organizerLedgerSummary.netPayableAmount)}</strong>
              </div>
              <div className="metricBox">
                <span>Balance</span>
                <strong>{formatNumber(organizerLedgerSummary.balanceAmount)}</strong>
              </div>
            </section>

            <section className="panel twoColumn">
              <article className="infoCard">
                <div className="panelHeader">
                  <h3>Organizer Summary</h3>
                  <span>{organizerLedgerSummary.organizer.district || "District not set"}</span>
                </div>
                <div className="selectedCard">
                  <strong>{organizerLedgerSummary.organizer.name}</strong>
                  <span>Mobile: {organizerLedgerSummary.organizer.mobile || "-"}</span>
                  <span>
                    Village: {organizerLedgerSummary.organizer.village || "-"} | District:{" "}
                    {organizerLedgerSummary.organizer.district || "-"}
                  </span>
                  <span>
                    Deduction {formatNumber(organizerLedgerSummary.deductionAmount)} | Net Payable{" "}
                    {formatNumber(organizerLedgerSummary.netPayableAmount)}
                  </span>
                  <span>
                    Paid {formatNumber(organizerLedgerSummary.paidAmount)} | Payment Entries{" "}
                    {organizerLedgerSummary.paymentCount}
                  </span>
                </div>
                <div className="tableWrap">
                  <table className="registrationTable compactTable">
                    <thead>
                      <tr>
                        <th>Reg. Code</th>
                        <th>Farmer</th>
                        <th>Village</th>
                        <th>Received</th>
                        <th>Commission</th>
                      </tr>
                    </thead>
                    <tbody>
                      {organizerLedgerSummary.linkedRegistrations.length ? (
                        organizerLedgerSummary.linkedRegistrations
                          .slice()
                          .sort((left, right) =>
                            left.cropRegistrationCode.localeCompare(right.cropRegistrationCode, "en", {
                              sensitivity: "base"
                            })
                          )
                          .map((registration) => (
                            <tr key={registration.id}>
                              <td>{registration.cropRegistrationCode}</td>
                              <td>{registration.farmerName}</td>
                              <td>{registration.village || "-"}</td>
                              <td>{formatNumber(Number(registration.totalReceivedQtl ?? 0))} QTL</td>
                              <td>
                                {formatNumber(
                                  roundQtl(
                                    Number(registration.totalReceivedQtl ?? 0) *
                                      Number(
                                        registration.organizerCommissionRatePerQtl ??
                                          organizerLedgerSummary.ratePerQtl
                                      )
                                  )
                                )}
                              </td>
                            </tr>
                          ))
                      ) : (
                        <tr>
                          <td colSpan={5} className="emptyStateCell">
                            No farmer is linked to this organizer yet.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </article>

              <article className="infoCard">
                <div className="panelHeader">
                  <h3>{editingOrganizerPaymentId ? "Edit Commission Payment" : "Record Commission Payment"}</h3>
                  <span>Fixed commission rate per QTL</span>
                </div>
                <div className="formGrid">
                  <label>
                    <span>Payment Date</span>
                    <input
                      type="date"
                      value={organizerPaymentDate}
                      onChange={(event) => setOrganizerPaymentDate(event.target.value)}
                    />
                  </label>
                  <label>
                    <span>Amount</span>
                    <input
                      type="number"
                      min="0"
                      step="0.01"
                      value={organizerPaymentAmount}
                      onChange={(event) => setOrganizerPaymentAmount(event.target.value)}
                    />
                  </label>
                  <label className="spanTwo">
                    <span>Transaction No.</span>
                    <input
                      value={organizerPaymentTransactionNo}
                      onChange={(event) => setOrganizerPaymentTransactionNo(event.target.value)}
                      placeholder="UTR / bank transaction no."
                    />
                  </label>
                  <label className="spanTwo">
                    <span>Remarks</span>
                    <input
                      value={organizerPaymentRemarks}
                      onChange={(event) => setOrganizerPaymentRemarks(event.target.value)}
                      placeholder="Optional remarks"
                    />
                  </label>
                </div>
                <div className="actionsFooter">
                  <button className="primaryButton" onClick={saveOrganizerPayment} type="button">
                    {editingOrganizerPaymentId ? "Update Payment" : "Add Payment"}
                  </button>
                  <button
                    className="secondaryButton"
                    onClick={() => downloadOrganizerCommissionVoucherPdf(organizerLedgerSummary)}
                    type="button"
                  >
                    Download Voucher
                  </button>
                  {editingOrganizerPaymentId ? (
                    <button
                      className="secondaryButton"
                      onClick={() => {
                        setEditingOrganizerPaymentId("");
                        setOrganizerPaymentDate(new Date().toISOString().slice(0, 10));
                        setOrganizerPaymentAmount("");
                        setOrganizerPaymentTransactionNo("");
                        setOrganizerPaymentRemarks("");
                      }}
                      type="button"
                    >
                      Cancel Edit
                    </button>
                  ) : null}
                </div>
              </article>
            </section>

            <section className="panel">
              <div className="panelHeader">
                <h3>Commission Payment History</h3>
                <span>{organizerLedgerPayments.length} payment(s)</span>
              </div>
              <div className="tableWrap">
                <table className="registrationTable compactTable">
                  <thead>
                    <tr>
                      <th>S.No.</th>
                      <th>Date</th>
                      <th>Transaction No.</th>
                      <th>Amount</th>
                      <th>Remarks</th>
                      <th>Entered By</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {organizerLedgerPayments.length ? (
                      organizerLedgerPayments.map((payment, index) => (
                        <tr key={payment.id}>
                          <td>{index + 1}</td>
                          <td>{formatDateDisplay(payment.paymentDate)}</td>
                          <td>{payment.transactionNo}</td>
                          <td>{formatNumber(Number(payment.amount ?? 0))}</td>
                          <td>{payment.remarks || "-"}</td>
                          <td>{payment.createdBy || "-"}</td>
                          <td>
                            <div className="inlineActionRow">
                              <button
                                className="smallButton"
                                onClick={() => beginEditOrganizerPayment(payment)}
                                type="button"
                              >
                                Edit
                              </button>
                              <button
                                className="smallButton"
                                onClick={() => deleteOrganizerPayment(payment)}
                                type="button"
                              >
                                Delete
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={7} className="emptyStateCell">
                          No organizer commission payment recorded yet.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </div>
        </div>
      ) : null}

      {slipModalOpen ? (
        <div
          className="modalOverlay"
          onClick={closeSlipModal}
          role="presentation"
        >
          <div
            className="modalCard"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
          >
            <div className="panelHeader">
              <h3>
                {slipType === "DAILY_CONSOLIDATED"
                  ? "Day Wise Consolidated Slip"
                  : slipType === "FARMER_OVERALL"
                    ? "Farmer Overall Slip"
                    : "Farmer Single Receipt Slip"}
              </h3>
              <button className="ghostButton" onClick={closeSlipModal} type="button">
                Close
              </button>
            </div>

            <div className="selectedCard">
              <strong>{slipRegistration?.cropRegistrationCode ?? "-"}</strong>
              <span>{slipRegistration?.farmerName ?? "-"}</span>
              <span>
                {slipRegistration?.village ?? "-"}, {slipRegistration?.district ?? "-"}
              </span>
              <span>
                Expected {formatNumber(slipRegistration?.expectedYieldQtl ?? 0)} QTL | Received{" "}
                {formatNumber(slipRegistration?.totalReceivedQtl ?? 0)} QTL
              </span>
            </div>

            <div className="formGrid">
              {slipType === "FARMER_SINGLE_RECEIPT" ? (
                <label>
                  <span>Receipt No.</span>
                  <select
                    value={slipReceiptNo}
                    onChange={(event) => {
                      setSlipReceiptNo(event.target.value);
                      setSlipPreview(null);
                    }}
                  >
                    <option value="">Select receipt</option>
                    {slipReceipts.map((item) => (
                      <option key={item.id} value={item.receiptNo}>
                        {item.receiptNo} - {item.receiptDate}
                      </option>
                    ))}
                  </select>
                </label>
              ) : slipType === "DAILY_CONSOLIDATED" ? (
                <label>
                  <span>Date</span>
                  <select
                    value={slipDate}
                    onChange={(event) => {
                      setSlipDate(event.target.value);
                      setSlipPreview(null);
                    }}
                  >
                    <option value="">Select date</option>
                    {Array.from(new Set(slipReceipts.map((item) => item.receiptDate))).map((item) => (
                      <option key={item} value={item}>
                        {item}
                      </option>
                    ))}
                  </select>
                </label>
              ) : (
                <div className="infoTile">
                  <span>Slip Type</span>
                  <strong>Overall farmer intake summary</strong>
                </div>
              )}
            </div>

            <div className="actionsFooter">
              <button
                className="primaryButton"
                onClick={() =>
                  previewSlip({
                    slipType,
                    slipRegistrationId,
                    slipReceiptNo,
                    slipDate
                  })
                }
                type="button"
              >
                Preview
              </button>
              <button className="secondaryButton" onClick={printSlipPdf} type="button">
                Download PDF
              </button>
            </div>

            <section className="panel">
              <div className="panelHeader">
                <h3>{slipPreview?.title ?? "Slip Preview"}</h3>
                <span>
                  {slipPreview ? `${slipPreview.pageSize} ${slipPreview.orientation}` : "Preview not loaded"}
                </span>
              </div>
              {slipPreview ? (
                slipPreview.template ? (
                  <div className="slipClassic">
                    <div className="slipClassicRule" />
                    <div className="slipClassicTitle">
                      <img className="slipClassicLogo" src={BRAND_LOGO_SRC} alt="Krishiv Seeds" />
                      <div>{COMPANY_NAME}</div>
                      <div>
                        {slipPreview.template === "DAILY_CONSOLIDATED_CLASSIC"
                          ? "DAILY CONSOLIDATED INTAKE SLIP"
                          : slipPreview.template === "FARMER_OVERALL_CLASSIC"
                            ? "FARMER OVERALL CONSOLIDATED INTAKE SLIP"
                            : "FARMER INTAKE RECEIPT"}
                      </div>
                    </div>
                    <div className="slipClassicRule" />

                    <div className="slipClassicMeta">
                      <div><strong>Slip No.</strong> : {slipPreview.slipNo}</div>
                      <div><strong>{slipPreview.summary[1]?.label ?? ""}</strong> : {slipPreview.summary[1]?.value ?? ""}</div>
                      <div><strong>{slipPreview.summary[2]?.label ?? ""}</strong> : {slipPreview.summary[2]?.value ?? ""}</div>
                      <div><strong>{slipPreview.summary[3]?.label ?? ""}</strong> : {slipPreview.summary[3]?.value ?? ""}</div>
                    </div>

                    <div className="slipClassicDetails">
                      {slipPreview.summary.slice(4).map((item) => (
                        <div className="slipClassicLine" key={item.label}>
                          <span>{item.label}</span>
                          <strong>{item.value}</strong>
                        </div>
                      ))}
                    </div>

                    <div className="slipClassicRule" />
                    <div className="slipClassicSection">Receipt Detail</div>
                    <div className="slipClassicRule" />

                    <div className="tableWrap">
                      <table className="slipClassicTable">
                        <thead>
                          <tr>
                            {slipPreview.tableColumns.map((column) => (
                              <th key={column}>{column}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {slipPreview.tableRows.map((row, index) => (
                            <tr key={`${slipPreview.slipType}-${index}`}>
                              {slipPreview.tableColumns.map((column) => (
                                <td key={`${column}-${index}`}>{String(row[column] ?? "")}</td>
                              ))}
                            </tr>
                          ))}
                          <tr className="slipClassicTotalRow">
                            <td colSpan={Math.max(slipPreview.tableColumns.length - 3, 1)}>Total</td>
                            <td>{slipPreview.totals.find((item) => item.label === "Total Bags")?.value ?? ""}</td>
                            <td>{slipPreview.totals.find((item) => item.label === "Total Gross")?.value ?? ""}</td>
                            <td>{slipPreview.totals.find((item) => item.label === "Total Net")?.value ?? ""}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    <div className="slipClassicRule" />
                    <div className="slipClassicMeta">
                      {slipPreview.totals.find((item) => item.label === "Prev. Received") ? (
                        <div><strong>Prev. Received</strong> : {slipPreview.totals.find((item) => item.label === "Prev. Received")?.value}</div>
                      ) : null}
                      {slipPreview.totals.find((item) => item.label === "Today Received") ? (
                        <div><strong>Today Received</strong> : {slipPreview.totals.find((item) => item.label === "Today Received")?.value}</div>
                      ) : null}
                      {slipPreview.totals.find((item) => item.label === "Total Net") ? (
                        <div><strong>Net Received</strong> : {slipPreview.totals.find((item) => item.label === "Total Net")?.value}</div>
                      ) : null}
                      {slipPreview.totals.find((item) => item.label === "Balance") ? (
                        <div><strong>Balance</strong> : {slipPreview.totals.find((item) => item.label === "Balance")?.value}</div>
                      ) : null}
                    </div>
                    {slipPreview.extraSections?.map((section) => (
                      <div key={section.title}>
                        <div className="slipClassicRule" />
                        <div className="slipClassicSection">{section.title}</div>
                        <div className="slipClassicDetails">
                          {section.lines.map((item) => (
                            <div className="slipClassicLine" key={`${section.title}-${item.label}`}>
                              <span>{item.label}</span>
                              <strong>{item.value}</strong>
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}
                    {slipPreview.footerNote.trim() ? (
                      <div className="slipClassicRemarks">
                        <strong>Remarks:</strong>
                        <p>{slipPreview.footerNote}</p>
                      </div>
                    ) : null}
                    <div className="slipClassicSignatures">
                      <span>Operator Sign : ______________</span>
                      <span>Farmer Sign : ______________</span>
                      <span>Godown Incharge : ______________</span>
                    </div>
                    <div className="slipClassicRule" />
                  </div>
                ) : (
                  <p>Slip preview format is not available.</p>
                )
              ) : (
                <p>Select the required receipt/date if needed, then click preview.</p>
              )}
            </section>
          </div>
        </div>
      ) : null}

      {depositViewRegistration && depositSummary ? (
        <div
          className="modalOverlay"
          onClick={() => setDepositViewRegistrationId("")}
          role="presentation"
        >
          <div
            className="modalCard"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
          >
            <div className="panelHeader">
              <h3>Registration Deposit View</h3>
              <button
                className="ghostButton"
                onClick={() => setDepositViewRegistrationId("")}
                type="button"
              >
                Close
              </button>
            </div>

            <div className="selectedCard">
              <strong>{depositViewRegistration.cropRegistrationCode}</strong>
              <span>{depositViewRegistration.farmerName}</span>
              <span>{depositViewRegistration.village}, {depositViewRegistration.district}</span>
              <span>
                Expected {formatNumber(depositViewRegistration.expectedYieldQtl)} QTL | Received{" "}
                {formatNumber(depositViewRegistration.totalReceivedQtl)} QTL
              </span>
            </div>

            <section className="panel metricsPanel">
              <div className="metricBox">
                <span>Total Deposit</span>
                <strong>{formatNumber(depositSummary.totalNetQtl)} QTL</strong>
              </div>
              <div className="metricBox">
                <span>Total Gross</span>
                <strong>{formatNumber(depositSummary.totalGrossQtl)} QTL</strong>
              </div>
              <div className="metricBox">
                <span>Total Bags</span>
                <strong>{depositSummary.totalBags}</strong>
              </div>
              <div className="metricBox">
                <span>Lots Created</span>
                <strong>{registrationLots.length}</strong>
              </div>
              <div className="metricBox">
                <span>Discrepancy Qty</span>
                <strong>{formatNumber(depositSummary.discrepancyQtyQtl)} QTL</strong>
              </div>
              <div className="metricBox">
                <span>Shifted Qty</span>
                <strong>{formatNumber(depositSummary.shiftedQtyQtl)} QTL</strong>
              </div>
              <div className="metricBox">
                <span>Shifted Bags</span>
                <strong>{depositSummary.shiftedBags}</strong>
              </div>
              <div className="metricBox">
                <span>Receipts</span>
                <strong>{registrationReceipts.length}</strong>
              </div>
            </section>

            <section className="panel">
              <div className="panelHeader">
                <h3>Stack-wise Deposit</h3>
                <span>{stackWiseDepositRows.length} stack row(s)</span>
              </div>
              <div className="tableWrap">
                <table>
                  <thead>
                    <tr>
                      <th>Godown</th>
                      <th>Stack</th>
                      <th>Bags</th>
                      <th>Gross</th>
                      <th>Net</th>
                    </tr>
                  </thead>
                  <tbody>
                    {stackWiseDepositRows.map((item) => (
                      <tr key={`${item.godownName}-${item.stackNo}`}>
                        <td>{item.godownName}</td>
                        <td>{item.stackNo}</td>
                        <td>{item.totalBags}</td>
                        <td>{formatNumber(item.totalGrossQtl)} QTL</td>
                        <td>{formatNumber(item.totalNetQtl)} QTL</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            <section className="panel">
              <div className="panelHeader">
                <h3>Lot-wise Deposit</h3>
                <span>{registrationLots.length} lot row(s)</span>
              </div>
              <div className="tableWrap">
                <table>
                  <thead>
                    <tr>
                      <th>Lot Code</th>
                      <th>Stack</th>
                      <th>Bags</th>
                      <th>Qty</th>
                      <th>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {registrationLots.map((lot) => (
                      <tr key={lot.id}>
                        <td>{lot.lotCode}</td>
                        <td>{lot.stackNo}</td>
                        <td>{lot.bags}</td>
                        <td>{formatNumber(lot.displayQtyQtl)} QTL</td>
                        <td>
                          <span className={`status ${lot.status.toLowerCase()}`}>{lot.status}</span>
                          {openDiscrepancyKeySet.has(
                            `${lot.cropRegistrationCode}::${lot.stackNo}`
                          ) ? (
                            <span className="status blocked">DISCREPANCY</span>
                          ) : null}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            <section className="panel">
              <div className="panelHeader">
                <h3>Receipt-wise Detail</h3>
                <span>{registrationReceiptLines.length} receipt line(s)</span>
              </div>
              <div className="tableWrap">
                <table>
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Receipt</th>
                      <th>Vehicle</th>
                      <th>Godown</th>
                      <th>Stack</th>
                      <th>Bags</th>
                      <th>Gross</th>
                      <th>Net</th>
                      <th>Moisture</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {registrationReceiptLines.map((item, index) => (
                      <tr key={`${item.receiptNo}-${index}`}>
                        <td>{item.receiptDate}</td>
                        <td>{item.receiptNo}</td>
                        <td>{item.vehicleNo || "-"}</td>
                        <td>{item.godownName}</td>
                        <td>{item.stackNo}</td>
                        <td>{item.bags}</td>
                        <td>{formatNumber(item.grossWeightQtl)} QTL</td>
                        <td>{formatNumber(item.netWeightQtl)} QTL</td>
                        <td>{formatNumber(item.moisturePercent)}%</td>
                        <td>
                          <button
                            className="secondaryButton smallButton"
                            disabled={!effectivePermissions?.canEdit}
                            onClick={() => startEditReceiptFromDeposit(item.receiptNo)}
                            type="button"
                          >
                            Edit Receipt
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          </div>
        </div>
      ) : null}
    </main>
  );
}
