"use client";

import { useEffect, useState } from "react";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import type {
  AppUser,
  AppRole,
  CertificationLot,
  DiscrepancyShift,
  FinancialVoucher,
  Godown,
  IntakeDiscrepancy,
  IntakeReceipt,
  IntakeReceiptLine,
  RegistrationRecord,
  RolePermissions,
  Stack
} from "./mvp";
import {
  calculateNetWeightQtl,
  calculateWeightPerBagKg,
  parseRegistrationWorkbook,
  parseReportWorkbookMeta
} from "./mvp";

type ViewKey =
  | "dashboard"
  | "masters"
  | "import"
  | "registrations"
  | "intake"
  | "intakeEdit"
  | "reports"
  | "finance"
  | "slips"
  | "lots"
  | "discrepancies"
  | "validations"
  | "backup"
  | "restore";
const API_BASE =
  process.env.NEXT_PUBLIC_API_BASE ??
  (typeof window !== "undefined"
    ? `${window.location.protocol}//${window.location.hostname}:4000`
    : "http://127.0.0.1:4000");
const IMPORT_PAGE_SIZE = 20;
type ImportSortKey = "farmerName" | "village" | "classStage";
type RegistrationSortKey =
  | "cropRegistrationCode"
  | "farmerName"
  | "village"
  | "expectedYieldQtl"
  | "totalReceivedQtl"
  | "balanceQtl"
  | "status";

const navItems: { key: ViewKey; label: string }[] = [
  { key: "dashboard", label: "Dashboard" },
  { key: "masters", label: "Masters" },
  { key: "import", label: "Farmer Master Import" },
  { key: "registrations", label: "Registration Master" },
  { key: "intake", label: "Intake Entry" },
  { key: "intakeEdit", label: "Intake Entry Edit" },
  { key: "reports", label: "Reports" },
  { key: "finance", label: "Financial Voucher" },
  { key: "slips", label: "Slip Print Center" },
  { key: "lots", label: "Lot Ledger" },
  { key: "discrepancies", label: "Discrepancy Register" },
  { key: "validations", label: "Validation Center" },
  { key: "backup", label: "Database Backup" },
  { key: "restore", label: "Database Restore" }
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
  | "FARMER_WISE_DETAIL"
  | "SUMMARY"
  | "DAILY_INTAKE_REGISTER"
  | "REGISTRATION_PENDING_RECEIVED"
  | "LOT_WISE_STOCK_LEDGER"
  | "STACK_WISE_STOCK_POSITION"
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

const reportTypeOptions: { value: ReportType; label: string }[] = [
  { value: "GODOWN_WISE_DETAIL", label: "Godown Wise Detail" },
  { value: "FARMER_WISE_DETAIL", label: "Farmer Wise Detail" },
  { value: "SUMMARY", label: "Summary" },
  { value: "DAILY_INTAKE_REGISTER", label: "Daily Intake Register" },
  { value: "REGISTRATION_PENDING_RECEIVED", label: "Registration Pending vs Received" },
  { value: "LOT_WISE_STOCK_LEDGER", label: "Lot-wise Stock Ledger" },
  { value: "STACK_WISE_STOCK_POSITION", label: "Stack-wise Stock Position" },
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
  return receipt.lines.reduce((sum, line) => sum + Number(line.qtyQtl ?? 0), 0);
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
  const [currentUser, setCurrentUser] = useState<AppUser | null>(null);
  const [currentPermissions, setCurrentPermissions] = useState<RolePermissions | null>(null);
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [registrations, setRegistrations] = useState<RegistrationRecord[]>([]);
  const [godowns, setGodowns] = useState<Godown[]>(defaultGodowns);
  const [stacks, setStacks] = useState<Stack[]>(defaultStacks);
  const [lots, setLots] = useState<CertificationLot[]>([]);
  const [discrepancies, setDiscrepancies] = useState<IntakeDiscrepancy[]>([]);
  const [discrepancyShifts, setDiscrepancyShifts] = useState<DiscrepancyShift[]>([]);
  const [financialVouchers, setFinancialVouchers] = useState<FinancialVoucher[]>([]);
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
  const [reportMode, setReportMode] = useState<ReportMode>("ALL");
  const [reportPreview, setReportPreview] = useState<ReportPreview | null>(null);
  const [slipSearch, setSlipSearch] = useState("");
  const [slipType, setSlipType] = useState<SlipType>("FARMER_SINGLE_RECEIPT");
  const [slipRegistrationId, setSlipRegistrationId] = useState("");
  const [slipReceiptNo, setSlipReceiptNo] = useState("");
  const [slipDate, setSlipDate] = useState("");
  const [slipPreview, setSlipPreview] = useState<SlipPreview | null>(null);
  const [slipModalOpen, setSlipModalOpen] = useState(false);
  const [voucherSearch, setVoucherSearch] = useState("");
  const [voucherRegistrationId, setVoucherRegistrationId] = useState("");
  const [voucherModalOpen, setVoucherModalOpen] = useState(false);
  const [voucherDate, setVoucherDate] = useState(new Date().toISOString().slice(0, 10));
  const [certifiedRate, setCertifiedRate] = useState("");
  const [discrepancyRate, setDiscrepancyRate] = useState("");
  const [voucherDeduction, setVoucherDeduction] = useState("0");
  const [voucherRemarks, setVoucherRemarks] = useState("");
  const [voucherPreview, setVoucherPreview] = useState<VoucherPreview | null>(null);
  const [importMessage, setImportMessage] = useState("Import the farmer master Excel to begin.");
  const [importAdminPassword, setImportAdminPassword] = useState("");
  const [backupDirectory, setBackupDirectory] = useState("D:\\KRISHIV seed DATA\\mongo-backups");
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
  const [shiftQtyQtl, setShiftQtyQtl] = useState("");
  const [shiftBags, setShiftBags] = useState("");
  const [shiftDate, setShiftDate] = useState(new Date().toISOString().slice(0, 10));
  const [shiftApprovedBy, setShiftApprovedBy] = useState("");
  const [shiftRemarks, setShiftRemarks] = useState("");
  const [validationSummary, setValidationSummary] = useState<ValidationSummary>(null);

  function authHeaders(extraHeaders?: Record<string, string>) {
    return {
      ...(currentUser
        ? {
            "x-user-email": currentUser.email,
            "x-user-role": currentUser.role
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

  function notifyUser(message: string, useDialog = true) {
    setToast(message);
    if (useDialog && typeof window !== "undefined") {
      window.alert(message);
    }
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
    };
    setCurrentUser(data.user);
    setCurrentPermissions(data.permissions);
    if (typeof window !== "undefined") {
      window.localStorage.setItem(
        "krishiv-auth",
        JSON.stringify({
          user: data.user,
          permissions: data.permissions
        })
      );
    }
    setToast(`Logged in as ${data.user.role}.`);
  }

  function handleLogout() {
    setCurrentUser(null);
    setCurrentPermissions(null);
    setLoginError("");
    setVoucherModalOpen(false);
    setSlipModalOpen(false);
    if (typeof window !== "undefined") {
      window.localStorage.removeItem("krishiv-auth");
    }
  }

  async function loadBootstrap() {
    const response = await fetchWithAuth(`${API_BASE}/api/seed/bootstrap`);
    if (!response.ok) {
      throw new Error("Unable to load seed data from API.");
    }

    const data = (await response.json()) as {
      registrations: RegistrationRecord[];
      godowns: Godown[];
      stacks: Stack[];
      lots: CertificationLot[];
      discrepancies: IntakeDiscrepancy[];
      discrepancyShifts: DiscrepancyShift[];
      financialVouchers: FinancialVoucher[];
      receipts: IntakeReceipt[];
      features?: { discrepancyWorkflow?: boolean };
    };

    setRegistrations(data.registrations ?? []);
    setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
    setStacks(data.stacks?.length ? data.stacks : defaultStacks);
    setLots(data.lots ?? []);
    setDiscrepancies(data.discrepancies ?? []);
    setDiscrepancyShifts(data.discrepancyShifts ?? []);
    setFinancialVouchers(data.financialVouchers ?? []);
    setReceipts(data.receipts ?? []);
    setFeatures({
      discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
    });
    setSelectedRegistrationId((current) => current || data.registrations?.[0]?.id || "");
    setSlipRegistrationId((current) => current || data.registrations?.[0]?.id || "");
    setSlipDate((current) => current || data.receipts?.[0]?.receiptDate || "");
    setReceiptNo(nextReceiptNo(data.receipts ?? []));
  }

  useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }
    const raw = window.localStorage.getItem("krishiv-auth");
    if (!raw) {
      return;
    }
    try {
      const parsed = JSON.parse(raw) as { user: AppUser; permissions: RolePermissions };
      setCurrentUser(parsed.user);
      setCurrentPermissions(parsed.permissions);
    } catch {
      window.localStorage.removeItem("krishiv-auth");
    }
  }, []);

  useEffect(() => {
    if (!currentUser) {
      return;
    }
    void loadBootstrap().catch((error) => {
      setToast(error instanceof Error ? error.message : "Unable to connect to API.");
      setGodowns(defaultGodowns);
      setStacks(defaultStacks);
    });
  }, [currentUser]);

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
    received: registrations.reduce((sum, item) => sum + item.totalReceivedQtl, 0),
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
  const visibleNavItems = navItems.filter(
    (item) => item.key !== "discrepancies" || features.discrepancyWorkflow
  );
  const isNavItemDisabled = (key: ViewKey) =>
    ((key === "backup" || key === "restore" || key === "masters") && !isAdminUser) ||
    ((key === "import") && !effectivePermissions?.canImport);
  const openDiscrepancyKeySet = new Set(
    openDiscrepancies.map((item) => `${item.cropRegistrationCode}::${item.stackNo}`)
  );
  const selectedDiscrepancy =
    discrepancies.find((item) => item.id === selectedDiscrepancyId) ?? openDiscrepancies[0] ?? null;
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
  const activeVoucher =
    (voucherRegistrationId ? voucherByRegistrationId.get(voucherRegistrationId) : null) ??
    voucherPreview?.voucher ??
    null;
  const voucherRows = registrations
    .filter((item) => item.status !== "BLOCKED")
    .filter((item) => {
      const query = voucherSearch.trim().toLowerCase();
      if (!query) {
        return true;
      }
      return (
        item.cropRegistrationCode.toLowerCase().includes(query) ||
        item.farmerName.toLowerCase().includes(query) ||
        item.village.toLowerCase().includes(query) ||
        item.district.toLowerCase().includes(query)
      );
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
  const topDiscrepancies = openDiscrepancies
    .slice()
    .sort((left, right) => right.excessQtyQtl - left.excessQtyQtl)
    .slice(0, 5);
  const recentReceipts = receipts.slice(0, 5);
  const topGodownStock = godowns
    .map((godown) => ({
      name: godown.name,
      qtyQtl: lots
        .filter((lot) => lot.godownId === godown.id)
        .reduce((sum, lot) => sum + Number(lot.currentQtyQtl ?? 0), 0),
      lots: lots.filter((lot) => lot.godownId === godown.id).length
    }))
    .sort((left, right) => right.qtyQtl - left.qtyQtl)
    .slice(0, 5);
  const shortSnapshot = [
    `${receipts.length} receipts`,
    `${dashboardMetrics.intakeBags} bags`,
    `${dashboardMetrics.totalLots} lots`,
    `${dashboardMetrics.discrepancyCount} open discrepancies`
  ].join(" | ");

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

    const data = (await response.json()) as {
      registrations: RegistrationRecord[];
      godowns: Godown[];
      stacks: Stack[];
      lots: CertificationLot[];
      discrepancies: IntakeDiscrepancy[];
      discrepancyShifts: DiscrepancyShift[];
      financialVouchers: FinancialVoucher[];
      receipts: IntakeReceipt[];
      features?: { discrepancyWorkflow?: boolean };
    };

    setRegistrations(data.registrations ?? []);
    setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
    setStacks(data.stacks?.length ? data.stacks : defaultStacks);
    setLots(data.lots ?? []);
    setDiscrepancies(data.discrepancies ?? []);
    setDiscrepancyShifts(data.discrepancyShifts ?? []);
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

    const isEditing = Boolean(editingReceiptNo);
    const currentReceiptNo = receiptNo;

    void (async () => {
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
      })) as {
        registrations: RegistrationRecord[];
        godowns: Godown[];
        stacks: Stack[];
        lots: CertificationLot[];
        discrepancies: IntakeDiscrepancy[];
        discrepancyShifts: DiscrepancyShift[];
        receipts: IntakeReceipt[];
        features?: { discrepancyWorkflow?: boolean };
      };

      setRegistrations(data.registrations ?? []);
      const nextGodowns = data.godowns?.length ? data.godowns : defaultGodowns;
      const nextStacks = data.stacks?.length ? data.stacks : defaultStacks;
      setGodowns(nextGodowns);
      setStacks(nextStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
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

  function deleteReceipt(receiptRefToDelete: string, receiptNoToDelete: string) {
    if (!requirePermission("canDelete", "Only Admin can delete intake entries.")) {
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
      const data = (await response.json()) as {
        registrations: RegistrationRecord[];
        godowns: Godown[];
        stacks: Stack[];
        lots: CertificationLot[];
        discrepancies: IntakeDiscrepancy[];
        discrepancyShifts: DiscrepancyShift[];
        financialVouchers: FinancialVoucher[];
        receipts: IntakeReceipt[];
        features?: { discrepancyWorkflow?: boolean };
      };
      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
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
    setShiftTargetGodownId(godowns[0]?.id ?? "");
    setShiftTargetStackNo("");
    setShiftQtyQtl(String(discrepancy.excessQtyQtl));
    setShiftBags(String(discrepancy.estimatedExcessBags));
    setShiftDate(new Date().toISOString().slice(0, 10));
    setShiftApprovedBy("");
    setShiftRemarks("");
    setToast(`Shift entry opened for ${discrepancy.discrepancyNo}.`);
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

      const data = (await response.json()) as {
        registrations: RegistrationRecord[];
        godowns: Godown[];
        stacks: Stack[];
        lots: CertificationLot[];
        discrepancies: IntakeDiscrepancy[];
        discrepancyShifts: DiscrepancyShift[];
        financialVouchers: FinancialVoucher[];
        receipts: IntakeReceipt[];
        features?: { discrepancyWorkflow?: boolean };
      };

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
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

      const data = (await response.json()) as {
        registrations: RegistrationRecord[];
        godowns: Godown[];
        stacks: Stack[];
        lots: CertificationLot[];
        discrepancies: IntakeDiscrepancy[];
        discrepancyShifts: DiscrepancyShift[];
        financialVouchers: FinancialVoucher[];
        receipts: IntakeReceipt[];
        features?: { discrepancyWorkflow?: boolean };
        validationSummary?: ValidationSummary;
      };

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setValidationSummary(data.validationSummary ?? null);
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

      const data = (await response.json()) as {
        registrations: RegistrationRecord[];
        godowns: Godown[];
        stacks: Stack[];
        lots: CertificationLot[];
        discrepancies: IntakeDiscrepancy[];
        discrepancyShifts: DiscrepancyShift[];
        financialVouchers: FinancialVoucher[];
        receipts: IntakeReceipt[];
        features?: { discrepancyWorkflow?: boolean };
        validationSummary?: ValidationSummary;
      };

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setValidationSummary(data.validationSummary ?? null);
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

      const data = (await response.json()) as {
        registrations: RegistrationRecord[];
        godowns: Godown[];
        stacks: Stack[];
        lots: CertificationLot[];
        discrepancies: IntakeDiscrepancy[];
        discrepancyShifts: DiscrepancyShift[];
        financialVouchers: FinancialVoucher[];
        receipts: IntakeReceipt[];
        features?: { discrepancyWorkflow?: boolean };
        validationSummary?: ValidationSummary;
      };

      setRegistrations(data.registrations ?? []);
      setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
      setStacks(data.stacks?.length ? data.stacks : defaultStacks);
      setLots(data.lots ?? []);
      setDiscrepancies(data.discrepancies ?? []);
      setDiscrepancyShifts(data.discrepancyShifts ?? []);
      setFinancialVouchers(data.financialVouchers ?? []);
      setReceipts(data.receipts ?? []);
      setFeatures({
        discrepancyWorkflow: Boolean(data.features?.discrepancyWorkflow)
      });
      setValidationSummary(data.validationSummary ?? null);
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

      const pdf = new jsPDF({
        orientation: "landscape",
        unit: "mm",
        format: "legal"
      });

      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(14);
      pdf.text("KRISHIV AGRI GENETICS LLP", 14, 14);
      pdf.setFontSize(11);
      pdf.text(preview.title, 14, 20);

      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(8);
      pdf.text(`Season: ${reportSeasonLabel}`, 14, 26);
      pdf.text(`Generated: ${new Date(preview.generatedAt).toLocaleString("en-IN")}`, 120, 26);

      autoTable(pdf, {
        startY: 30,
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
        startY: (metricsTable.lastAutoTable?.finalY ?? 30) + 4,
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
    const nextSlipLots = lotLedgerRows.filter((item) => item.cropRegistrationId === nextRegistration.id);
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
        Gross: Number(line.grossWeightQtl ?? 0),
        Net: Number(line.netWeightQtl ?? line.qtyQtl ?? 0)
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
        Net: sumReceiptNetQty(receipt)
      }));
      const lotSummaryLines = nextSlipLots.length
        ? nextSlipLots.map((lot) => ({
            label: `L${lot.lotNo}`,
            value: `Stack ${lot.stackNo}   ${lot.bags} Bags   ${formatNumber(lot.displayQtyQtl)} QTL`
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
        Gross: Number(line.grossWeightQtl ?? 0),
        Net: Number(line.netWeightQtl ?? line.qtyQtl ?? 0)
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
    setSlipType("FARMER_SINGLE_RECEIPT");
    setSlipRegistrationId(registrations[0]?.id ?? "");
    setSlipReceiptNo("");
    setSlipDate(receipts[0]?.receiptDate ?? "");
    setSlipPreview(null);
    setSlipModalOpen(false);
  }

  function openVoucherModal(registrationId: string) {
    if (!requirePermission("canVoucher", "Your role cannot generate or edit financial vouchers.")) {
      return;
    }
    const existingVoucher = voucherByRegistrationId.get(registrationId) ?? null;
    setVoucherRegistrationId(registrationId);
    setVoucherDate(existingVoucher?.voucherDate ?? new Date().toISOString().slice(0, 10));
    setCertifiedRate(existingVoucher ? String(existingVoucher.certifiedRatePerQtl) : "");
    setDiscrepancyRate(existingVoucher ? String(existingVoucher.discrepancyRatePerQtl) : "");
    setVoucherDeduction(existingVoucher ? String(existingVoucher.deductionAmount) : "0");
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
        remarks: voucherRemarks
      })
    });

    if (!response.ok) {
      const errorBody = await response.json().catch(() => null);
      throw new Error(errorBody?.message || "Unable to generate financial voucher.");
    }

    const data = (await response.json()) as {
      registrations: RegistrationRecord[];
      godowns: Godown[];
      stacks: Stack[];
      lots: CertificationLot[];
      discrepancies: IntakeDiscrepancy[];
      discrepancyShifts: DiscrepancyShift[];
      financialVouchers: FinancialVoucher[];
      receipts: IntakeReceipt[];
      features?: { discrepancyWorkflow?: boolean };
    };

    setRegistrations(data.registrations ?? []);
    setGodowns(data.godowns?.length ? data.godowns : defaultGodowns);
    setStacks(data.stacks?.length ? data.stacks : defaultStacks);
    setLots(data.lots ?? []);
    setDiscrepancies(data.discrepancies ?? []);
    setDiscrepancyShifts(data.discrepancyShifts ?? []);
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

  function downloadVoucherPdf(voucher: FinancialVoucher) {
    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a5"
    });
    const left = 10;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const right = pageWidth - 10;
    const hasDiscrepancy = Number(voucher.discrepancyQtyQtl ?? 0) > 0;
    const voucherMetaFontSize = 8.8;
    let y = 8;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setDrawColor(91, 61, 38);
    doc.line(left, y, right, y);
    y += 4;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.text("KRISHIV AGRI GENETICS LLP", pageWidth / 2, y, { align: "center" });
    y += 5;
    doc.text("FARMER OVERALL INTAKE CUM PAYMENT VOUCHER", pageWidth / 2, y, { align: "center" });
    y += 4;
    doc.line(left, y, right, y);
    y += 5;

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
      ["Less Deductions", "", "", formatNumber(voucher.deductionAmount)],
      ["Net Payable Amount", "", "", formatNumber(voucher.netPayableAmount)]
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

    doc.save(`${voucher.voucherNo.replace(/[^A-Za-z0-9_-]+/g, "_")}.pdf`);
    notifyUser(`Voucher ${voucher.voucherNo} downloaded.`);
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
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const left = 10;
      const right = pageWidth - 10;
      let y = 10;
      const titleText =
        slipPreview.template === "DAILY_CONSOLIDATED_CLASSIC"
          ? "DAILY CONSOLIDATED INTAKE SLIP"
          : slipPreview.template === "FARMER_OVERALL_CLASSIC"
            ? "FARMER OVERALL CONSOLIDATED INTAKE SLIP"
            : "FARMER INTAKE RECEIPT";
      const totalMap = new Map(slipPreview.totals.map((item) => [item.label, item.value]));
      const metaEntries = slipPreview.summary.slice(0, 4);
      const detailEntries = slipPreview.summary.slice(4);
      const fileName = `${(slipPreview.slipNo ?? slipPreview.title).replace(/[^A-Za-z0-9_-]+/g, "_")}.pdf`;

      doc.setDrawColor(91, 61, 38);
      doc.setLineWidth(0.3);
      doc.line(left, y, right, y);
      y += 4;

      doc.setFont("helvetica", "bold");
      doc.setFontSize(slipPreview.pageSize === "A5" ? 12 : 14);
      doc.text("KRISHIV AGRI GENETICS LLP", pageWidth / 2, y, { align: "center" });
      y += 5;
      doc.text(titleText, pageWidth / 2, y, { align: "center" });
      y += 4;
      doc.line(left, y, right, y);
      y += 5;

      doc.setFontSize(slipPreview.pageSize === "A5" ? 9 : 10);
      metaEntries.forEach((item, index) => {
        const label = index === 0 ? "Slip No." : item.label;
        const value = index === 0 ? slipPreview.slipNo ?? item.value : item.value;
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

      const body = slipPreview.tableRows.map((row) =>
        slipPreview.tableColumns.map((column) => String(row[column] ?? ""))
      );
      body.push([
        "Total",
        ...slipPreview.tableColumns.slice(1, Math.max(slipPreview.tableColumns.length - 3, 1)).map(() => ""),
        totalMap.get("Total Bags") ?? "",
        totalMap.get("Total Gross") ?? "",
        totalMap.get("Total Net") ?? ""
      ]);

      autoTable(doc, {
        startY: y,
        head: [slipPreview.tableColumns],
        body,
        theme: "grid",
        styles: {
          fontSize: slipPreview.pageSize === "A5" ? 7.5 : 8,
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
      const extraSections = slipPreview.extraSections ?? [];
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

      if (slipPreview.footerNote.trim()) {
        if (y > pageHeight - 24) {
          doc.addPage();
          y = 12;
        }
        doc.setFont("helvetica", "bold");
        doc.text("Remarks:", left, y);
        y += 4;
        doc.setFont("helvetica", "normal");
        const remarksLines = doc.splitTextToSize(slipPreview.footerNote, right - left);
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

      doc.save(fileName);
      notifyUser("Slip PDF downloaded.");
    } catch (error) {
      notifyUser(error instanceof Error ? error.message : "Unable to generate slip PDF.");
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
            <h1>KRISHIV Seed Intake</h1>
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
        <div>
          <h1>KRISHIV Seed Intake</h1>
          <p>{reportSeasonLabel}</p>
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

      <div className="layout">
        <aside className="sidebar">
          <div className="sidebarBrand">
            <div className="seal">KA</div>
            <div>
              <strong>KRISHIV AGRI GENETICS LLP</strong>
              <p>Raw seed intake and certification traceability</p>
            </div>
          </div>

          <nav className="navList">
            {visibleNavItems.map((item) => (
              <button
                className={`navButton ${activeView === item.key ? "active" : ""}`}
                disabled={isNavItemDisabled(item.key)}
                key={item.key}
                onClick={() => setActiveView(item.key)}
                type="button"
              >
                {item.label}
              </button>
            ))}
          </nav>

        </aside>

        <section className="contentArea">
          <div className="contentHeader">
            <h2>{visibleNavItems.find((item) => item.key === activeView)?.label ?? "KRISHIV Seed Intake"}</h2>
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
              <section className="panel metricsPanel">
                <div className="metricBox">
                  <span>Expected Yield</span>
                  <strong>{formatNumber(dashboardMetrics.expectedYield)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Received</span>
                  <strong>{formatNumber(dashboardMetrics.received)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Pending</span>
                  <strong>{formatNumber(dashboardMetrics.pending)} QTL</strong>
                </div>
                <div className="metricBox">
                  <span>Open Lots</span>
                  <strong>{dashboardMetrics.openLots}</strong>
                </div>
                <div className="metricBox">
                  <span>Total Intake Bags</span>
                  <strong>{dashboardMetrics.intakeBags}</strong>
                </div>
                <div className="metricBox">
                  <span>Total Lots Created</span>
                  <strong>{dashboardMetrics.totalLots}</strong>
                </div>
                <div className="metricBox">
                  <span>Full Lots</span>
                  <strong>{dashboardMetrics.fullLots}</strong>
                </div>
                <div className="metricBox">
                  <span>Active Registrations</span>
                  <strong>{dashboardMetrics.activeRegistrations}</strong>
                </div>
              </section>

              <section className="panel twoColumn">
                <article className="infoCard">
                  <h3>Stock Snapshot</h3>
                  <ul>
                    <li>{receipts.length} intake receipts saved in MongoDB</li>
                    <li>{godowns.length} godowns and {stacks.length} stacks available for intake</li>
                    <li>{dashboardMetrics.shiftedBags} bags shifted through {dashboardMetrics.shiftedCases} discrepancy shift entries</li>
                    <li>{formatNumber(dashboardMetrics.shiftedQty)} QTL already moved out through discrepancy handling</li>
                  </ul>
                </article>
                <article className="infoCard">
                  <h3>Discrepancy Snapshot</h3>
                  <ul>
                    <li>{dashboardMetrics.discrepancyCount} open discrepancy cases currently under review</li>
                    <li>{dashboardMetrics.discrepancyBags} bags still marked under discrepancy</li>
                    <li>{formatNumber(dashboardMetrics.discrepancyQty)} QTL excess quantity still pending shift or resolution</li>
                    <li>{new Set(openDiscrepancies.map((item) => `${item.godownName}-${item.stackNo}`)).size} stack groups are presently affected</li>
                  </ul>
                </article>
              </section>

              <section className="panel twoColumn">
                <article className="infoCard">
                  <h3>Recent Receipts</h3>
                  <ul>
                    {recentReceipts.map((receipt) => (
                      <li key={receipt.id}>
                        {receipt.receiptNo} | {receipt.cropRegistrationCode} |{" "}
                        {formatNumber(receipt.lines.reduce((sum, line) => sum + line.qtyQtl, 0))} QTL
                      </li>
                    ))}
                  </ul>
                </article>
                <article className="infoCard">
                  <h3>Godown Snapshot</h3>
                  <ul>
                    {topGodownStock.map((item) => (
                      <li key={item.name}>
                        {item.name} | {formatNumber(item.qtyQtl)} QTL | {item.lots} lots
                      </li>
                    ))}
                  </ul>
                </article>
              </section>

              <section className="panel twoColumn">
                <article className="infoCard">
                  <h3>Top Discrepancies</h3>
                  <ul>
                    {topDiscrepancies.length > 0 ? (
                      topDiscrepancies.map((item) => (
                        <li key={item.id}>
                          {item.cropRegistrationCode} | Stack {item.stackNo} | Excess{" "}
                          {formatNumber(item.excessQtyQtl)} QTL
                        </li>
                      ))
                    ) : (
                      <li>No open discrepancy cases</li>
                    )}
                  </ul>
                </article>
                <article className="infoCard">
                  <h3>Short Snapshot</h3>
                  <ul>
                    <li>Received {formatNumber(dashboardMetrics.received)} QTL against expected {formatNumber(dashboardMetrics.expectedYield)} QTL</li>
                    <li>Pending balance is {formatNumber(dashboardMetrics.pending)} QTL</li>
                    <li>{dashboardMetrics.discrepancyBags} bags are pending discrepancy resolution</li>
                    <li>{dashboardMetrics.shiftedBags} bags have already been shifted out of marked stacks</li>
                  </ul>
                </article>
              </section>

              <section className="panel twoColumn">
                <article className="infoCard">
                  <h3>Module Coverage</h3>
                  <ul>
                    <li>Farmer master import with pagination, sorting, and Excel export</li>
                    <li>Registration master with launch-to-intake workflow</li>
                    <li>Intake entry, edit, delete, and auto lot creation</li>
                    <li>Discrepancy register with shift entry and shift history</li>
                  </ul>
                </article>
                <article className="infoCard">
                  <h3>Current Control Mode</h3>
                  <ul>
                    <li>{features.discrepancyWorkflow ? "Over-intake is saved and flagged into discrepancy workflow" : "Over-intake is hard blocked at save time"}</li>
                    <li>Lot cap remains fixed at 200 QTL</li>
                    <li>Stack-wise segregation remains active for lot continuity</li>
                    <li>Certification-facing stock should exclude unresolved discrepancy quantity</li>
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
                              </div>
                            ) : (
                              <button
                                className="secondaryButton smallButton"
                                onClick={() => setDepositViewRegistrationId(item.id)}
                                type="button"
                              >
                                Deposit View
                              </button>
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
                      editingReceiptNo ? !effectivePermissions?.canEdit : !effectivePermissions?.canEntry
                    }
                    onClick={saveReceipt}
                    type="button"
                  >
                    {editingReceiptNo ? "Update Intake Receipt" : "Save Intake Receipt"}
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
                  <label>
                    <span>Season</span>
                    <input value={reportSeasonLabel} onChange={(event) => setReportSeasonLabel(event.target.value)} />
                  </label>
                  <label>
                    <span>From Date</span>
                    <input type="date" value={reportFromDate} onChange={(event) => setReportFromDate(event.target.value)} />
                  </label>
                  <label>
                    <span>To Date</span>
                    <input type="date" value={reportToDate} onChange={(event) => setReportToDate(event.target.value)} />
                  </label>
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
                    <span>District</span>
                    <input value={reportDistrict} onChange={(event) => setReportDistrict(event.target.value)} />
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
                    <span>Farmer Name</span>
                    <input value={reportFarmerName} onChange={(event) => setReportFarmerName(event.target.value)} />
                  </label>
                  <label>
                    <span>Report Mode</span>
                    <select value={reportMode} onChange={(event) => setReportMode(event.target.value as ReportMode)}>
                      <option value="ALL">Accepted + Discrepancy</option>
                      <option value="ACCEPTED_ONLY">Accepted Only</option>
                      <option value="DISCREPANCY_ONLY">Discrepancy Only</option>
                    </select>
                  </label>
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
                    value={voucherSearch}
                    onChange={(event) => setVoucherSearch(event.target.value)}
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
                        <th>Total Net</th>
                        <th>Discrepancy</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {voucherRows.map((item, index) => {
                        const discrepancyInfo =
                          registrationDiscrepancyMap.get(item.id) ?? { qtyQtl: 0, bags: 0 };
                        const existingVoucher = voucherByRegistrationId.get(item.id);
                        return (
                          <tr key={item.id}>
                            <td>{index + 1}</td>
                            <td>{item.cropRegistrationCode}</td>
                            <td>{item.farmerName}</td>
                            <td>{item.village}</td>
                            <td>{item.district}</td>
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
                      })}
                    </tbody>
                  </table>
                </div>
              </section>

              <section className="panel">
                <div className="panelHeader">
                  <h3>Voucher Register</h3>
                  <span>{financialVouchers.length} voucher(s)</span>
                </div>
                <div className="tableWrap">
                  <table className="registrationTable compactTable">
                    <thead>
                      <tr>
                        <th>Voucher No.</th>
                        <th>Date</th>
                        <th>Reg. Code</th>
                        <th>Farmer</th>
                        <th>Net Qty</th>
                        <th>Discrepancy Qty</th>
                        <th>Net Payable</th>
                        <th>Status</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {financialVouchers.map((voucher) => (
                        <tr key={voucher.id}>
                          <td>{voucher.voucherNo}</td>
                          <td>{voucher.voucherDate}</td>
                          <td>{voucher.cropRegistrationCode}</td>
                          <td>{voucher.farmerName}</td>
                          <td>{formatNumber(voucher.totalNetQtyQtl)} QTL</td>
                          <td>{formatNumber(voucher.discrepancyQtyQtl)} QTL</td>
                          <td>{formatNumber(voucher.netPayableAmount)}</td>
                          <td><span className="status active">{voucher.status}</span></td>
                          <td>
                            <button
                              className="smallButton"
                              type="button"
                              onClick={() => downloadVoucherPdf(voucher)}
                            >
                              Download PDF
                            </button>
                          </td>
                        </tr>
                      ))}
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
                  <span>Search registration and open the required slip format</span>
                </div>
                <div className="filtersBar">
                  <input
                    placeholder="Search by reg. code, farmer, village, district, or class"
                    value={slipSearch}
                    onChange={(event) => setSlipSearch(event.target.value)}
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
                <div className="tableHint">
                  {slipRegistrationRows.length} registration row(s)
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
                          <tr key={item.id}>
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
                              {item.status !== "RESOLVED" ? (
                                <button
                                  className="secondaryButton smallButton"
                                  onClick={() => startShiftEntry(item)}
                                  type="button"
                                >
                                  Shift Excess
                                </button>
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
                  <h3>Certification Lots</h3>
                  <span>{lots.length} total lots</span>
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
                <span>Deductions</span>
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
                    <div>KRISHIV AGRI GENETICS LLP</div>
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
                          <td colSpan={3}>Less Deductions</td>
                          <td>{formatNumber(voucherPreview.voucher.deductionAmount)}</td>
                        </tr>
                        <tr className="slipClassicTotalRow">
                          <td colSpan={3}>Net Payable Amount</td>
                          <td>{formatNumber(voucherPreview.voucher.netPayableAmount)}</td>
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
                      <div>KRISHIV AGRI GENETICS LLP</div>
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
