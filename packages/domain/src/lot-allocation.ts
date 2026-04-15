import type {
  IntakeLineInput,
  LotAllocationPlanLine,
  OpenLotSnapshot,
  RegistrationSnapshot
} from "./types";

const LOT_CAP_QTL = 200;

function roundQtl(value: number): number {
  return Number(value.toFixed(2));
}

function assertRegistrationEligible(
  registration: RegistrationSnapshot,
  incomingQtyQtl: number,
  toleranceQtl = 0
): void {
  if (registration.status !== "ACTIVE") {
    throw new Error("Registration is not active for intake.");
  }

  if (registration.certifiedAreaAcre <= 0) {
    throw new Error("Certified area must be greater than zero.");
  }

  if (registration.expectedYieldQtl <= 0) {
    throw new Error("Expected yield must be greater than zero.");
  }

  const projectedTotal = roundQtl(registration.totalReceivedQtl + incomingQtyQtl);
  const allowed = roundQtl(registration.allowedIntakeQtl + toleranceQtl);

  if (projectedTotal > allowed) {
    throw new Error("Incoming quantity exceeds registration allowed intake.");
  }
}

function getNextLotNo(openLots: OpenLotSnapshot[]): number {
  if (openLots.length === 0) {
    return 1;
  }

  return Math.max(...openLots.map((lot) => lot.lotNo)) + 1;
}

function buildLotCode(seasonCode: string, regCode: string, lotNo: number): string {
  return `${seasonCode}/${regCode}/L${lotNo}`;
}

export function planLotAllocation(params: {
  registration: RegistrationSnapshot;
  intakeLine: IntakeLineInput;
  openLots: OpenLotSnapshot[];
  toleranceQtl?: number;
}): LotAllocationPlanLine[] {
  const { registration, intakeLine, openLots, toleranceQtl = 0 } = params;

  if (intakeLine.qtyQtl <= 0) {
    throw new Error("Intake quantity must be greater than zero.");
  }

  assertRegistrationEligible(registration, intakeLine.qtyQtl, toleranceQtl);

  const compatibleLots = openLots
    .filter((lot) =>
      lot.cropRegistrationId === intakeLine.cropRegistrationId &&
      lot.farmerId === intakeLine.farmerId &&
      lot.crop === intakeLine.crop &&
      lot.variety === intakeLine.variety &&
      lot.classStage === intakeLine.classStage &&
      lot.stackId === intakeLine.stackId &&
      lot.godownId === intakeLine.godownId &&
      lot.qualityStatus === intakeLine.qualityStatus &&
      lot.status === "OPEN"
    )
    .sort((a, b) => a.lotNo - b.lotNo);

  const plans: LotAllocationPlanLine[] = [];
  let remainingQty = roundQtl(intakeLine.qtyQtl);
  let nextLotNo = getNextLotNo(openLots);
  let usedCompatibleLot = false;

  for (const lot of compatibleLots) {
    if (remainingQty <= 0) {
      break;
    }

    const lotBalance = roundQtl(LOT_CAP_QTL - lot.currentQtyQtl);
    if (lotBalance <= 0) {
      continue;
    }

    const allocatedQtyQtl = roundQtl(Math.min(remainingQty, lotBalance));
    remainingQty = roundQtl(remainingQty - allocatedQtyQtl);
    usedCompatibleLot = true;

    plans.push({
      lotId: lot.id,
      lotNo: lot.lotNo,
      lotCode: lot.lotCode,
      godownId: lot.godownId,
      stackId: lot.stackId,
      allocatedQtyQtl,
      resultingLotQtyQtl: roundQtl(lot.currentQtyQtl + allocatedQtyQtl),
      creationReason: "CONTINUATION",
      requiresCreate: false
    });
  }

  while (remainingQty > 0) {
    const allocatedQtyQtl = roundQtl(Math.min(remainingQty, LOT_CAP_QTL));
    const lotNo = nextLotNo++;
    remainingQty = roundQtl(remainingQty - allocatedQtyQtl);

    plans.push({
      lotNo,
      lotCode: buildLotCode(registration.seasonCode, registration.cropRegistrationCode, lotNo),
      godownId: intakeLine.godownId,
      stackId: intakeLine.stackId,
      allocatedQtyQtl,
      resultingLotQtyQtl: allocatedQtyQtl,
      creationReason: usedCompatibleLot ? "QTY_SPLIT" : compatibleLots.length === 0 && openLots.length > 0 ? "NEW_STACK" : "NEW_FIRST_LOT",
      requiresCreate: true
    });
  }

  return plans;
}
