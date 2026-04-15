export type RegistrationStatus =
  | "ACTIVE"
  | "BLOCKED"
  | "CLOSED"
  | "EXHAUSTED"
  | "REJECTED";

export type QualityStatus =
  | "ACCEPTED"
  | "HOLD"
  | "REJECTED";

export interface RegistrationSnapshot {
  id: string;
  seasonCode: string;
  cropRegistrationCode: string;
  farmerId: string;
  farmerName: string;
  crop: string;
  variety: string;
  classStage: string;
  certifiedAreaAcre: number;
  expectedYieldQtl: number;
  allowedIntakeQtl: number;
  totalReceivedQtl: number;
  status: RegistrationStatus;
}

export interface OpenLotSnapshot {
  id: string;
  lotNo: number;
  lotCode: string;
  cropRegistrationId: string;
  farmerId: string;
  crop: string;
  variety: string;
  classStage: string;
  godownId: string;
  stackId: string;
  currentQtyQtl: number;
  maxAllowedQtyQtl: number;
  qualityStatus: QualityStatus;
  status: "OPEN" | "FULL" | "CLOSED";
}

export interface IntakeLineInput {
  cropRegistrationId: string;
  farmerId: string;
  crop: string;
  variety: string;
  classStage: string;
  godownId: string;
  stackId: string;
  qtyQtl: number;
  qualityStatus: QualityStatus;
}

export interface LotAllocationPlanLine {
  lotId?: string;
  lotNo: number;
  lotCode: string;
  godownId: string;
  stackId: string;
  allocatedQtyQtl: number;
  resultingLotQtyQtl: number;
  creationReason: "CONTINUATION" | "NEW_STACK" | "QTY_SPLIT" | "NEW_FIRST_LOT";
  requiresCreate: boolean;
}
