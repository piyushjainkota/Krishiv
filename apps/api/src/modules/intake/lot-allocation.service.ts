import { planLotAllocation } from "@krishiv/domain/lot-allocation";
import type {
  IntakeLineInput,
  LotAllocationPlanLine,
  OpenLotSnapshot,
  RegistrationSnapshot
} from "@krishiv/domain/types";

export class LotAllocationService {
  plan(params: {
    registration: RegistrationSnapshot;
    intakeLine: IntakeLineInput;
    openLots: OpenLotSnapshot[];
    toleranceQtl?: number;
  }): LotAllocationPlanLine[] {
    return planLotAllocation(params);
  }
}
