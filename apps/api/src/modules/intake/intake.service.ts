import type { CreateIntakeReceiptInput } from "./intake.schemas";
import { LotAllocationService } from "./lot-allocation.service";

export class IntakeService {
  constructor(
    private readonly lotAllocationService = new LotAllocationService()
  ) {}

  async previewAllocation(input: CreateIntakeReceiptInput) {
    return input.lines.map((line, index) => {
      const registration = {
        id: line.cropRegistrationId,
        seasonCode: "2026KH",
        cropRegistrationCode: `REG-${index + 1}`,
        farmerId: line.farmerId,
        farmerName: "Demo Farmer",
        crop: line.crop,
        variety: line.variety,
        classStage: line.classStage,
        certifiedAreaAcre: 5,
        expectedYieldQtl: 250,
        allowedIntakeQtl: 250,
        totalReceivedQtl: 0,
        status: "ACTIVE" as const
      };

      return {
        lineIndex: index,
        allocations: this.lotAllocationService.plan({
          registration,
          intakeLine: line,
          openLots: []
        })
      };
    });
  }
}
