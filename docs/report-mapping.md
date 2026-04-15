# Report Mapping

## Godown Wise Detail

### Filters

- season
- date range
- godown
- crop
- variety
- class stage

### Row Grain

One row per intake receipt allocation grouped by receipt line and lot allocation.

### Data Mapping

- crop -> `crop_registrations.crop`
- variety -> `crop_registrations.variety`
- class/stage -> `crop_registrations.class_stage`
- seed grower name -> `farmers.name`
- father/husband name -> `farmers.father_husband_name`
- village -> `farmers.village`
- expected yield -> `crop_registrations.expected_yield_qtl`
- intake date -> `intake_receipts.receipt_datetime`
- moisture % -> `intake_receipt_lines.moisture_percent`
- no. of bags -> `intake_receipt_lines.no_of_bags`
- weight per bag -> `intake_receipt_lines.weight_per_bag_kg`
- total quantity -> `lot_allocations.allocated_qty_qtl`
- farmer identification code -> `farmers.farmer_code` fallback `crop_registrations.crop_registration_code`
- stored bags -> `intake_receipt_lines.no_of_bags` unless later bag ledger exists
- stored weight per bag -> `intake_receipt_lines.weight_per_bag_kg`
- stored quantity -> `lot_allocations.allocated_qty_qtl`
- block/stack no. -> `stacks.stack_no`
- lot no. -> `certification_lots.lot_code`
- remarks -> coalesce line remarks and allocation remarks

## Farmer Wise Detail

### Filters

- season
- farmer
- village
- crop
- registration code
- date range

### Row Grain

One row per intake receipt line allocation under a farmer registration.

### Data Mapping

- crop -> `crop_registrations.crop`
- variety -> `crop_registrations.variety`
- class/stage -> `crop_registrations.class_stage`
- grower name -> `farmers.name`
- father/husband name -> `farmers.father_husband_name`
- village -> `farmers.village`
- area offered -> `crop_registrations.certified_area_acre`
- expected yield -> `crop_registrations.expected_yield_qtl`
- field P.V. bags -> optional future field survey source
- intake date -> `intake_receipts.receipt_datetime`
- intake receipt no. -> `intake_receipts.receipt_no`
- moisture % -> `intake_receipt_lines.moisture_percent`
- no. of bags -> `intake_receipt_lines.no_of_bags`
- total quantity -> `lot_allocations.allocated_qty_qtl`
- farmer id / registration code -> `crop_registrations.crop_registration_code`
- lot no. -> `certification_lots.lot_code`
- stack no. -> `stacks.stack_no`
- godown -> `godowns.name`
- remarks -> merged remarks

## Summary

### Filters

- season
- crop
- variety
- class stage
- godown optional

### Row Grain

One row per crop + variety + class stage combination for the selected filter set.

### Aggregations

- expected yield -> sum of `crop_registrations.expected_yield_qtl`
- actual intake bags -> sum of `intake_receipt_lines.no_of_bags`
- actual intake weight per bag -> weighted average of bag weights if needed for display
- actual intake total quantity -> sum of `lot_allocations.allocated_qty_qtl`
- raw seed handled bags -> current raw stock bags derived from stock ledger
- raw seed handled quantity -> current raw stock quantity from open lot balances
- balance bags available -> future bag ledger derived field
- quantity available -> current available raw stock
- graded seed fields -> future grading module outputs
- stored in godown -> grouped stack listing from current lot positions
- remarks -> summary remarks or exceptions

## Reconciliation Rules

1. farmer-wise total quantity must equal sum of linked lot allocations for same filters
2. godown-wise total quantity must equal sum of same source allocations for same filters
3. summary total quantity must match farmer-wise total and godown-wise total
4. any report export row must include hidden source ids for traceability where format allows
