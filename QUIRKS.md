# Excel quirks kept for parity

Places where the Excel logic looks wrong or fragile. The Python port reproduces
them so its numbers match the workbook exactly. Each one is marked `# EXCEL-QUIRK:`
in the code. Review them, then fix them deliberately; a fix will move the
numbers away from Excel.

## Program
- Common/parking area (C30) and parking area (C38) count only the ground and 2nd
  floors; 3rd/4th floor parking SF (131,832 SF here) is excluded.

## Dev Budget
- Row 18 predevelopment interest: cost is 1e-6, but `N18 = 770,170.98 / H18` puts a
  hardcoded $770,170.98 into eligible basis (Q18), making non-basis R18 negative.
- P119 adds P101 (an empty row) instead of P100 and omits P39 and P52 (all 0 today).
- I119 per-unit total omits the Legal subtotal (I52).
- F169 land value has a hardcoded +$2,180,000 typed into the formula.
- Rows 49/50 apply the eligible factor to total cost H rather than residential share K
  (no effect while the non-residential share is 0).

## LIHTC & Dev Fee
- H174 paid developer fee for 4% new construction is typed as $3,000,000; the
  "lesser of $6M or (c)" test in its label is not applied.
- G35/G36 labels say 35-50% and <=35% AMI; criteria are >30%-<51% and <31%.
- G44: when the threshold basis limit binds, allowable basis = limit - acquisition basis.
- E76/E78 add H97 (5% of acquisition basis, 9% acq/rehab section) for every deal type.
- H197 special-needs total uses H164 (regular 4% section) instead of H185 (equal today).

## Sources & Uses
- E29 equity plug does not subtract E34, E38, E41, but E43 includes them; any
  non-zero value there makes construction sources exceed uses.

## Unit Takeoff / Unit Mix Summary
- Unit Takeoff rows 97/98: AMI reconciliation sums rows 6-68 only; variants run to 94.
- Unit Takeoff P98: 50% AMI target is typed 31 instead of linking to Unit Mix Summary R16
  (P97 shows -71).
- Unit Takeoff K118 "CHECK vs K96": plan takeoff has 309 units (339 15th St) vs 1,017.
- Unit Mix Summary scenario rows (W/Z in 68-73, 79-86, 92-97, 104-109) read blank
  cells and are always 0.

## Financing & Debt Sizing
- D46/E46: "test year" income lookup lands one column past the test year, so DSCR
  sizing uses year 17 income for test year 16 (noted on the tab as legacy).
- C187: recycled-bond amount in the rate blend is a hardcode (8,836,818.61), not B187.
- Rows 57-59 labelled "months" are payment periods; Debt 3 pays annually, so its
  55 "IO months" are 55 years.

## Operating Cash Flow
- J73: year 1 pays the full partnership management fee without the cash cap later years apply.
- K73:BO73: IF has no else branch (FALSE when cash is short); years 59-60 blank.
- X100 sale: NOI = income before debt service + monitoring/trustee/service fees
  (Y51:Y53), not reserves or property tax; subtracts empty X85; subtracts the
  year-16 loan balance (D167) for a year-15 sale.
- Row 109 profit share: years J..P exclude lease-up income (row 96) and carryback
  balance; Q..X include both.
- J131 subtracts J111/2, an empty label row.
- Typed constants: 0.75 investor share of lease-up income (I96), -85,000 in L125,
  $2.0M seller carryback (F106), 0.0616 application rate (D171).

## Flow of Funds
- M18 state equity at 8609 is a typed 0; a state 8609 installment would never fund.
- Row 89 sizes its closing remainder on $F$13 instead of its own end month F89.
- L73 refers to itself as its closing deposit (resolves to 0); row 73's S-curve
  branch subtracts the deposit monthly instead of scaling (D-L). Row 73 is Straightline today.
- Taxable pre-funding draw at month F13+K14 counts the conversion month's uses twice
  (SUMIFS ">month,<=F17" plus "=F17").
- Row 191 F14 branch sums K191:$L191, including itself (only reachable if F14 is 0 or 1).
- Row 207 `L205-(R167-R170)` uses relative R refs, comparing against six months later.
- Checks J170, J189:J191 sum only through month 40 (AZ). D191 (taxable loan, feeds
  D32) counts only draws before the C of O funding month; E191 counts all, so J191
  reads "Error" in Excel too.

## Comps / Rent Builder
- Amenity vintage points compare the first four characters of the year with
  "2020"/"2010"/...; only years ending in 0 score (2024, 2009 get 0). Likely meant the decade.
- Subject amenity premiums (rows 151/153) price W/D $75 and finishes $25 vs $100/$75
  in the comp library, and omit A/C and smart home: different scales.
- Location Premiums rows 21-32 are empty slots with a 1.34 multiplier (unused).
- Rent Builder F53:F56 AMI average weights by total units incl. manager units at 0% AMI.
- Rent Builder G57 Actual AMI skips rows 21 and 30 (row 21 is a manager line today).

## OpEx Builder / OpEx & Other Income
- Comp pick "Hyder 1" (col P) is outside the T65:CL65 lookup range: #N/A if chosen without a G override.
- AH63 "Alternates" SUMIF covers rows 71-172; other comps use 73-175.
- H20 builder total excludes contingency but I20 includes it, so H20 != I20 x units.
- Comp size adjustment `1-(1017-comp units)/100*0.1` produces factors as low as 0.06.
- Federal Minimum lookup matches only Elevator rows on region (merged column B).
- Comp building ages use YEAR(TODAY()), so values shift each Jan 1 (as in Excel).
- Master Inputs E83 contingency is stored as text '5%'.

## Leasing
- 70%/80%/market Studio rows (35/43/51) count in segment totals but have no lease-up
  formulas, so those studios never lease. 70% Studio row has no AMI in column C.
- Manager pre-leasing: month -2 for 3BR/2BR, month -1 for 1BR/Studio.
- C59 manager Studio AMI links to Rent Builder D63 instead of "MGR" (no numeric effect).
- L63/M63 sum rows 8-58, omitting manager Studio row 59.
- Parking income (row 69) is driven by 70% AMI cumulative leased (N37), not total leased (N63). Likely a real bug.
- U100 COUNT range drifted to N7:U8 (counts 16 instead of 8; value 0 today).
