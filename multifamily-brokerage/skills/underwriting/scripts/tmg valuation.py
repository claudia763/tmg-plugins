#!/usr/bin/env python3
"""
TMG Multifamily Valuation Model — Python port
=============================================

Reverse-engineered from the Excel underwriting workbook
("DG  Westlake Apartments  862026.xlsx"), specifically the three broker-facing
tabs and the hidden engine they drive:

    Assumptions  ->  operating / growth / acquisition / loan / syndication inputs
    Value-Add    ->  renovation, amenity, and operations programs (top-3 per
                     category flow into income, rent growth, expenses, budget)
    Factors      ->  exit-cap build-up (sale-comp cap rate + risk adjustments)
    UW - F&C     ->  (hidden) Year-1 underwriting, 6-year proforma, debt,
                     reversion, project & LP cash flows  — replicated here

Pure standard library (no Excel, no numpy) so it runs on any Linux server
with Python 3.8+.

PER-DEAL USE (skill Phase A): copy this file into the job folder and replace
the PROPERTY / T12_MONTHLY / ASSUMPTIONS / VALUE_ADD_ITEMS / FACTORS config
blocks with the new deal's data; the shipped values are the worked example
(Westlake Apartments, Lubbock TX, validated against the Excel model to
~0.000% on every metric). EXCEL_TARGETS is optional — leave it from the
prior deal and ignore the validation table, or update it if you have a
filled-out workbook to back into.

Target metrics reproduced within +/- 2% of the Excel model:
    Project IRR, Average Cash-on-Cash, DSCR (T-3 and Pro Forma Year 1).

Run:  python3 tmg_valuation.py          # prints proforma + metrics + validation
"""

from __future__ import annotations
import math

# ============================================================================
# 1. PROPERTY / T-12 REFERENCE DATA  (Master tab + monthly actuals in UW - F&C)
# ----------------------------------------------------------------------------
# Monthly series run oldest -> newest (Aug-2025 .. Jul-2026 in the workbook).
# T-12 = sum of 12 months; T-3 annualized = sum of last 3 months * 4.
# ============================================================================

PROPERTY = {
    "name": "Westlake Apartments",
    "units": 174,
    "rentable_sf": 170_451,
    "avg_unit_rent_market": 950.5229885057471,   # Master!J34 (mkt rent / unit / mo)
    "total_tax_rate": 0.01769366,                # Master!B24 (all taxing entities)
}

T12_MONTHLY = {
    # -------- income (negative = deduction from Gross Potential Rent) --------
    "market_rent":   [165391.0] * 12,
    "loss_to_lease": [-5393.916547] * 12,
    "vacancy":       [-5775.450119, -5775.450119, -11050.223453, -10170.203453,
                      3894.076547, -10456.603453, -12089.753453, -8396.453453,
                      -6557.263453, -7920.973453, -6561.193453, -9453.493453],
    "concessions":   [0.0] * 12,
    "bad_debt":      [0.0] * 12,
    "rubs_electric": [0.0] * 12,
    "rubs_water":    [0.0] * 12,
    "rubs_trash":    [0.0] * 12,
    "other_income":  [81.666667, 81.666667, 60.0, 5.0, 180.0, 140.0,
                      0.0, 257.46, 146.2, 312.58, 100.0, 0.0],
    # -------- expenses -------------------------------------------------------
    "contract_services":   [1960.123333, 1960.123333, 869.8, 446.49, 4564.08,
                            965.11, 1682.16, 965.11, 965.11, 2076.35, 369.73, 195.94],
    "repairs_maintenance": [3330.056667, 3330.056667, 879.51, 5497.51, 3613.15,
                            2265.08, 2827.98, 4363.07, 2477.33, 1929.27, 4143.14, 5468.96],
    "administration":      [588.746667, 588.746667, 156.33, 1056.83, 553.08,
                            1093.19, 233.28, 1233.28, 851.15, 592.0, 233.28, 236.28],
    "marketing":           [1724.766667, 1724.766667, 1565.13, 1422.1, 2187.07,
                            2154.14, 1220.48, 1952.21, 1830.06, 2500.98, 2190.41, 2266.92],
    "payroll":             [20566.136667, 20566.136667, 18586.33, 19355.22, 23756.86,
                            24992.07, 20406.76, 24038.4, 27452.32, 25874.42, 19298.37, 19922.86],
    "util_electric":       [3202.163333, 3202.163333, 1842.27, 2637.48, 5126.74,
                            4813.05, 6308.59, 2898.86, 4487.42, 5231.28, 5636.95, 5829.88],
    "util_water_sewer":    [5373.65, 5373.65, 4505.63, 6562.27, 5053.05, 5437.91,
                            6261.8, 6759.91, 4992.86, 5098.4, 5208.01, 5266.17],
    "util_trash":          [4004.226667, 4004.226667, 5324.13, 3231.45, 3457.1,
                            3969.09, 3085.5, 3906.58, 3741.54, 4841.07, 4704.52, 3181.14],
    "util_gas":            [0.0] * 12,
    "management_fee":      [4629.1, 4629.1, 4470.21, 4494.96, 4922.13, 4310.41,
                            4437.22, 4555.74, 4607.58, 4571.66, 4606.08, 4516.31],
    "insurance":           [11907.0, 11907.0, 11907.0, 11907.0, 11907.0, 11599.0,
                            11907.0, 11599.0, 11907.0, 12907.0, 11599.0, 11599.0],
    "taxes":               [12754.62]*7 + [12754.92]*5,
}

# Agency (Fannie) expense benchmarks, $/unit/yr — 'UW - F&C' columns CQ:CS.
# Used for any expense line flagged "x" on the Assumptions tab.
AGENCY_BENCHMARKS = {          # "Fannie" comp set (Assumptions!G27)
    "contract_services": 250,
    "repairs_maintenance": 650,
    "administration": 250,
    "marketing": 150,
    "payroll": 1400,           # (not flagged here — overridden at $200k)
    "insurance": 800,
}

# ============================================================================
# 2. ASSUMPTIONS TAB  (blue broker-input cells)
# ============================================================================

ASSUMPTIONS = {
    # ---- Operating, pro forma Year 1 ----
    "year1_rent_growth": 0.03,          # G17
    "loss_to_lease_pct": 0.00,          # G18  (% of market rent)
    "vacancy_pct": 0.06,                # G19  (% of GPR)
    "concessions_pct": 0.00,            # G20
    "bad_debt_pct": 0.00,               # G21
    "other_income_override": None,      # G25  (None = T-12 + Value-Add items)

    # ---- Expenses, Year 1.  "agency" -> benchmark $/unit; number -> $ override;
    #      None -> T-12 actual ----
    "expense_comp_set": "Fannie",       # G27
    "contract_services": "agency",      # F28 = x
    "repairs_maintenance": "agency",    # F29 = x
    "administration": "agency",         # F30 = x
    "marketing": "agency",              # F31 = x
    "payroll": 200_000,                 # G32 ($ override)
    "util_electric": None,              # T-12
    "util_water_sewer": None,           # T-12 net of Value-Add water savings
    "util_trash": None,                 # T-12
    "util_gas": None,                   # T-12
    "management_fee_pct": 0.03,         # G37  (% of EGI)
    "insurance": "agency",              # F38 = x  ($800/unit)
    "tax_assessment_factor": 0.75,      # G39
    "capex_reserve_per_unit": 350,      # G41

    # ---- Growth rates, years 2+ ----
    "gpr_growth": 0.02,                 # G43
    "other_income_growth": 0.02,        # G44
    "expense_growth": 0.02,             # G45
    "tax_growth": 0.02,                 # G46

    # Stabilized (reversion-year) economic loss — Assumptions!G59, inserted
    # 8/2026 directly below Terminal Cap Rate, default 8.0%. Flows to
    # 'UW - F&C' AU50; the sheet then walks it back toward acquisition:
    #   AT50 (yr5) = AU50 + 0.005      AS50 (yr4) = AT50 + 0.005
    #   AR50 (yr3) = AS50 + 0.005
    #   AQ50 (yr2) = (AR50 + SUM(AM8:AM10)) / 2 + 0.005
    # (AM8:AM10 = Year-1 vacancy + concessions + bad debt).
    # See econ_loss_schedule(); replaces the old hardcoded AQ50:AU50 constants.
    "reversion_econ_loss": 0.08,        # G59

    # ---- Acquisition / disposition ----
    "purchase_price": 11_500_000,       # G50 (strike)
    "capex_financed_in_loan": 0,        # G53
    "value_add_budget_override": None,  # G54 (None = Value-Add tab H74)
    "additional_lender_reserves": 0,    # G55 (returned at end of year 1)
    "loan_origination_pct": 0.03,       # G56 (% of loan, paid from equity)
    "hold_period_years": 5,             # G57
    "terminal_cap_override": None,      # G58 (None = Factors build-up)
    "sales_expense_pct": 0.03,          # G60

    # ---- New debt (Free & Clear) — Freddie Mac Conventional terms ----
    "ltv": 0.75,                        # G64
    "interest_rate": 0.0613,            # G65
    "io_years": 3,                      # G67
    "loan_term_months": 120,            # G68
    "amortization_months": 360,         # G69
    "supplemental_loan": 0,             # G70
    "supplemental_rate": 0.0,           # G71

    # ---- Syndication ----
    "preferred_return": 0.0,            # G88
    "lp_share": 0.80,                   # G89
    "gp_coinvest": 0,                   # G91
    "asset_mgmt_fee_pct": 0.015,        # G92 (% of gross revenue)
}

# ============================================================================
# 3. VALUE-ADD TAB  (matches the expanded Value-Add tab, Aug-2026 revision)
# ----------------------------------------------------------------------------
# Each item: category 'renovation' | 'amenity' | 'operations'.
# Only the FIRST THREE included items per category count (the sheet matches
# "r"/"rr"/"rrr" flag strings in row order) — include AT MOST three per
# category; validate_value_add() errors on more rather than silently dropping.
#
# Effects and where they land in the engine:
#   rent            renovation NOI -> extra rent growth spread over years 1-3
#   other_income    annual NOI added to Other Income (Year 1 base)
#   water_savings   subtracted from the water/sewer expense line
#   electric_savings / trash_savings / payroll_savings
#                   subtracted from the matching expense line (only when that
#                   assumption is None, i.e. T-12-driven — mirrors water)
#   bad_debt_savings  added to Year-1 net rental income (economic-loss offset)
#   tax_savings     subtracted from the Year-1 RE tax line
#   other_opex_savings  subtracted from total Year-1 opex (unclassified)
#   rubs_water / rubs_electric / rubs_trash   added to the RUBS income lines
#   prereq          cost-only row (no NOI wiring)
#   assumption      NARRATIVE row — no NOI wiring here; achieve the effect by
#                   setting the matching ASSUMPTIONS input instead (expense
#                   override / benchmark, vacancy_pct, tax_assessment_factor).
#
# Manual-count rows (Down-Unit, Non-Revenue, EV, STR, Corporate) ship with
# n=0 / noi=0 — set the real count AND $NOI before including, or the
# validator errors. Costs (numeric only) sum into the equity budget (H74).
# ============================================================================

VALUE_ADD_ITEMS = [
    # (name, category, include?, units, cost/unit, $NOI/unit/mo, effect)

    # ---- Renovations -------------------------------------------------------
    ("Light Interior Renovations",    "renovation", False, 87,  1500, 50, "rent"),
    ("Moderate Interior Renovations", "renovation", False, 174, 5000, 100, "rent"),
    ("Premium Interior Renovations",  "renovation", False, 87,  5000, 150, "rent"),
    ("Exterior Renovation & Def. Maint.", "renovation", False, 174, 3000, 0, "prereq"),
    ("Washer/Dryer Hookups",          "renovation", False, 87,  5000, 50, "rent"),
    ("Leasing Office & Gym Reno",     "renovation", False, 2, 100000, 0, "prereq"),
    ("Down-Unit Restoration",         "renovation", False, 0, 25000, 0, "rent"),      # n & noi manual
    ("Non-Revenue Unit Conversion",   "renovation", False, 0,  5000, 0, "rent"),      # n & noi manual
    ("Smart Home Technology Package", "renovation", False, 174,  750, 25, "rent"),
    ("Individual HVAC Replacement",   "renovation", False, 174, 6000, 75, "rent"),
    ("Garage / Storage Additions",    "renovation", False, 34.8, 2500, 100, "rent"),

    # ---- Amenity / Other Income -------------------------------------------
    ("Exterior Amenities",            "amenity", False, 174, 2000, 0, "prereq"),
    ("Pet Fees & Rent",               "amenity", True,  34.8, 0, 30, "other_income"),
    ("Gated Parking",                 "amenity", False, 174, 500, 10, "other_income"),
    ("Reserved Parking",              "amenity", True,  34.8, 0, 25, "other_income"),
    ("Reserved & Covered Parking",    "amenity", False, 34.8, 1500, 50, "other_income"),
    ("Property WIFI",                 "amenity", False, 174, 0, 10, "other_income"),
    ("Cable & Internet Package",      "amenity", True,  174, 0, 20, "other_income"),
    ("Valet Trash",                   "amenity", False, 174, 0, 10, "other_income"),
    ("Package Delivery Lockers",      "amenity", False, 87, 500, 10, "other_income"),
    ("Self Storage Units",            "amenity", False, 20, 1000, 50, "other_income"),
    ("Short Term Rentals / AirBnB",   "amenity", False, 0, 0, 0, "other_income"),     # n & noi manual
    ("Laundry & Vending Income",      "amenity", False, 34.8, 2000, 40, "other_income"),
    ("In-Unit Washer & Dryers",       "amenity", False, 87, 1000, 50, "other_income"),
    ("Solar Panels",                  "amenity", False, 17.4, 10000, 50, "other_income"),
    ("Fee Income Optimization",       "amenity", False, 174, 0, 15, "other_income"),
    ("Renters Insurance Program",     "amenity", False, 174, 0, 12, "other_income"),
    ("Telecom Door-Fee Agreements",   "amenity", False, 174, 0,  8, "other_income"),
    ("EV Charging Stations",          "amenity", False, 0, 5000, 30, "other_income"), # n = station count
    ("Cell Tower / Billboard Lease",  "amenity", False, 1, 0, 10, "other_income"),
    ("Corporate / Furnished Housing", "amenity", False, 0, 8000, 200, "other_income"),# n manual; 200 = NOI net of opex
    ("Amenity Fee",                   "amenity", False, 174, 0, 15, "other_income"),

    # ---- Opex Improvement & RUBS ------------------------------------------
    ("Water Conservation: Fixtures",  "operations", False, 174, 250, 20, "water_savings"),
    ("Lease-up / Occupancy Improvements", "operations", False, 174, 0, 0, "assumption"),
    ("Water Conservation: Plumbing",  "operations", False, 174, 1500, 50, "water_savings"),
    ("Implement/Increase Water RUBS", "operations", False, 174, 0, 35, "rubs_water"),
    ("Implement/Increase Electric RUBS", "operations", False, 174, 0, 22, "rubs_electric"),
    ("Misc. Billback (Pest/Trash/Gas)", "operations", False, 174, 0, 23, "rubs_trash"),
    ("Reduce Opex to Comp Averages",  "operations", False, 174, 0, 0, "assumption"),
    ("Reduce Insurance Expense",      "operations", False, 174, 0, 0, "assumption"),
    ("Reduce Tax Expense / Protest",  "operations", False, 174, 0, 0, "assumption"),
    ("LED / Common-Area Electric Retrofit", "operations", False, 174, 150, 12, "electric_savings"),
    ("Water Submetering",             "operations", False, 174, 450, 45, "rubs_water"),
    ("Trash Compactor Conversion",    "operations", False, 174, 300, 10, "trash_savings"),
    ("Revenue Mgmt / Screening",      "operations", False, 174, 0, 25, "bad_debt_savings"),
    ("Payroll Optimization",          "operations", False, 174, 0, 20, "payroll_savings"),
    ("Utility Bill Audit",            "operations", False, 174, 0,  8, "other_opex_savings"),
    ("Tax Exemption (HFC/PFC)",       "operations", False, 174, 0, 150, "tax_savings"),
]

# ----------------------------------------------------------------------------
# Trigger metadata — the Value-Add tab's "AGENT INSTRUCTIONS" column encoded
# as data. Conflicts are BY NAME (row numbers on the tab shift when rows are
# inserted; names do not — this also fixes the tab's stale "row 41" backlink
# on the Corporate/Furnished row, which should point at the STR row).
# source: RR=rent roll, T12=trailing financials, COMPS=comp tabs, WEB=web
# research, SITE=physical verification required, MODEL=model input, LEGAL=
# ordinance/statute check.  verify: 'docs' = safe on document review alone;
# 'site' = DO NOT model income until physically verified; 'flag' = narrative
# flag-for-buyer only, never modeled without executed evidence.
# ----------------------------------------------------------------------------

VALUE_ADD_RULES = {
    "Light Interior Renovations": dict(
        trigger="Rent Comparison shows renovated comps at a premium of <= $50/unit/mo over subject.",
        source=["COMPS"], verify="docs",
        conflicts=["Moderate Interior Renovations", "Premium Interior Renovations"]),
    "Moderate Interior Renovations": dict(
        trigger="Rent Comparison premium $75-125/unit/mo AND Final_RR renovation status shows a partial reno set.",
        source=["COMPS", "RR"], verify="docs",
        conflicts=["Light Interior Renovations", "Premium Interior Renovations"]),
    "Premium Interior Renovations": dict(
        trigger="Rent Comparison premium > $125/unit/mo AND comp finish level is stainless/granite.",
        source=["COMPS"], verify="docs",
        conflicts=["Light Interior Renovations", "Moderate Interior Renovations"]),
    "Exterior Renovation & Def. Maint.": dict(
        trigger="An interior tier is selected AND subject exterior trails the comp set.",
        source=["COMPS", "SITE"], verify="site", conflicts=[],
        notes="Prerequisite row; cost only. Scope needs site inspection or PCA."),
    "Washer/Dryer Hookups": dict(
        trigger="Final_RR shows no in-unit W/D and comps offer connections.",
        source=["RR", "COMPS", "SITE"], verify="site",
        conflicts=["In-Unit Washer & Dryers"],
        notes="Plumbing/electric capacity and closet space must be confirmed on site."),
    "Leasing Office & Gym Reno": dict(
        trigger="Comp set offers a renovated office or gym the subject lacks.",
        source=["COMPS", "SITE"], verify="site", conflicts=[]),
    "Down-Unit Restoration": dict(
        trigger="Final_RR occupancy status flags down/offline/non-rentable units. Set n to the exact count.",
        source=["RR", "SITE"], verify="site", conflicts=[],
        notes="Cost/unit varies widely by cause (fire vs flood vs deferred maint) — scope on site."),
    "Non-Revenue Unit Conversion": dict(
        trigger="Final_RR lease type / occupancy status shows model, office, employee, or storage units. Set n.",
        source=["RR"], verify="docs", conflicts=[],
        notes="Verify with seller the unit is genuinely convertible."),
    "Smart Home Technology Package": dict(
        trigger="Comps offer smart locks/thermostats and subject does not, OR premium interiors selected.",
        source=["COMPS", "WEB"], verify="docs", conflicts=[]),
    "Individual HVAC Replacement": dict(
        trigger="Year built pre-1980 AND T-12 shows large owner-paid electric consistent with central plant/window units.",
        source=["T12", "WEB", "SITE"], verify="site",
        conflicts=["Implement/Increase Electric RUBS"],
        notes="Conversion shifts electric to tenants — do not also model electric RUBS on the same load."),
    "Garage / Storage Additions": dict(
        trigger="Comps charge for garage or storage rental.",
        source=["COMPS", "SITE"], verify="site", conflicts=[]),

    "Exterior Amenities": dict(
        trigger="Comps have dog park/playground/pool the subject lacks. Requires confirmed excess land.",
        source=["COMPS", "SITE"], verify="site", conflicts=[],
        notes="Prerequisite row; cost only."),
    "Pet Fees & Rent": dict(
        trigger="ALWAYS include a small amount ($10-15/unit-equivalent NOI). Increase to the full $30 only if "
                "pet-friendly features exist at the subject or Exterior Amenities is selected.",
        source=["COMPS", "T12"], verify="docs", conflicts=[],
        notes="Check T-12 other income for existing pet income first — do not double count."),
    "Gated Parking": dict(
        trigger="Comps have controlled-access gating the subject lacks.",
        source=["COMPS", "SITE"], verify="site", conflicts=[]),
    "Reserved Parking": dict(
        trigger="ALWAYS include unless the subject already charges for reserved parking (check T-12 detail).",
        source=["T12"], verify="docs", conflicts=[]),
    "Reserved & Covered Parking": dict(
        trigger="Include unless subject already has carports; stacks on Reserved Parking.",
        source=["SITE"], verify="site", conflicts=[],
        notes="Verification wins over the always-include default: flag but model $0 until the "
              "surface-parking layout is confirmed."),
    "Property WIFI": dict(
        trigger="Premium interiors selected OR comps offer property WiFi.",
        source=["COMPS"], verify="docs", conflicts=[]),
    "Cable & Internet Package": dict(
        trigger="Premium interiors selected OR comps offer bulk internet. Check T-12 for an existing bulk contract.",
        source=["COMPS", "T12"], verify="docs",
        conflicts=["Telecom Door-Fee Agreements"],
        notes="Bulk package and exclusive-marketing door fees are usually mutually exclusive contracts."),
    "Valet Trash": dict(
        trigger="Units >= 200 AND (premium interiors selected OR comps offer valet). Check T-12 for existing billing.",
        source=["T12", "COMPS"], verify="docs", conflicts=[]),
    "Package Delivery Lockers": dict(
        trigger="Units >= 200. Requires common-area square footage.",
        source=["SITE"], verify="site", conflicts=[]),
    "Self Storage Units": dict(
        trigger="Demonstrated storage demand AND confirmed excess land, ideally with major-road frontage.",
        source=["WEB", "SITE"], verify="site", conflicts=[]),
    "Short Term Rentals / AirBnB": dict(
        trigger="Near airport / major business district / tourist driver AND municipality does not PROHIBIT STRs "
                "(license-required jurisdictions like OKC qualify as allowed). Set n and noi manually.",
        source=["WEB", "LEGAL"], verify="docs",
        conflicts=["Corporate / Furnished Housing"],
        notes="If prohibited, route to Corporate / Furnished Housing instead."),
    "Laundry & Vending Income": dict(
        trigger="Final_RR shows units without W/D connections. Confirm room condition and any route contract on site.",
        source=["RR", "SITE"], verify="site", conflicts=[]),
    "In-Unit Washer & Dryers": dict(
        trigger="Final_RR or amenities confirm EXISTING connections.",
        source=["RR"], verify="docs",
        conflicts=["Washer/Dryer Hookups"],
        notes="Same-unit-set conflict: hookups must already exist for this row."),
    "Solar Panels": dict(
        trigger="Modeled hold >= 7 years only. Roof age/load/orientation need inspection.",
        source=["MODEL", "SITE"], verify="site", conflicts=[]),
    "Fee Income Optimization": dict(
        trigger="T-12 fee income/unit below the comp fee schedules (scrape comp property websites for app/admin/pet fees).",
        source=["T12", "WEB"], verify="docs",
        conflicts=["Reduce Opex to Comp Averages"],
        notes="Conflict only if comp expense benchmarks are net of fee income — check before stacking."),
    "Renters Insurance Program": dict(
        trigger="T-12 shows no renters-insurance / master-liability income line.",
        source=["T12"], verify="docs", conflicts=[]),
    "Telecom Door-Fee Agreements": dict(
        trigger="T-12 shows no door-fee / revenue-share income; confirm no exclusive agreement with seller.",
        source=["T12"], verify="docs",
        conflicts=["Cable & Internet Package"]),
    "EV Charging Stations": dict(
        trigger="Built 2015+ or comp set demonstrates EV demand (PlugShare/aerials). n = station count. "
                "Panel capacity needs site check.",
        source=["WEB", "SITE"], verify="site", conflicts=[]),
    "Cell Tower / Billboard Lease": dict(
        trigger="SITUATIONAL. Highway frontage / rooftop access / carrier interest. Never model speculatively — "
                "omit unless a lease or LOI is evidenced.",
        source=["WEB", "SITE"], verify="flag", conflicts=[]),
    "Corporate / Furnished Housing": dict(
        trigger="Near hospitals / universities / corporate campuses, OR STRs are prohibited (use in place of the "
                "STR row). Set n and confirm the NOI figure is net of furnishing/turn costs.",
        source=["WEB"], verify="docs",
        conflicts=["Short Term Rentals / AirBnB"]),
    "Amenity Fee": dict(
        trigger="Exterior Amenities selected or an amenity package already exists. Check T-12 for an existing fee.",
        source=["T12"], verify="docs", conflicts=[]),

    "Water Conservation: Fixtures": dict(
        trigger="T-12 water > $65/unit/mo AND property pays water.",
        source=["T12"], verify="docs",
        conflicts=["Reduce Opex to Comp Averages"]),
    "Lease-up / Occupancy Improvements": dict(
        trigger="Final_RR physical occupancy trails comp average. Do NOT model above the comp ceiling.",
        source=["RR", "COMPS"], verify="docs", conflicts=[],
        notes="Narrative row: achieve via ASSUMPTIONS vacancy_pct, not a wired dollar line."),
    "Water Conservation: Plumbing": dict(
        trigger="T-12 water > $100/unit/mo (supply-line loss, not fixture waste). Repair scope needs inspection.",
        source=["T12", "SITE"], verify="site",
        conflicts=["Reduce Opex to Comp Averages"]),
    "Implement/Increase Water RUBS": dict(
        trigger="Comps bill water back at a higher recovery ratio than subject T-12.",
        source=["T12", "COMPS"], verify="docs",
        conflicts=["Water Submetering"]),
    "Implement/Increase Electric RUBS": dict(
        trigger="Comps bill electric back / tenant-metered AND subject T-12 shows owner-paid electric. "
                "Confirm meter configuration.",
        source=["T12", "COMPS", "SITE"], verify="docs",
        conflicts=["Individual HVAC Replacement"]),
    "Misc. Billback (Pest/Trash/Gas)": dict(
        trigger="Comps bill these back AND the lines are owner-paid in subject T-12. Net out anything already recovered.",
        source=["T12", "COMPS"], verify="docs", conflicts=[]),
    "Reduce Opex to Comp Averages": dict(
        trigger="T-12 total opex/unit exceeds comp average.",
        source=["T12", "COMPS"], verify="docs",
        conflicts=["Reduce Insurance Expense", "Reduce Tax Expense / Protest",
                   "Trash Compactor Conversion", "Payroll Optimization", "Utility Bill Audit",
                   "LED / Common-Area Electric Retrofit",
                   "Water Conservation: Fixtures", "Water Conservation: Plumbing"],
        notes="Component rows double count against this umbrella row — net out or pick one side. "
              "(List extends the tab's netting note, which omitted LED and water conservation.)"),
    "Reduce Insurance Expense": dict(
        trigger="T-12 insurance/unit exceeds comp average. Verify against loss runs / quotes.",
        source=["T12", "COMPS"], verify="docs",
        conflicts=["Reduce Opex to Comp Averages"],
        notes="Narrative row: achieve via the ASSUMPTIONS insurance override or 'agency' benchmark."),
    "Reduce Tax Expense / Protest": dict(
        trigger="Assessed value above market or T-12 taxes exceed comp averages.",
        source=["T12", "WEB"], verify="docs",
        conflicts=["Tax Exemption (HFC/PFC)", "Reduce Opex to Comp Averages"],
        notes="Narrative row: achieve via tax_assessment_factor. An HFC/PFC exemption supersedes a protest."),
    "LED / Common-Area Electric Retrofit": dict(
        trigger="T-12 owner-paid common-area electric exceeds comp average. Fixture type needs a site look.",
        source=["T12", "SITE"], verify="site",
        conflicts=["Reduce Opex to Comp Averages"]),
    "Water Submetering": dict(
        trigger="Property pays water AND jurisdiction permits true submetering (check per state — OK generally "
                "permits; TX under PUC rules). Plumbing configuration needs site check.",
        source=["T12", "LEGAL", "SITE"], verify="site",
        conflicts=["Implement/Increase Water RUBS"]),
    "Trash Compactor Conversion": dict(
        trigger="Units >= 200 AND T-12 refuse expense exceeds comp average. Placement needs site space/access.",
        source=["T12", "SITE"], verify="site",
        conflicts=["Reduce Opex to Comp Averages"]),
    "Revenue Mgmt / Screening": dict(
        trigger="T-12 bad debt > 2% of GPR.",
        source=["T12"], verify="docs", conflicts=[],
        notes="Distinct from lease-up: occupancy and bad debt are separate lines — both may be included."),
    "Payroll Optimization": dict(
        trigger="T-12 payroll/unit exceeds comp average. Do not staff below what the unit count requires.",
        source=["T12", "COMPS"], verify="docs",
        conflicts=["Reduce Opex to Comp Averages"]),
    "Utility Bill Audit": dict(
        trigger="ALWAYS include — contingency-based, no upfront cost. Model only a modest recovery.",
        source=["T12"], verify="docs",
        conflicts=["Reduce Opex to Comp Averages"]),
    "Tax Exemption (HFC/PFC)": dict(
        trigger="State is TEXAS only (no Oklahoma equivalent) AND an HFC/PFC structure is available to the buyer. "
                "FLAG FOR BUYER — never model as achieved without executed structure documents.",
        source=["WEB", "LEGAL"], verify="flag",
        conflicts=["Reduce Tax Expense / Protest"]),
}


def validate_value_add(items=None, rules=None, state=None):
    """Sanity-check the include flags against VALUE_ADD_RULES.

    Returns (errors, warnings). Errors are selections the tab logic forbids;
    warnings are selections that need evidence attached (site verification,
    double-count netting). Call before run_model(); nothing here mutates.
    """
    items = items if items is not None else VALUE_ADD_ITEMS
    rules = rules if rules is not None else VALUE_ADD_RULES
    errors, warnings = [], []
    included = [it for it in items if it[2]]
    names = {it[0] for it in included}

    counts = {}
    for (name, cat, inc, n, cost_pu, noi_pu_mo, effect) in included:
        counts[cat] = counts.get(cat, 0) + 1
        r = rules.get(name, {})
        for other in r.get("conflicts", []):
            if other in names and name < other:   # report each pair once
                pair = f"'{name}' + '{other}'"
                if other in ("Reduce Opex to Comp Averages",) or                    name in ("Reduce Opex to Comp Averages",):
                    warnings.append(f"Double-count risk: {pair} — net one out.")
                else:
                    errors.append(f"Mutually exclusive: {pair}.")
        if r.get("verify") == "site":
            warnings.append(f"'{name}' requires PHYSICAL ASSET VERIFICATION before income is modeled.")
        if r.get("verify") == "flag":
            warnings.append(f"'{name}' is flag-for-buyer only — confirm executed evidence (lease/LOI/structure docs).")
        if effect not in ("prereq", "assumption") and (not n or not noi_pu_mo):
            errors.append(f"'{name}' included with n={n}, noi/unit/mo={noi_pu_mo} — set real values first.")
        if name == "Tax Exemption (HFC/PFC)" and state and str(state).upper() not in ("TX", "TEXAS"):
            errors.append("HFC/PFC exemption is Texas-only; state is %s." % state)
    for cat, c in counts.items():
        if c > 3:
            errors.append(f"{c} '{cat}' items included — the model only counts the first THREE; trim to 3.")
    return errors, warnings

# ============================================================================
# 4. FACTORS TAB  (exit-cap build-up)
# ----------------------------------------------------------------------------
# terminal cap = MROUND( sale-comp cap
#                        + market_occ_at_reversion / 10000
#                        + sum(included adjustments, bps) / 10000 , 0.0025 )
# (replicates Assumptions!G58; the occupancy term is tiny but faithful)
# ============================================================================

FACTORS = {
    "sale_comp_cap_rate": 0.06660952380952381,   # 'Agency Loan-Sale Comps'!Z40
    "market_occ_at_reversion": 0.90,             # Factors!N17
    "avg_5yr_market_rent_growth": 0.0275,        # Factors!N16 (reference)
    "cap_rate_adjustments": [
        # (label, bps, included)
        ("Tertiary Location",            150, False),
        ("Poor demographics / low AMI",  100, False),
        ("Low Unit Count",                50, False),
        ("LURA",                         150, False),
        ("High Section 8 Concentration", 100, False),
        ("Short Term Rental Concentration", 100, False),
        ("High Insurance Costs",         150, False),
        ("Old Vintage (1980s or older)",  25, True),   # J23 = x
        ("Other Income below average",     0, True),   # auto-flag, 0 bps
        ("Contract Services high",         0, True),   # auto-flag, 0 bps
        ("Repairs & Maintenance high",     0, True),   # auto-flag, 0 bps
        ("Administration high",            0, True),   # auto-flag, 0 bps
    ],
}

# Reference outputs for validation (Westlake Apartments).
# Year-1 metrics were validated to ~0.000% against the 8/6/2026 filled-out
# workbook and are unchanged by the G59 economic-loss update. The out-year
# metrics (IRR, CoC, equity multiple, LP IRR, exit price) are this engine's
# values under the new G59 schedule (G59 = 8.0%) — re-validate them against
# the next filled-out copy of the updated workbook and paste in its cached
# values.
EXCEL_TARGETS = {
    "project_irr": 0.2337289397,        # engine-derived under G59 schedule
    "avg_cash_on_cash": 0.1308252417,   # engine-derived under G59 schedule
    "t3_dscr": 1.3076268612470674,      # Excel-validated (Y1 / trailing)
    "proforma_y1_dscr": 1.862208060526909,   # Excel-validated
    "equity_multiple": 2.4492772459,    # engine-derived under G59 schedule
    "lp_irr": 0.1876298753,             # engine-derived under G59 schedule
    "noi_year1": 984572.6792013333,     # Excel-validated
    "loan_amount": 8_625_000,           # Excel-validated
    "equity_required": 3_133_750,       # Excel-validated
    "terminal_cap": 0.07,               # Excel-validated
    "sale_price_at_exit": 14471437.2280,  # engine-derived under G59 schedule
}


# ============================================================================
# FINANCIAL PRIMITIVES (Excel-equivalent)
# ============================================================================

def pmt(rate: float, nper: int, pv: float) -> float:
    """Excel PMT (payment per period, negative for a positive pv)."""
    if rate == 0:
        return -pv / nper
    return -pv * rate / (1 - (1 + rate) ** -nper)


def cumprinc(rate: float, nper: int, pv: float, start: int, end: int) -> float:
    """Excel CUMPRINC: cumulative principal paid between periods start..end."""
    payment = pmt(rate, nper, pv)
    bal, total = pv, 0.0
    for per in range(1, end + 1):
        interest = bal * rate
        principal = payment + interest      # payment is negative
        if per >= start:
            total += principal
        bal += principal
    return total


def irr(cashflows, guess: float = 0.1) -> float:
    """Excel-style IRR (annual periods). Newton with bisection fallback."""
    if all(cf >= 0 for cf in cashflows) or all(cf <= 0 for cf in cashflows):
        return float("nan")            # IRR undefined without a sign change
    def npv(r):
        return sum(cf / (1 + r) ** i for i, cf in enumerate(cashflows))
    r = guess
    for _ in range(100):
        f = npv(r)
        h = 1e-7
        d = (npv(r + h) - f) / h
        if abs(d) < 1e-12:
            break
        r2 = r - f / d
        if abs(r2 - r) < 1e-10:
            return r2
        r = r2
    lo, hi = -0.9999, 10.0
    flo = npv(lo)
    for _ in range(200):
        mid = (lo + hi) / 2
        fmid = npv(mid)
        if abs(fmid) < 1e-9:
            return mid
        if (flo < 0) == (fmid < 0):
            lo, flo = mid, fmid
        else:
            hi = mid
    return (lo + hi) / 2


def mround(x: float, multiple: float) -> float:
    """Excel MROUND (round half away from zero to nearest multiple)."""
    return multiple * math.floor(x / multiple + 0.5)


# ============================================================================
# MODEL
# ============================================================================

def t12(series):            # trailing-12 total
    return sum(series)


def t3_annualized(series):  # last 3 months x 4
    return sum(series[-3:]) * 4


def econ_loss_schedule(a):
    """Out-year economic loss, 'UW - F&C' AQ50:AU50 (post-8/2026 model).

    Anchored at the stabilized/reversion input (Assumptions!G59 ->
    reversion_econ_loss) and stepped +50 bps per year walking back toward
    acquisition; year 2 averages the year-3 loss with the Year-1
    underwritten economic loss, plus 50 bps.
    Keys are absolute proforma years (2..6; year 6 = reversion column AU).
    """
    rev = a["reversion_econ_loss"]                                   # AU50
    y5 = rev + 0.005                                                 # AT50
    y4 = y5 + 0.005                                                  # AS50
    y3 = y4 + 0.005                                                  # AR50
    y1_loss = a["vacancy_pct"] + a["concessions_pct"] + a["bad_debt_pct"]
    y2 = (y3 + y1_loss) / 2 + 0.005                                  # AQ50
    return {2: y2, 3: y3, 4: y4, 5: y5, 6: rev}


def value_add_summary(items, units_total):
    """Replicates the Value-Add tab roll-ups (top 3 per category)."""
    out = {"other_income": 0.0, "renovation_noi": 0.0, "water_savings": 0.0,
           "rubs_water": 0.0, "rubs_electric": 0.0, "rubs_trash": 0.0,
           "electric_savings": 0.0, "trash_savings": 0.0,
           "payroll_savings": 0.0, "bad_debt_savings": 0.0,
           "tax_savings": 0.0, "other_opex_savings": 0.0,
           "budget": 0.0, "selected": []}
    counts = {"renovation": 0, "amenity": 0, "operations": 0}
    for (name, cat, inc, n, cost_pu, noi_pu_mo, effect) in items:
        if not inc or counts[cat] >= 3:
            continue
        counts[cat] += 1
        annual_noi = noi_pu_mo * n * 12
        total_cost = (cost_pu or 0) * n
        out["budget"] += total_cost
        out["selected"].append((name, cat, annual_noi, total_cost))
        if cat == "renovation" and effect == "rent":
            out["renovation_noi"] += annual_noi
        elif effect == "other_income":
            out["other_income"] += annual_noi
        elif effect in out and effect not in ("budget", "selected",
                                              "renovation_noi", "other_income"):
            out[effect] += annual_noi
        # 'prereq' and 'assumption' rows: cost only, no NOI wiring.
    return out


def terminal_cap_rate(factors, override=None):
    """Replicates Assumptions!G58 exit-cap build-up from the Factors tab."""
    if override is not None:
        return override
    bps = sum(b for (_, b, inc) in factors["cap_rate_adjustments"] if inc)
    raw = (factors["sale_comp_cap_rate"]
           + factors["market_occ_at_reversion"] / 10000
           + bps / 10000)
    return mround(raw, 0.0025)


def run_model(prop=PROPERTY, t12m=T12_MONTHLY, a=ASSUMPTIONS,
              va_items=VALUE_ADD_ITEMS, factors=FACTORS, verbose=False):
    units = prop["units"]
    price = a["purchase_price"]
    hold = a["hold_period_years"]

    va = value_add_summary(va_items, units)

    # ---------------- Year-1 underwriting ('UW - F&C' col AK) ----------------
    t12_market_rent = t12(t12m["market_rent"])
    gpr1 = t12_market_rent * (1 + a["year1_rent_growth"])          # AK5 (LTL=0 -> AK7)
    gpr1 -= gpr1 * a["loss_to_lease_pct"]
    econ_loss1 = a["vacancy_pct"] + a["concessions_pct"] + a["bad_debt_pct"]
    net_rental_income1 = gpr1 * (1 - econ_loss1) \
        + va["bad_debt_savings"]                                    # AK11 (+ VA bad-debt offset)

    rubs1 = {
        "rubs_electric": t12(t12m["rubs_electric"]) + va["rubs_electric"],
        "rubs_water":    t12(t12m["rubs_water"]) + va["rubs_water"],
        "rubs_trash":    t12(t12m["rubs_trash"]) + va["rubs_trash"],
    }
    rubs_income1 = sum(rubs1.values())                              # AK14

    if a["other_income_override"] is not None:
        other_income1 = a["other_income_override"]
    else:
        other_income1 = t12(t12m["other_income"]) + va["other_income"]  # AK18

    egi1 = net_rental_income1 + rubs_income1 + other_income1        # AK19

    def expense_line(key):
        setting = a[key]
        if setting == "agency":
            return AGENCY_BENCHMARKS[key] * units
        if setting is None:
            return t12(t12m[key])
        return float(setting)

    exp = {}
    for key in ("contract_services", "repairs_maintenance", "administration",
                "marketing", "payroll"):
        exp[key] = expense_line(key)
    if a["payroll"] is None:
        exp["payroll"] -= va["payroll_savings"]                     # VA payroll optimization
    controllable = sum(exp.values())                                # AK26

    electric = expense_line("util_electric")
    if a["util_electric"] is None:
        electric -= va["electric_savings"]                          # VA LED retrofit
    exp["util_electric"] = electric
    water = expense_line("util_water_sewer")
    if a["util_water_sewer"] is None:
        water -= va["water_savings"]                                # Assumptions!G34
    exp["util_water_sewer"] = water
    trash = expense_line("util_trash")
    if a["util_trash"] is None:
        trash -= va["trash_savings"]                                # VA compactor conversion
    exp["util_trash"] = trash
    exp["util_gas"] = expense_line("util_gas")
    utilities = (exp["util_electric"] + exp["util_water_sewer"]
                 + exp["util_trash"] + exp["util_gas"])             # AK31

    mgmt_fee1 = a["management_fee_pct"] * egi1                      # AK32
    insurance1 = expense_line("insurance")                          # AK33
    taxes1 = a["tax_assessment_factor"] * prop["total_tax_rate"] * price \
        - va["tax_savings"]                                         # AK34 (- VA exemption, if modeled)
    capex = a["capex_reserve_per_unit"] * units                     # AK36

    opex1 = (controllable + utilities + mgmt_fee1 + insurance1 + taxes1
             - va["other_opex_savings"])                            # AK35 (- VA utility audit etc.)
    total_exp1 = opex1 + capex                                      # AK37
    noi1 = egi1 - total_exp1                                        # AK38

    # ---------------- Debt ---------------------------------------------------
    loan = price * a["ltv"] + a["capex_financed_in_loan"] + a["supplemental_loan"]  # AL63
    rate = a["interest_rate"]
    io_ds = rate * (price * a["ltv"] + a["capex_financed_in_loan"])  # IO-period annual DS
    amort_pay_annual = -pmt(rate / 12, a["amortization_months"],
                            price * a["ltv"] + a["capex_financed_in_loan"]) * 12  # AL42
    loan_constant = -pmt(rate / 12, a["amortization_months"], 1) * 12             # AL64

    va_budget = (a["value_add_budget_override"]
                 if a["value_add_budget_override"] is not None else va["budget"])
    total_costs = price + va_budget + a["capex_financed_in_loan"] \
        + a["additional_lender_reserves"]                            # AI54
    equity = total_costs - loan + loan * a["loan_origination_pct"]   # AL72

    # ---------------- Multi-year proforma (years 1..hold, + reversion) -------
    years = list(range(1, hold + 1)) + ["reversion"]
    gpr, econ, rental, other, rubs_y, revenue = {}, {}, {}, {}, {}, {}
    gpr[1] = gpr1
    econ[1] = econ_loss1
    rental[1] = net_rental_income1
    other[1] = other_income1
    rubs_y[1] = rubs_income1
    econ_sched = econ_loss_schedule(a)
    for i, y in enumerate(years[1:], start=2):
        gpr[y] = gpr[years[i - 2]] * (1 + a["gpr_growth"])
        econ[y] = econ_sched[min(i, 6)]   # i = absolute proforma year
        rental[y] = gpr[y] * (1 - econ[y])
        other[y] = other[years[i - 2]] * (1 + a["other_income_growth"])
        rubs_y[y] = rubs_y[years[i - 2]] * (1 + a["other_income_growth"])
    for y in years:
        revenue[y] = rental[y] + other[y] + rubs_y[y]                # AP63..AU63

    growth_lines1 = {                       # expense lines that grow at expense_growth
        "contract_services": exp["contract_services"],
        "repairs_maintenance": exp["repairs_maintenance"],
        "administration": exp["administration"],
        "marketing": exp["marketing"],
        "payroll": exp["payroll"],
        "utilities": utilities,
        "management_fee": mgmt_fee1,
        "insurance": insurance1,
    }
    term_cap = terminal_cap_rate(factors, a["terminal_cap_override"])

    noi, opex_total = {1: noi1}, {1: total_exp1}
    for i, y in enumerate(years[1:], start=2):
        g = (1 + a["expense_growth"]) ** (i - 1)
        e = sum(v * g for v in growth_lines1.values())
        if y == "reversion":
            # Buyer reassessment at sale: 'UW - F&C' AU75
            taxes_y = (revenue["reversion"] * 0.5 / term_cap
                       * prop["total_tax_rate"] * a["tax_assessment_factor"])
        else:
            taxes_y = taxes1 * (1 + a["tax_growth"]) ** (i - 1)
        opex_total[y] = e + taxes_y + capex
        noi[y] = revenue[y] - opex_total[y]                          # AP78..AU78

    # ---------------- Reversion / sale ---------------------------------------
    sale_price = noi["reversion"] / term_cap                         # AL66
    sales_expense = sale_price * a["sales_expense_pct"]              # AL67
    amort_months_in_hold = max(0, (hold - a["io_years"]) * 12)       # AI57
    principal_paydown = -cumprinc(rate / 12, a["amortization_months"],
                                  price * a["ltv"] + a["capex_financed_in_loan"],
                                  1, amort_months_in_hold) if amort_months_in_hold else 0.0
    loan_payoff = loan - principal_paydown                           # AL68
    net_sale_proceeds = sale_price - sales_expense - loan_payoff     # AL69

    # ---------------- Annual cash flows & metrics ----------------------------
    debt_service, ncf, coc = {}, {}, {}
    for y in range(1, hold + 1):
        debt_service[y] = io_ds if y <= a["io_years"] else amort_pay_annual
        ncf[y] = noi[y] - debt_service[y]                            # AP82..
        coc[y] = ncf[y] / equity                                     # AP83..
    project_cfs = [-equity]
    for y in range(1, hold + 1):
        cf = ncf[y]
        if y == 1:
            cf += a["additional_lender_reserves"]                    # returned yr 1
        if y == hold:
            cf += net_sale_proceeds
        project_cfs.append(cf)

    project_irr = irr(project_cfs)                                   # AI51
    avg_coc = sum(coc.values()) / hold                               # AY71
    equity_multiple = sum(project_cfs[1:]) / equity                  # AY70
    annualized_roi = (sum(project_cfs) / equity) / hold              # AY69 (ROI / hold)

    # DSCRs
    t3_egi = (t3_annualized(t12m["market_rent"]) + t3_annualized(t12m["loss_to_lease"])
              + t3_annualized(t12m["vacancy"]) + t3_annualized(t12m["concessions"])
              + t3_annualized(t12m["bad_debt"]) + t3_annualized(t12m["other_income"])
              + t3_annualized(t12m["rubs_electric"]) + t3_annualized(t12m["rubs_water"])
              + t3_annualized(t12m["rubs_trash"]))                   # AB19
    t12_utilities = (t12(t12m["util_electric"]) + t12(t12m["util_water_sewer"])
                     + t12(t12m["util_trash"]) + t12(t12m["util_gas"]))
    # T-3 normalized NOI ('UW - F&C' AB41): T-3 income against proforma
    # expenses, swapping proforma utilities for trailing actual utilities.
    t3_noi_normalized = t3_egi - total_exp1 + utilities - t12_utilities
    t3_dscr = t3_noi_normalized / amort_pay_annual                   # AC42
    y1_dscr = noi1 / debt_service[1]                                 # AP84

    # ---------------- LP waterfall (rows 86-105) -----------------------------
    lp_equity = equity - a["gp_coinvest"]
    lp_cfs = [-lp_equity]
    for y in range(1, hold + 1):
        amf = revenue[y] * a["asset_mgmt_fee_pct"]
        pref = a["preferred_return"] * lp_equity
        residual = ncf[y] - amf - pref
        lp_cf = pref + (a["lp_share"] * residual if residual > 0 else residual)
        if y == hold:
            potential_sale = noi[hold] / term_cap                    # AT97 (hold-yr NOI)
            gross_profit = (potential_sale
                            - potential_sale * a["sales_expense_pct"]
                            - loan_payoff)
            net_profit = gross_profit - equity
            lp_cf += lp_equity + a["lp_share"] * net_profit
        lp_cfs.append(lp_cf)
    lp_irr = irr(lp_cfs)

    results = {
        "value_add": va,
        "terminal_cap": term_cap,
        "gpr_year1": gpr1, "egi_year1": egi1, "noi": noi,
        "revenue": revenue, "opex_total": opex_total,
        "loan_amount": loan, "loan_constant": loan_constant,
        "annual_ds_io": io_ds, "annual_ds_amortizing": amort_pay_annual,
        "monthly_pmt_post_io": amort_pay_annual / 12,
        "total_costs": total_costs, "equity_required": equity,
        "sale_price_at_exit": sale_price, "loan_payoff": loan_payoff,
        "net_sale_proceeds": net_sale_proceeds,
        "project_cashflows": project_cfs, "cash_on_cash": coc,
        "project_irr": project_irr, "avg_cash_on_cash": avg_coc,
        "equity_multiple": equity_multiple, "roi": equity_multiple - 1,
        "annualized_roi": annualized_roi,
        "t3_dscr": t3_dscr, "proforma_y1_dscr": y1_dscr,
        "debt_yield_y1": noi1 / loan,
        "lp_cashflows": lp_cfs, "lp_irr": lp_irr,
    }
    return results


def solve_price_for_irr(target_irr: float, lo=1e6, hi=50e6, tol=1e-6) -> float:
    """Goal-seek purchase price to a target project IRR (mirrors the model's
    low/high pricing bands, which solve price against the buyers' target IRR)."""
    a = dict(ASSUMPTIONS)
    def f(p):
        a["purchase_price"] = p
        v = run_model(a=a)["project_irr"]
        if math.isnan(v):
            v = -1.0                   # undefined IRR -> treat as deeply below target
        return v - target_irr
    flo, fhi = f(lo), f(hi)
    for _ in range(80):
        mid = (lo + hi) / 2
        fm = f(mid)
        if abs(fm) < tol:
            return mid
        if (flo < 0) == (fm < 0):
            lo, flo = mid, fm
        else:
            hi, fhi = mid, fm
    return (lo + hi) / 2


# ============================================================================
# REPORT + VALIDATION
# ============================================================================

def main():
    r = run_model()
    hold = ASSUMPTIONS["hold_period_years"]
    money = lambda x: f"${x:,.0f}"
    pct = lambda x: f"{x*100:.2f}%"

    print("=" * 74)
    print(f"{PROPERTY['name']} — {PROPERTY['units']} units — Free & Clear scenario")
    print("=" * 74)
    print(f"Purchase Price          {money(ASSUMPTIONS['purchase_price'])}"
          f"    ({money(ASSUMPTIONS['purchase_price']/PROPERTY['units'])}/unit)")
    print(f"Loan ({pct(ASSUMPTIONS['ltv'])} LTV)        {money(r['loan_amount'])}"
          f"   @ {pct(ASSUMPTIONS['interest_rate'])}, {ASSUMPTIONS['io_years']}y IO,"
          f" {ASSUMPTIONS['amortization_months']}mo am")
    print(f"Equity Required         {money(r['equity_required'])}")
    print(f"Terminal Cap (Factors)  {pct(r['terminal_cap'])}")
    va = r["value_add"]
    print(f"Value-Add other income  {money(va['other_income'])}/yr "
          f"({', '.join(s[0] for s in va['selected'])})")
    print("-" * 74)
    print(f"{'Year':<12}{'Revenue':>14}{'Expenses':>14}{'NOI':>14}"
          f"{'Debt Svc':>14}{'CF':>14}")
    for y in list(range(1, hold + 1)) + ["reversion"]:
        ds = "" if y == "reversion" else money(-(r['annual_ds_io'] if y <= ASSUMPTIONS['io_years'] else r['annual_ds_amortizing']))
        cf = "" if y == "reversion" else money(r['noi'][y] - (r['annual_ds_io'] if y <= ASSUMPTIONS['io_years'] else r['annual_ds_amortizing']))
        print(f"{str(y):<12}{money(r['revenue'][y]):>14}{money(-r['opex_total'][y]):>14}"
              f"{money(r['noi'][y]):>14}{ds:>14}{cf:>14}")
    print("-" * 74)
    print(f"Sale @ exit  {money(r['sale_price_at_exit'])}   less sales exp & payoff"
          f" -> net proceeds {money(r['net_sale_proceeds'])}")
    print(f"Project cash flows: {[round(cf) for cf in r['project_cashflows']]}")
    print("=" * 74)
    print("RETURN METRICS")
    print(f"  Project IRR            {pct(r['project_irr'])}")
    print(f"  Avg Cash-on-Cash       {pct(r['avg_cash_on_cash'])}")
    print(f"  Equity Multiple        {r['equity_multiple']:.2f}x")
    print(f"  Annualized ROI         {pct(r['annualized_roi'])}")
    print(f"  T-3 DSCR               {r['t3_dscr']:.4f}")
    print(f"  Pro Forma Yr-1 DSCR    {r['proforma_y1_dscr']:.4f}")
    print(f"  Debt Yield (Yr 1)      {pct(r['debt_yield_y1'])}")
    print(f"  LP IRR                 {pct(r['lp_irr'])}")
    print("=" * 74)
    print("VALIDATION vs EXCEL MODEL (tolerance +/- 2%)")
    checks = [
        ("Project IRR", "project_irr"),
        ("Avg Cash-on-Cash", "avg_cash_on_cash"),
        ("T-3 DSCR", "t3_dscr"),
        ("Pro Forma Yr-1 DSCR", "proforma_y1_dscr"),
        ("Equity Multiple", "equity_multiple"),
        ("LP IRR", "lp_irr"),
        ("Year-1 NOI", "noi_year1"),
        ("Loan Amount", "loan_amount"),
        ("Equity Required", "equity_required"),
        ("Terminal Cap", "terminal_cap"),
        ("Sale Price at Exit", "sale_price_at_exit"),
    ]
    all_ok = True
    for label, key in checks:
        excel = EXCEL_TARGETS[key]
        pyv = r["noi"][1] if key == "noi_year1" else r[key]
        diff = (pyv - excel) / excel if excel else 0.0
        ok = abs(diff) <= 0.02
        all_ok &= ok
        print(f"  {label:<22} python={pyv:>16,.6f}  excel={excel:>16,.6f}"
              f"  diff={diff*100:>+7.3f}%  {'PASS' if ok else 'FAIL'}")
    print("-" * 74)
    print("ALL TARGET METRICS WITHIN +/-2%" if all_ok else "*** MISMATCH — SEE ABOVE ***")
    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
