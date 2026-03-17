"""
Test: Build inputs from Ember_Template.xlsx assumptions, run calculate(), compare to Excel expected outputs.
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
from calc import calculate

# ── Build the complete inputs dict matching Excel template assumptions ──

inputs = {
    # Tract Inputs
    "project_name": "Test Project",
    "address": "TBD Address",
    "gross_acreage": 800,
    "purchase_price_per_acre": 45000,
    "closing_costs_pct": 0.045,
    "land_escalator": 0.05,
    "closing_date": "2026-06-30",
    "det_storage_rate": 1.1,
    "det_depth": 9,
    "det_num_projects": 6,
    "parks_pct": 0.03,
    "drill_site_acres": 2,
    "commercial_pod_acres": 30,
    "residential_pod_acres": 60,

    # Plant facilities (Tract Inputs)
    "plants": [
        {"type": "Water Plant", "acres": 3.5},
        {"type": "Lift Station", "acres": 0.75},
        {"type": "Lift Station", "acres": 0.75},
        {"type": "WWTP", "acres": 10},
        {"type": "Lift Station", "acres": 0.75},
        {"type": "None", "acres": 0},
        {"type": "None", "acres": 0},
        {"type": "None", "acres": 0},
    ],

    # Amenities (Tract Inputs)
    "amenities": [
        {"type": "Large Amenity Center", "acres": 6},
        {"type": "Small Amenity Center", "acres": 3},
        {"type": "Pocket Park", "acres": 0.5},
        {"type": "Pocket Park", "acres": 0.5},
        {"type": "Pocket Park", "acres": 0.5},
        {"type": "Small Amenity Center", "acres": 3},
    ],

    # Other net-outs
    "other_netouts": [
        {"name": "School Site", "acres": 15},
    ],

    # Roads (Tract Inputs)
    "roads": [
        {"type": "2 Lane", "lf": 1500, "width": 60, "road_setback": 35, "landscaping_setback": 25},
        {"type": "4 Lane", "lf": 2000, "width": 80, "road_setback": 35, "landscaping_setback": 25},
        {"type": "4 Lane", "lf": 2000, "width": 80, "road_setback": 35, "landscaping_setback": 25},
        {"type": "2 Lane", "lf": 1500, "width": 60, "road_setback": 35, "landscaping_setback": 25},
        {"type": "2 Lane", "lf": 2000, "width": 80, "road_setback": 35, "landscaping_setback": 25},
    ],

    # ── Cost Inputs ──
    "default_other_pct": 0.17,
    "sectional_other_pct": 0.17,
    "landscaping_other_pct": 0.12,
    "contingency": 0.05,
    "site_work_pct": 0.01,
    "fenced_pct": 0.25,
    "cost_per_mailbox": 200,
    "cost_per_streetlight": 1700,
    "default_start_month": 1,

    # Land takedowns
    "takedowns": [
        {"period": 0, "pct": 0.5, "purchase_price": 18000000, "closing_costs": 810000, "total": 18810000},
        {"period": 36, "pct": 0.5, "purchase_price": 20837250, "closing_costs": 810000, "total": 21647250},
        {"period": 48, "pct": 0, "purchase_price": 0, "closing_costs": 0, "total": 0},
    ],

    # Plant costs
    "plant_costs": [
        {"type": "Water Plant", "acres": 3.5, "base_cost": 6000000, "other_pct": 0.17, "total_cost": 7020000, "start_month": 1, "duration": 8, "ph2_base_cost": 3000000, "ph2_other_pct": 0.17, "phase2_total": 3510000, "ph2_start_month": 37, "ph2_duration": 8},
        {"type": "Lift Station", "acres": 0.75, "base_cost": 1500000, "other_pct": 0.17, "total_cost": 1755000, "start_month": 1, "duration": 3, "ph2_base_cost": 0, "ph2_other_pct": 0.17, "phase2_total": 0, "ph2_start_month": 37, "ph2_duration": 3},
        {"type": "Lift Station", "acres": 0.75, "base_cost": 1500000, "other_pct": 0.17, "total_cost": 1755000, "start_month": 24, "duration": 3, "ph2_base_cost": 0, "ph2_other_pct": 0.17, "phase2_total": 0, "ph2_start_month": 60, "ph2_duration": 3},
        {"type": "WWTP", "acres": 10, "base_cost": 9000000, "other_pct": 0.17, "total_cost": 10530000, "start_month": 1, "duration": 8, "ph2_base_cost": 4000000, "ph2_other_pct": 0.17, "phase2_total": 4680000, "ph2_start_month": 37, "ph2_duration": 8},
        {"type": "Lift Station", "acres": 0.75, "base_cost": 1500000, "other_pct": 0.17, "total_cost": 1755000, "start_month": 48, "duration": 3, "ph2_base_cost": 0, "ph2_other_pct": 0.17, "phase2_total": 0, "ph2_start_month": 84, "ph2_duration": 3},
    ],

    # Amenity costs
    "amenity_costs": [
        {"type": "Large Amenity Center", "acres": 6, "base_cost": 8000000, "other_pct": 0.17, "total_cost": 9360000, "start_month": 12, "duration": 12},
        {"type": "Small Amenity Center", "acres": 3, "base_cost": 5000000, "other_pct": 0.17, "total_cost": 5850000, "start_month": 60, "duration": 8},
        {"type": "Pocket Park", "acres": 0.5, "base_cost": 500000, "other_pct": 0.17, "total_cost": 585000, "start_month": 80, "duration": 2},
        {"type": "Pocket Park", "acres": 0.5, "base_cost": 500000, "other_pct": 0.17, "total_cost": 585000, "start_month": 100, "duration": 2},
        {"type": "Pocket Park", "acres": 0.5, "base_cost": 500000, "other_pct": 0.17, "total_cost": 585000, "start_month": 120, "duration": 2},
        {"type": "Small Amenity Center", "acres": 3, "base_cost": 4000000, "other_pct": 0.17, "total_cost": 4680000, "start_month": 100, "duration": 8},
    ],

    # Detention costs
    "det_costs": [
        {"acres": 21.185, "base_cost": 2366222, "other_pct": 0.17, "total_cost": 2768480, "start_month": 1, "duration": 9, "delivery_period": 9, "landscaping_per_sf": 2, "total_landscaping": 620140},
        {"acres": 21.185, "base_cost": 2366222, "other_pct": 0.17, "total_cost": 2768480, "start_month": 16, "duration": 9, "delivery_period": 24, "landscaping_per_sf": 2, "total_landscaping": 620140},
        {"acres": 21.185, "base_cost": 2366222, "other_pct": 0.17, "total_cost": 2768480, "start_month": 31, "duration": 9, "delivery_period": 39, "landscaping_per_sf": 2, "total_landscaping": 620140},
        {"acres": 21.185, "base_cost": 2366222, "other_pct": 0.17, "total_cost": 2768480, "start_month": 46, "duration": 9, "delivery_period": 54, "landscaping_per_sf": 2, "total_landscaping": 620140},
        {"acres": 21.185, "base_cost": 2366222, "other_pct": 0.17, "total_cost": 2768480, "start_month": 61, "duration": 9, "delivery_period": 69, "landscaping_per_sf": 2, "total_landscaping": 620140},
        {"acres": 21.185, "base_cost": 2366222, "other_pct": 0.17, "total_cost": 2768480, "start_month": 76, "duration": 9, "delivery_period": 84, "landscaping_per_sf": 2, "total_landscaping": 620140},
    ],

    # Other costs
    "other_costs": [
        {"name": "School Site", "acres": 15, "base_cost": 15000, "other_pct": 0.17, "total_cost": 17550, "start_month": 1, "duration": 3},
    ],

    # Road costs
    "road_costs": [
        {"type": "2 Lane", "lf": 1500, "wsd_per_lf": 450, "paving_per_lf": 343, "base_cost": 1189500, "other_pct": 0.17, "total_cost": 1391715, "start_month": 1, "duration": 11, "delivery_period": 11, "landscaping_per_sf": 2, "total_landscaping": 168000, "light_spacing": 250, "streetlight_cost": 20400},
        {"type": "4 Lane", "lf": 2000, "wsd_per_lf": 460, "paving_per_lf": 663, "base_cost": 2246000, "other_pct": 0.17, "total_cost": 2627820, "start_month": 12, "duration": 13, "delivery_period": 24, "landscaping_per_sf": 2, "total_landscaping": 224000, "light_spacing": 250, "streetlight_cost": 27200},
        {"type": "4 Lane", "lf": 2000, "wsd_per_lf": 460, "paving_per_lf": 663, "base_cost": 2246000, "other_pct": 0.17, "total_cost": 2627820, "start_month": 48, "duration": 13, "delivery_period": 60, "landscaping_per_sf": 2, "total_landscaping": 224000, "light_spacing": 250, "streetlight_cost": 27200},
        {"type": "2 Lane", "lf": 1500, "wsd_per_lf": 450, "paving_per_lf": 343, "base_cost": 1189500, "other_pct": 0.17, "total_cost": 1391715, "start_month": 72, "duration": 11, "delivery_period": 82, "landscaping_per_sf": 2, "total_landscaping": 168000, "light_spacing": 250, "streetlight_cost": 20400},
        {"type": "2 Lane", "lf": 2000, "wsd_per_lf": 450, "paving_per_lf": 343, "base_cost": 1586000, "other_pct": 0.17, "total_cost": 1855620, "start_month": 96, "duration": 13, "delivery_period": 108, "landscaping_per_sf": 2, "total_landscaping": 224000, "light_spacing": 250, "streetlight_cost": 27200},
    ],

    # Lot sizes (sectional development)
    "lot_sizes": [
        {"ff": 25, "on": 0, "yield_per_ac": 8.25, "pace": 5, "home_price": 200000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 4, "fence_cost_per_ff": 94,
         "build_time": 3, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 2000, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 30, "on": 0, "yield_per_ac": 5.54, "pace": 5, "home_price": 360000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 4, "fence_cost_per_ff": 94,
         "build_time": 3, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 3600, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 35, "on": 0, "yield_per_ac": 8.25, "pace": 6, "home_price": 275000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 4, "fence_cost_per_ff": 94,
         "build_time": 3, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 2800, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 40, "on": 1, "yield_per_ac": 5.5, "pace": 7, "home_price": 330168, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 4, "fence_cost_per_ff": 94,
         "build_time": 3, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 3300, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 45, "on": 1, "yield_per_ac": 5, "pace": 6, "home_price": 380000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 4, "fence_cost_per_ff": 94,
         "build_time": 4, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 3800, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 50, "on": 1, "yield_per_ac": 4.5, "pace": 5, "home_price": 430000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 4, "fence_cost_per_ff": 94,
         "build_time": 4, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 4300, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 55, "on": 0, "yield_per_ac": 4, "pace": 5, "home_price": 500000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 4, "fence_cost_per_ff": 94,
         "build_time": 4, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 5000, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 60, "on": 1, "yield_per_ac": 3.5, "pace": 2, "home_price": 580000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 4, "fence_cost_per_ff": 94,
         "build_time": 5, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 5800, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 65, "on": 0, "yield_per_ac": 3, "pace": 2, "home_price": 615000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 3, "fence_cost_per_ff": 94,
         "build_time": 5, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 6200, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 70, "on": 0, "yield_per_ac": 2.5, "pace": 1, "home_price": 675000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 3, "fence_cost_per_ff": 94,
         "build_time": 6, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 6800, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 75, "on": 0, "yield_per_ac": 2, "pace": 1, "home_price": 720000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 3, "fence_cost_per_ff": 94,
         "build_time": 6, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 7200, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
        {"ff": 80, "on": 1, "yield_per_ac": 1.5, "pace": 0.75, "home_price": 750000, "wsd_per_ff": 290, "paving_per_ff": 220, "dev_start_month": 1, "landscaping_per_lot": 2000, "urd_per_lot": 35, "lots_per_streetlight": 3, "fence_cost_per_ff": 94,
         "build_time": 7, "av_pct": 0.85, "premium_per_ff": 25, "escalation": 0.06, "fence_per_ff": 65, "marketing_fee": 7500, "lot_av_pct": 0.5, "lot_tax_rate": 0.022},
    ],

    # DMF
    "dmf_pct": 0.025,

    # Professional services
    "prof_svc_pct": 0.015,

    # Operating costs
    "personnel_monthly": 50000,
    "marketing_personnel_monthly": 15000,
    "legal_monthly": 10000,
    "mud_monthly": 35000,
    "mud_pct": 0.20,
    "insurance_monthly": 10000,
    "bookkeeping_monthly": 10000,

    # ── Revenue Inputs ──
    "timing_method": "50/25/25",
    "bem_period": 9,
    "bem_pct": 0.18,
    "brokerage_fees": 0.03,
    "lot_closing_costs": 0.015,
    "take1_pct": 0.5,
    "take2_pct": 0.25,
    "take3_pct": 0.25,

    # $/FF by year (flat $1800 for years 0-10)
    "price_per_ff": {str(yr): 1800 for yr in range(11)},

    # Residential pods
    "res_pod_count": 1,
    "res_pods": [
        {"pod": 1, "price_per_acre": 120000, "closing_costs_pct": 0.045, "implied_lots_per_acre": 3.5, "impact_fee_per_lot": 10000, "sale_period": 12},
        {"pod": 2, "price_per_acre": 120000, "closing_costs_pct": 0.045, "implied_lots_per_acre": 3.5, "impact_fee_per_lot": 10000, "sale_period": 36},
        {"pod": 3, "price_per_acre": 120000, "closing_costs_pct": 0.045, "implied_lots_per_acre": 3.5, "impact_fee_per_lot": 10000, "sale_period": 60},
        {"pod": 4, "price_per_acre": 120000, "closing_costs_pct": 0.045, "implied_lots_per_acre": 3.5, "impact_fee_per_lot": 10000, "sale_period": 84},
        {"pod": 5, "price_per_acre": 120000, "closing_costs_pct": 0.045, "implied_lots_per_acre": 3.5, "impact_fee_per_lot": 10000, "sale_period": 108},
        {"pod": 6, "price_per_acre": 120000, "closing_costs_pct": 0.045, "implied_lots_per_acre": 3.5, "impact_fee_per_lot": 10000, "sale_period": 132},
    ],

    # Commercial pods
    "comm_pod_count": 6,
    "comm_pods": [
        {"pod": 1, "price_per_sf": 8, "closing_costs_pct": 0.045, "sale_period": 12, "av_per_acre": 1200000, "av_delay_months": 18},
        {"pod": 2, "price_per_sf": 8, "closing_costs_pct": 0.045, "sale_period": 36, "av_per_acre": 1200000, "av_delay_months": 18},
        {"pod": 3, "price_per_sf": 8, "closing_costs_pct": 0.045, "sale_period": 60, "av_per_acre": 1200000, "av_delay_months": 18},
        {"pod": 4, "price_per_sf": 8, "closing_costs_pct": 0.045, "sale_period": 84, "av_per_acre": 1200000, "av_delay_months": 18},
        {"pod": 5, "price_per_sf": 8, "closing_costs_pct": 0.045, "sale_period": 108, "av_per_acre": 1200000, "av_delay_months": 18},
        {"pod": 6, "price_per_sf": 8, "closing_costs_pct": 0.045, "sale_period": 132, "av_per_acre": 1200000, "av_delay_months": 18},
    ],

    # MUD & WCID bonds
    "mud_bond": {
        "toggle": 1,
        "debt_ratio": 0.12,
        "first_bond_period": 48,
        "bond_interval": 12,
        "pct_to_dev": 0.85,
        "receivables_fee": 0.025,
    },
    "wcid_bond": {
        "toggle": 1,
        "debt_ratio": 0.042,
        "first_bond_period": 48,
        "bond_interval": 12,
        "pct_to_dev": 0.85,
        "receivables_fee": 0.025,
    },

    # Lookups
    "lk_plants": [
        {"type": "WWTP", "acres": 10, "duration": 8},
        {"type": "Water Plant", "acres": 3.5, "duration": 8},
        {"type": "Lift Station", "acres": 0.75, "duration": 3},
        {"type": "None", "acres": 0, "duration": 0},
    ],
    "lk_amenities": [
        {"type": "Pocket Park", "duration": 2},
        {"type": "Small Amenity Center", "duration": 8},
        {"type": "Large Amenity Center", "duration": 12},
        {"type": "None", "duration": 0},
    ],
    "lk_roads": [
        {"type": "2 Lane", "wsd": 450, "paving": 343},
        {"type": "4 Lane", "wsd": 460, "paving": 663},
    ],
}

# ── Run calculation ──
out = calculate(inputs)

# ── Excel expected values (from Project Performance tab) ──
expected = {
    # Revenue
    "rev_lot_sales":     178875000,
    "rev_mud":           78079200,
    "rev_wcid":          27327720,
    "rev_premiums":      2484375,
    "rev_fence_fees":    1614844,
    "rev_escalations":   2750203,
    "rev_mktg_fees":     8580950,
    "rev_res_pods":      8976000,
    "rev_comm_pods":     9983952,
    "total_revenue":     318672214,

    # Costs
    "cost_land":         40457250,
    "cost_plants":       31005000,
    "cost_amenities":    21645000,
    "cost_detention":    16610880,
    "cost_lot_dev":      59297062,   # "Sections" in Excel
    "cost_other":        17550,
    "cost_roads":        9894690,
    "cost_fencing":      2335312,
    "cost_dry_utilities":1101166,
    "cost_site_work":    1788750,
    "cost_lot_landscaping": 9442917,  # "Landscaping" in Excel row 27
    "cost_legal":        1190000,
    "cost_lot_taxes":    1967625,
    "cost_mud_hoa":      875000,
    "cost_insurance":    1290000,
    "cost_marketing":    8580950,
    "cost_brokerage":    5366250,
    "cost_closing":      2683125,
    "cost_mailboxes":    420900,
    "cost_prof_svc":     4780083,
    "cost_contingency":  7895921,
    "gross_costs":       228645432,

    # Margins
    "gross_margin_amt":  90026782,
    "below_line_dmf":    4053187,
    "below_line_personnel": 7885000,
    "below_line_bookkeeping": 1290000,
    "net_margin_amt":    74163114,
}

# ── Compare ──
print("=" * 80)
print("EMBER TEMPLATE VERIFICATION: calc.py vs Excel")
print("=" * 80)

pass_count = 0
fail_count = 0
total = len(expected)

for key, excel_val in expected.items():
    calc_val = out.get(key, "MISSING")
    if calc_val == "MISSING":
        print(f"  MISSING  {key:30s}  Excel={excel_val:>15,}")
        fail_count += 1
        continue

    diff = abs(calc_val - excel_val)
    pct = (diff / abs(excel_val) * 100) if excel_val != 0 else 0
    tolerance = 0.5  # 0.5% tolerance for rounding
    status = "PASS" if pct <= tolerance else "FAIL"

    if status == "PASS":
        pass_count += 1
    else:
        fail_count += 1

    marker = "  " if status == "PASS" else ">>"
    print(f"{marker} {status:4s}  {key:30s}  Calc={calc_val:>15,}  Excel={excel_val:>15,}  Diff={diff:>10,}  ({pct:.2f}%)")

print("=" * 80)
print(f"RESULT: {pass_count}/{total} passed, {fail_count}/{total} failed")
print("=" * 80)

# Also print key KPIs and debug info
print(f"\nTotal Lots: {out.get('total_lots', '?')} (Excel: 2,104.5)")
print(f"Project Length: {out.get('project_length_months', '?')} months ({out.get('project_length_years', '?')} years)")
print(f"IRR: {out.get('unlevered_irr', 0)*100:.1f}%")
print(f"Gross Margin %: {out.get('gross_margin_pct', 0)*100:.1f}%")
print(f"Dev Acres: {out.get('dev_acres', '?')} / Res Dev Acres: {out.get('residential_dev_acres', '?')}")

# Debug lot rows
print("\nLot breakdown:")
for lr in out.get('lot_rows', []):
    if lr.get('total_lots', 0) > 0:
        print(f"  FF={lr['ff']:3.0f}  lots={lr['total_lots']:7.1f}  full_secs={lr['full_sections']}  last={lr['last_lots']}")

# Debug land cost
print(f"\ncost_land output: {out.get('cost_land', '?')} (this is base, not escalated)")
print(f"land_total_cost (with escalation): {out.get('land_total_cost', '?')}")

# Debug plant cost
print(f"\ncost_plants: {out.get('cost_plants', '?')}")
for pr in out.get('plant_cost_rows', []):
    if pr['total_cost'] > 0 or pr['ph2_total_cost'] > 0:
        print(f"  {pr['type']:15s}  ph1={pr['total_cost']:>12,.0f}  ph2={pr['ph2_total_cost']:>12,.0f}")

# Debug marketing
print(f"\ncost_marketing: {out.get('cost_marketing', '?')} (calc.py: mkt_pct * lot_rev)")
print(f"rev_mktg_fees:  {out.get('rev_mktg_fees', '?')} (sum of per-lot mktg fees)")
