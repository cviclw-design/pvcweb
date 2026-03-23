#!/usr/bin/env python3
import pandas as pd
import logging
import math
import os
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.formatting.rule import CellIsRule

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "downloads")
OUTPUT_DIR = os.path.join(BASE_DIR, "downloads")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "pvc_report.xlsx")
INPUT_FILE = os.path.join(DATA_DIR, "final_pvc_dataset.xlsx")
IEEMA_FILE = os.path.join(BASE_DIR, "IEEMA.xlsx")

GST_FACTOR = 1.18

DATE_COLS = [
    "pvc_base_date",
    "call_date",
    "scheduled_date",
    "sup_date",
    "orig_dp",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
)

logger = logging.getLogger("PVC")


def safe_float(x):
    try:
        return float(x)
    except Exception:
        return 0.0


def safe_round(x, n=2):
    try:
        return round(float(x), n)
    except Exception:
        return None


def truncate_4(x):
    try:
        return math.trunc(float(x) * 100) / 100
    except Exception:
        return None


def to_month_start(d):
    d = pd.to_datetime(d, errors="coerce")
    if pd.isna(d):
        return pd.NaT
    return pd.Timestamp(d.year, d.month, 1)


def previous_month(d):
    d = to_month_start(d)
    if pd.isna(d):
        return pd.NaT
    return d - relativedelta(months=1)


# =====================================================
# IEEMA HANDLING
# =====================================================
WPI_COEFF = {
    "copper": 40,
    "crgo": 24,
    "ms": 8,
    "insmat": 4,
    "transoil": 8,
    "wpi": 8,
}


def load_ieema():
    df = pd.read_excel(IEEMA_FILE)
    df.columns = df.columns.str.lower().str.strip()

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    else:
        df["date"] = pd.to_datetime(
            df["year"].astype(str) + "-" + df["month"].astype(str) + "-01",
            errors="coerce",
        )

    df["date"] = df["date"].apply(to_month_start)
    for col in WPI_COEFF.keys():
        if col in df.columns:
            df[col] = df[col].astype(float).round(2)
    return df.set_index("date").sort_index()


def get_ieema_df():
    return load_ieema()


def ieema_row(df, date, previous=False):
    if pd.isna(date):
        return None
    target = previous_month(date) if previous else to_month_start(date)
    eligible = df[df.index <= target]
    return eligible.iloc[-1] if not eligible.empty else None


def pvc_percent(base_date, current_date, ieema_df):
    base = ieema_row(ieema_df, base_date, previous=False)
    curr = ieema_row(ieema_df, current_date, previous=True)

    if base is None or curr is None:
        return None

    total = 0.0
    for k, w in WPI_COEFF.items():
        b = base.get(k)
        c = curr.get(k)
        if b and c:
            contrib = round(w * ((c - b) / b))
            total += contrib

    return round(total, 4)


def pvc_percent_detailed(base_date, current_date, ieema_df, scenario):
    base = ieema_row(ieema_df, base_date, previous=False)
    curr = ieema_row(ieema_df, current_date, previous=True)

    if base is None or curr is None:
        return None

    row = {
        "scenario": scenario,
        "base_month": base.name,
        "current_month": curr.name,
    }

    total = 0.0
    for k, w in WPI_COEFF.items():
        b = base.get(k)
        c = curr.get(k)
        if b and c:
            contrib = round(w * ((c - b) / b), 4)
            total += contrib
        else:
            contrib = None

        row[f"{k}_base"] = round(b,2) if b is not None else None
        row[f"{k}_current"] = round(c,2) if c is not None else None
        row[f"{k}_weight"] = w
        row[f"{k}_contribution_pct"] = contrib

    row["pvc_percent"] = round(total,4)
    return row


# =====================================================
# SINGLE RECORD CALCULATION FOR FLASK
# =====================================================
def calculate_single_record_from_dict(d, ieema_df):
    """
    Calculate PVC & LD for ONE record for the web app.
    Keys expected in d:
      acc_qty, basic_rate, freight_rate_per_unit,
      pvc_base_date, call_date,
      orig_dp, refixeddp, extendeddp, sup_date,
      lower_rate, lower_freight, lower_basic_date
    """

    pvc_base_date = pd.to_datetime(d.get("pvc_base_date"), errors="coerce")
    call_date = pd.to_datetime(d.get("call_date"), errors="coerce")

    orig_dp = pd.to_datetime(d.get("orig_dp"), errors="coerce")
    refixed_dp = pd.to_datetime(d.get("refixeddp"), errors="coerce")
    extended_dp = pd.to_datetime(d.get("extendeddp"), errors="coerce")

    # scheduled_date used for contractual PVC
    scheduled_date = refixed_dp if pd.notna(refixed_dp) else orig_dp

    sup_date = pd.to_datetime(d.get("sup_date"), errors="coerce")
    lower_basic_date = pd.to_datetime(d.get("lower_basic_date"), errors="coerce")

    acc_qty = safe_float(d.get("acc_qty"))
    basic_rate = safe_float(d.get("basic_rate"))
    freight_rate = safe_float(d.get("freight_rate_per_unit"))

    lower_rate = safe_float(d.get("lower_rate"))
    lower_freight = safe_float(d.get("lower_freight"))

    freight = freight_rate * acc_qty

    # PVC percentages
    pct_a2 = pvc_percent(pvc_base_date, call_date, ieema_df)
    pct_b2 = pvc_percent(pvc_base_date, scheduled_date, ieema_df)
    pct_c1 = pvc_percent(lower_basic_date, call_date, ieema_df)
    pct_d1 = pvc_percent(lower_basic_date, scheduled_date, ieema_df)
    pvc_ps_a2 = basic_rate * pct_a2 / 100 if pct_a2 else 0
    pvc_ps_b2 = basic_rate * pct_b2 / 100 if pct_b2 else 0
    pvc_ps_c1 = lower_rate * pct_c1 / 100 if pct_c1 else 0
    pvc_ps_d1 = lower_rate * pct_d1 / 100 if pct_d1 else 0
    base_amt = basic_rate * acc_qty
    lower_amt = lower_rate * acc_qty
    lower_freight_total = lower_freight * acc_qty

    pvc_actual = (
        (base_amt + base_amt * pct_a2 / 100 + freight) * GST_FACTOR
        if pct_a2 is not None
        else 0
    )
    pvc_contractual = (
        (base_amt + base_amt * pct_b2 / 100 + freight) * GST_FACTOR
        if pct_b2 is not None
        else 0
    )
    lower_actual = (
        (lower_amt + lower_amt * pct_c1 / 100 + lower_freight_total) * GST_FACTOR
        if pct_c1 is not None
        else 0
    )
    lower_contractual = (
        (lower_amt + lower_amt * pct_d1 / 100 + lower_freight_total) * GST_FACTOR
        if pct_d1 is not None
        else 0
    )

     # ----- LD base logic -----
    # 1) extended_dp present -> LD from extended_dp to supply (no PVC on this)
    # 2) elif refixed_dp present -> LD from refixed_dp to supply
    # 3) else -> LD from orig_dp to supply
    delay_days = 0
    ld_weeks_new = 0
    ld_rate_pct_new = 0
    ld_applicable = True
    ld_base = None
    if pd.notna(extended_dp):
        ld_applicable = False
        ld_base = extended_dp
    elif pd.notna(refixed_dp):
        ld_base = refixed_dp
    else:
        ld_base = orig_dp

    if ld_applicable and pd.notna(sup_date) and pd.notna(ld_base):
        delay_days = max((sup_date - ld_base).days, 0)
    else:
        delay_days = 0

    if not ld_applicable or delay_days <= 0:
        ld_applicable = False
        delay_days = 0
        ld_weeks_new = 0
        ld_rate_pct_new = 0
    else:
        ld_weeks_new = math.ceil(delay_days / 7)
        ld_rate_pct_new = min(ld_weeks_new * 0.5, 10)

    ld_amt_actual = max(pvc_actual, 0) * ld_rate_pct_new / 100
    ld_amt_contractual = max(pvc_contractual, 0) * ld_rate_pct_new / 100
    

    pvc_actual_less_ld_new = pvc_actual - ld_amt_actual if pvc_actual else None
    pvc_contractual_less_ld_new = (
        pvc_contractual - ld_amt_contractual if pvc_contractual else None
    )

    lower_actual_less_ld = (
        lower_actual - (max(lower_actual, 0) * ld_rate_pct_new / 100)
        if lower_actual
        else None
    )
    lower_contractual_less_ld = (
        lower_contractual - (max(lower_contractual, 0) * ld_rate_pct_new / 100)
        if lower_contractual
        else None
    )

    rateapplied = str(d.get("rateapplied", "")).strip().lower()

    if rateapplied == "lower rate applicable":
        candidates_new = {
            "A2": pvc_actual,
            "B2": pvc_contractual,
            "C1": lower_actual,
            "D1": lower_contractual,
        }

    elif rateapplied == "lower rate and ld comparison":
        candidates_new = {
            "A2": pvc_actual_less_ld_new,
            "B2": pvc_contractual_less_ld_new,
            "C1": lower_actual,
            "D1": lower_contractual,
        }

    elif rateapplied == "lower rate with ld in further extension":
        candidates_new = {
            "A2": pvc_actual_less_ld_new,
            "B2": pvc_contractual_less_ld_new,
            "C1": lower_actual_less_ld,
            "D1": lower_contractual_less_ld,
        }

    else:
        candidates_new = {
            "A2": pvc_actual_less_ld_new,
            "B2": pvc_contractual_less_ld_new,
            "C1": lower_actual,
            "D1": lower_contractual,
    }    
    candidates_new = {k: v for k, v in candidates_new.items() if v and v > 0}

    selected_scenario_new = (
        min(candidates_new, key=candidates_new.get) if candidates_new else None
    )
    fair_price_new = candidates_new.get(selected_scenario_new, 0)

    result_row = {
        "acc_qty": acc_qty,
        "basic_rate": basic_rate,
        "pvc_base_date": pvc_base_date,
        "lower_rate": lower_rate,
        "lower_freight": lower_freight,
        "lower_freight_total": safe_round(lower_freight_total),
        "lower_basic_date": lower_basic_date,
        "freight_rate_per_unit": freight_rate,
        "freight": safe_round(freight),
        "orig_dp": orig_dp,
        "refixeddp": refixed_dp,
        "extendeddp": extended_dp,
        "scheduled_date": scheduled_date,
        "call_date": call_date,
        "sup_date": sup_date,
        "pvc_actual": safe_round(pvc_actual),
        "pvc_contractual": safe_round(pvc_contractual),
        "lower_actual": safe_round(lower_actual),
        "lower_contractual": safe_round(lower_contractual),
        "delay_days": delay_days,
        "ld_weeks_new": ld_weeks_new,
        "ld_rate_pct_new": safe_round(ld_rate_pct_new),
        "ld_applicable": ld_applicable,
        "ld_amt_actual": safe_round(ld_amt_actual),
        "ld_amt_contractual": safe_round(ld_amt_contractual),
        "pvc_actual_less_ld_new": safe_round(pvc_actual_less_ld_new),
        "pvc_contractual_less_ld_new": safe_round(pvc_contractual_less_ld_new),
        "fair_price_new": safe_round(fair_price_new),
        "selected_scenario_new": selected_scenario_new,
        "pvc_per_set_a2": safe_round(pvc_ps_a2),
        "pvc_per_set_b2": safe_round(pvc_ps_b2),
        "pvc_per_set_c1": safe_round(pvc_ps_c1),
        "pvc_per_set_d1": safe_round(pvc_ps_d1),
    }

    scenario_amounts = {
        "A2": safe_round(pvc_actual_less_ld_new),
        "B2": safe_round(pvc_contractual_less_ld_new),
        "C1": safe_round(lower_actual),
        "D1": safe_round(lower_contractual),
    }
    result_row["scenario_amounts"] = scenario_amounts

    scenario_details = []
    for sc, bd, cd in [
        ("A2", pvc_base_date, call_date),
        ("B2", pvc_base_date, scheduled_date),
        ("C1", lower_basic_date, call_date),
        ("D1", lower_basic_date, scheduled_date),
    ]:
        if not ld_applicable and sc != "A2":
            continue
        det = pvc_percent_detailed(bd, cd, ieema_df, sc)
        if det:
            scenario_details.append(det)
    result_row["scenario_details"] = scenario_details

    return result_row


# =====================================================
# BATCH MAIN (Excel input → Excel output)
# =====================================================
def main():
    logger.info("🚀 PVC ANALYSIS STARTED")

    df = pd.read_excel(INPUT_FILE)
    ieema = load_ieema()

    for c in DATE_COLS:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    final_rows = []
    scenario_rows = []

    for r in df.itertuples(index=False):
        if pd.isna(r.pvc_base_date):
            logger.warning("Missing pvc_base_date")
        if pd.isna(r.call_date):
            logger.warning("Missing call_date")
        if pd.isna(r.scheduled_date):
            logger.warning("Missing scheduled_date")
        if pd.isna(r.sup_date):
            logger.warning("Missing sup_date")

        acc_qty = safe_float(r.acc_qty)
        basic_rate = safe_float(r.basic_rate)

        freight = safe_float(r.freight_rate_per_unit) * acc_qty
        lower_rate = safe_float(r.lower_rate)
        lower_freight = safe_float(r.lower_freight)

        pct_a2 = pvc_percent(r.pvc_base_date, r.call_date, ieema)
        pct_b2 = pvc_percent(r.pvc_base_date, r.scheduled_date, ieema)
        pct_c1 = pvc_percent(r.lower_basic_date, r.call_date, ieema)
        pct_d1 = pvc_percent(r.lower_basic_date, r.scheduled_date, ieema)

        base_amt = basic_rate * acc_qty
        lower_amt = lower_rate * acc_qty
        lower_freight_total = lower_freight * acc_qty

        pvc_actual = (
            (base_amt + base_amt * pct_a2 / 100 + freight) * GST_FACTOR
            if pct_a2 is not None
            else 0
        )
        pvc_contractual = (
            (base_amt + base_amt * pct_b2 / 100 + freight) * GST_FACTOR
            if pct_b2 is not None
            else 0
        )
        lower_actual = (
            (lower_amt + lower_amt * pct_c1 / 100 + lower_freight_total) * GST_FACTOR
            if pct_c1 is not None
            else 0
        )
        lower_contractual = (
            (lower_amt + lower_amt * pct_d1 / 100 + lower_freight_total) * GST_FACTOR
            if pct_d1 is not None
            else 0
        )

        pvc_ps_a2 = basic_rate * pct_a2 / 100 if pct_a2 else 0
        pvc_ps_b2 = basic_rate * pct_b2 / 100 if pct_b2 else 0
        pvc_ps_c1 = lower_rate * pct_c1 / 100 if pct_c1 else 0
        pvc_ps_d1 = lower_rate * pct_d1 / 100 if pct_d1 else 0

            # ----- LD base logic -----
    delay_days = 0
    ld_weeks_new = 0
    ld_rate_pct_new = 0
    ld_applicable = True
    ld_base = None

    if pd.notna(extended_dp):
        ld_base = extended_dp
    elif pd.notna(refixed_dp):
        ld_base = refixed_dp
    else:
        ld_base = orig_dp

    if pd.notna(sup_date) and pd.notna(ld_base):
        delay_days = max((sup_date - ld_base).days, 0)

    if delay_days <= 0:
        ld_applicable = False
        delay_days = 0
        ld_weeks_new = 0
        ld_rate_pct_new = 0
    else:
        ld_weeks_new = math.ceil(delay_days / 7)
        ld_rate_pct_new = min(ld_weeks_new * 0.5, 10)

        ld_amt_actual = max(pvc_actual, 0) * ld_rate_pct_new / 100
        ld_amt_contractual = max(pvc_contractual, 0) * ld_rate_pct_new / 100

        pvc_actual_less_ld_new = pvc_actual - ld_amt_actual if pvc_actual else None
        pvc_contractual_less_ld_new = (
            pvc_contractual - ld_amt_contractual if pvc_contractual else None
        )

        lower_actual_less_ld = (
            lower_actual - (max(lower_actual, 0) * ld_rate_pct_new / 100)
            if lower_actual
            else None
        )
        lower_contractual_less_ld = (
            lower_contractual - (max(lower_contractual, 0) * ld_rate_pct_new / 100)
            if lower_contractual
            else None
        )

        rateapplied = str(d.get("rateapplied", "")).strip().lower()

        if rateapplied == "Lower rate applicable":
            candidates_new = {
                "A2": pvc_actual,
                "B2": pvc_contractual,
                "C1": lower_actual,
                "D1": lower_contractual,
            }
        elif rateapplied == "Lower rate and LD comparison":
            candidates_new = {
                "A2": pvc_actual_less_ld_new,
                "B2": pvc_contractual_less_ld_new,
                "C1": lower_actual,
                "D1": lower_contractual,
            }
        elif rateapplied == "Lower rate with LD in further extension":
            candidates_new = {
                "A2": pvc_actual_less_ld_new,
                "B2": pvc_contractual_less_ld_new,
                "C1": lower_actual_less_ld,
                "D1": lower_contractual_less_ld,
            }
        else:
            candidates_new = {
                "A2": pvc_actual_less_ld_new,
                "B2": pvc_contractual_less_ld_new,
                "C1": lower_actual,
                "D1": lower_contractual,
            }

        candidates_new = {k: v for k, v in candidates_new.items() if v and v > 0}
        selected_scenario_new = (
            min(candidates_new, key=candidates_new.get) if candidates_new else None
        )
        fair_price_new = candidates_new.get(selected_scenario_new, 0)

        final_rows.append(
            {
                "acc_qty": acc_qty,
                "basic_rate": basic_rate,
                "pvc_base_date": r.pvc_base_date,
                "lower_rate": lower_rate,
                "lower_freight": lower_freight,
                "lower_freight_total": lower_freight_total,
                "lower_basic_date": r.lower_basic_date,
                "freight_rate_per_unit": safe_float(r.freight_rate_per_unit),
                "freight": safe_round(freight),
                "orig_dp": r.orig_dp,
                "scheduled_date": r.scheduled_date,
                "sup_date": r.sup_date,
                "call_date": r.call_date,
                "rateapplied": r.rateapplied,
                "pvc_actual": safe_round(pvc_actual),
                "pvc_contractual": safe_round(pvc_contractual),
                "lower_actual": safe_round(lower_actual),
                "lower_contractual": safe_round(lower_contractual),
                "delay_days": delay_days,
                "ld_weeks_new": ld_weeks_new,
                "ld_rate_pct_new": safe_round(ld_rate_pct_new),
                "ld_amt_actual": safe_round(ld_amt_actual),
                "ld_amt_contractual": safe_round(ld_amt_contractual),
                "pvc_actual_less_ld_new": safe_round(pvc_actual_less_ld_new),
                "pvc_contractual_less_ld_new": safe_round(pvc_contractual_less_ld_new),
                "lower_actual_less_ld": safe_round(lower_actual_less_ld),
                "lower_contractual_less_ld": safe_round(lower_contractual_less_ld),
                "fair_price_new": safe_round(fair_price_new),
                "selected_scenario_new": selected_scenario_new,
                "pvc_per_set_a2": safe_round(pvc_ps_a2),
                "pvc_per_set_b2": safe_round(pvc_ps_b2),
                "pvc_per_set_c1": safe_round(pvc_ps_c1),
                "pvc_per_set_d1": safe_round(pvc_ps_d1),
            }
        )

        for sc, bd, cd in [
            ("A2", r.pvc_base_date, r.call_date),
            ("B2", r.pvc_base_date, r.scheduled_date),
            ("C1", r.lower_basic_date, r.call_date),
            ("D1", r.lower_basic_date, r.scheduled_date),
        ]:
            det = pvc_percent_detailed(bd, cd, ieema, sc)
            if det:
                det.update(
                    {
                        "selected_scenario_new": selected_scenario_new,
                        "rateapplied": r.rateapplied,
                    }
                )
                scenario_rows.append(det)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame(final_rows).to_excel(
            writer, index=False, sheet_name="PVC_FINAL"
        )
        pd.DataFrame(scenario_rows).to_excel(
            writer, index=False, sheet_name="PVC_SCENARIO_INDEX_DETAILS"
        )

    logger.info("✅ PVC FINAL REPORT CREATED SUCCESSFULLY")


if __name__ == "__main__":
    main()