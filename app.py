"""
Labor Law Claim Generator Bot - Levin Telraz Law Firm
Generates Israeli labor law claims (כתבי תביעה) based on client intake data.
"""

import json
import math
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os
import io

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "lt-labor-law-bot-secret-key-2026")
app.config["PERMANENT_SESSION_LIFETIME"] = 86400  # 24 hours in seconds

APP_PASSWORD = os.environ.get("APP_PASSWORD", "LT2026")

# ── Israeli Labor Law Constants ──────────────────────────────────────────────

MINIMUM_WAGE_2024 = 5880.02  # NIS monthly
MINIMUM_WAGE_HOURLY_2024 = 32.30
PENSION_EMPLOYER_RATE = 0.065  # 6.5%
PENSION_EMPLOYEE_RATE = 0.06   # 6%
SEVERANCE_RATE = 0.0833        # 8.33%
RECUPERATION_DAY_VALUE = 418   # NIS per day (2024)

# Recuperation days entitlement by year
RECUPERATION_DAYS = {
    1: 5,   # Year 1
    2: 6,   # Year 2
    3: 6,   # Year 3
    4: 7,   # Years 4-10
    5: 7,
    6: 7,
    7: 7,
    8: 7,
    9: 7,
    10: 7,
    11: 8,  # Years 11-15
    15: 8,
    16: 9,  # Years 16-19
    20: 10, # Year 20+
}

# Vacation days entitlement by year (6-day work week)
VACATION_DAYS_6DAY = {
    1: 14, 2: 14, 3: 14, 4: 14,
    5: 16,
    6: 18, 7: 21,
    8: 22, 9: 23, 10: 24, 11: 25, 12: 26, 13: 27, 14: 28,
}

# Vacation days entitlement by year (5-day work week)
VACATION_DAYS_5DAY = {
    1: 12, 2: 12, 3: 12, 4: 12,
    5: 13,
    6: 15, 7: 18,
    8: 19, 9: 20, 10: 21, 11: 22, 12: 23, 13: 24, 14: 24,
}

# Jewish holidays per year (typically 9 days)
HOLIDAY_DAYS_PER_YEAR = 9

# Overtime rates
OVERTIME_125_RATE = 0.25  # First 2 hours
OVERTIME_150_RATE = 0.50  # Beyond 2 hours


def calculate_employment_duration(start_date, end_date):
    """Calculate employment duration in years and months."""
    delta = relativedelta(end_date, start_date)
    total_months = delta.years * 12 + delta.months
    years = total_months / 12
    return {
        "years": delta.years,
        "months": delta.months,
        "total_months": total_months,
        "decimal_years": round(years, 2),
    }


def calculate_determining_salary(base_salary, commissions=0, extras=0):
    """Calculate the determining salary (שכר קובע) for labor law purposes."""
    return base_salary + commissions + extras


def calculate_severance(determining_salary, years_decimal):
    """Calculate severance pay (פיצויי פיטורים)."""
    return round(determining_salary * years_decimal, 2)


def calculate_vacation_entitlement(years_decimal, work_days_per_week, daily_rate):
    """Calculate vacation entitlement and monetary value."""
    table = VACATION_DAYS_6DAY if work_days_per_week == 6 else VACATION_DAYS_5DAY
    total_days = 0
    full_years = int(years_decimal)
    fraction = years_decimal - full_years

    for y in range(1, full_years + 1):
        key = min(y, max(table.keys()))
        total_days += table.get(y, table[max(k for k in table if k <= y)])

    if fraction > 0 and full_years + 1 <= max(table.keys()):
        next_year = full_years + 1
        next_entitlement = table.get(next_year, table[max(k for k in table if k <= next_year)])
        total_days += round(next_entitlement * fraction, 2)

    return {
        "total_days": round(total_days, 2),
        "value": round(total_days * daily_rate, 2),
    }


def calculate_recuperation(years_decimal, daily_value=RECUPERATION_DAY_VALUE):
    """Calculate recuperation pay (דמי הבראה)."""
    total_days = 0
    full_years = int(years_decimal)
    fraction = years_decimal - full_years

    for y in range(1, full_years + 1):
        if y <= 3:
            days = RECUPERATION_DAYS.get(y, 6)
        elif y <= 10:
            days = 7
        elif y <= 15:
            days = 8
        elif y <= 19:
            days = 9
        else:
            days = 10
        total_days += days

    if fraction > 0:
        next_y = full_years + 1
        if next_y <= 3:
            days = RECUPERATION_DAYS.get(next_y, 6)
        elif next_y <= 10:
            days = 7
        else:
            days = 8
        total_days += round(days * fraction, 2)

    return {
        "total_days": round(total_days, 2),
        "value": round(total_days * daily_value, 2),
    }


def calculate_holiday_pay(years_decimal, daily_rate, days_paid=0, rate_paid=0):
    """Calculate holiday pay entitlement (דמי חגים)."""
    total_days = round(HOLIDAY_DAYS_PER_YEAR * years_decimal)
    entitled_value = total_days * daily_rate
    paid_value = days_paid * rate_paid
    difference = entitled_value - paid_value
    return {
        "total_days": total_days,
        "entitled_value": round(entitled_value, 2),
        "paid_value": round(paid_value, 2),
        "difference": round(max(0, difference), 2),
    }


def calculate_pension_gaps(monthly_salary, months, employer_rate=PENSION_EMPLOYER_RATE,
                           amount_deposited=0):
    """Calculate pension deposit gaps (הפרשי הפרשות לפנסיה)."""
    total_owed = round(monthly_salary * months * employer_rate, 2)
    gap = round(total_owed - amount_deposited, 2)
    return {
        "total_owed": total_owed,
        "deposited": amount_deposited,
        "gap": max(0, gap),
    }


def calculate_overtime(weekly_overtime_125, weekly_overtime_150, hourly_rate, months):
    """Calculate overtime pay owed (שעות נוספות) - basic weekly input mode."""
    rate_125 = hourly_rate * 1.25
    rate_150 = hourly_rate * 1.50
    surcharge_125 = hourly_rate * OVERTIME_125_RATE
    surcharge_150 = hourly_rate * OVERTIME_150_RATE

    monthly_125 = weekly_overtime_125 * 4 * surcharge_125
    monthly_150 = weekly_overtime_150 * 4 * surcharge_150

    total = round((monthly_125 + monthly_150) * months, 2)
    return {
        "mode": "basic",
        "monthly_125": round(monthly_125, 2),
        "monthly_150": round(monthly_150, 2),
        "total": total,
        "rate_125": round(rate_125, 2),
        "rate_150": round(rate_150, 2),
        "surcharge_125": round(surcharge_125, 2),
        "surcharge_150": round(surcharge_150, 2),
    }


def calculate_overtime_global(hourly_wage, standard_daily_hours, actual_daily_hours,
                              global_ot_hours, work_days_per_week, months):
    """Calculate overtime with global OT comparison (125%/150% daily breakdown)."""
    daily_ot = max(0, actual_daily_hours - standard_daily_hours)
    daily_ot_125 = min(daily_ot, 2)
    daily_ot_150 = max(0, daily_ot - 2)

    rate_125 = hourly_wage * 1.25
    rate_150 = hourly_wage * 1.50

    work_days_per_month = round(work_days_per_week * 4.33, 1)

    monthly_ot_125_hours = round(daily_ot_125 * work_days_per_month, 2)
    monthly_ot_150_hours = round(daily_ot_150 * work_days_per_month, 2)
    monthly_should_pay = round(monthly_ot_125_hours * rate_125 + monthly_ot_150_hours * rate_150, 2)

    global_125_hours = min(global_ot_hours, daily_ot_125 * work_days_per_month) if daily_ot > 0 else min(global_ot_hours, 2 * work_days_per_month)
    global_150_hours = max(0, global_ot_hours - global_125_hours)
    monthly_paid = round(global_125_hours * rate_125 + global_150_hours * rate_150, 2)

    monthly_difference = round(max(0, monthly_should_pay - monthly_paid), 2)
    total = round(monthly_difference * months, 2)

    return {
        "mode": "global",
        "hourly_wage": hourly_wage,
        "standard_daily_hours": standard_daily_hours,
        "actual_daily_hours": actual_daily_hours,
        "daily_ot": round(daily_ot, 2),
        "daily_ot_125": round(daily_ot_125, 2),
        "daily_ot_150": round(daily_ot_150, 2),
        "rate_125": round(rate_125, 2),
        "rate_150": round(rate_150, 2),
        "work_days_per_month": work_days_per_month,
        "monthly_ot_125_hours": monthly_ot_125_hours,
        "monthly_ot_150_hours": monthly_ot_150_hours,
        "monthly_should_pay": monthly_should_pay,
        "global_ot_hours": global_ot_hours,
        "global_125_hours": round(global_125_hours, 2),
        "global_150_hours": round(global_150_hours, 2),
        "monthly_paid": monthly_paid,
        "monthly_difference": monthly_difference,
        "total": total,
        "months": months,
    }


def safe_float(val, default=0):
    """Safely convert a value to float, returning default on failure."""
    if val is None or val == '':
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def safe_int(val, default=0):
    """Safely convert a value to int, returning default on failure."""
    if val is None or val == '':
        return default
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return default


def calculate_all_claims(data):
    """Master calculation function for all claim components."""
    start_str = (data.get("start_date") or "").strip()
    end_str = (data.get("end_date") or "").strip()

    if not start_str or not end_str:
        raise ValueError("יש להזין תאריך תחילת עבודה ותאריך סיום עבודה")

    try:
        start = datetime.strptime(start_str, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(f"תאריך תחילת עבודה אינו תקין: {start_str}")

    try:
        end = datetime.strptime(end_str, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(f"תאריך סיום עבודה אינו תקין: {end_str}")

    if end <= start:
        raise ValueError("תאריך סיום העבודה חייב להיות מאוחר מתאריך ההתחלה")

    duration = calculate_employment_duration(start, end)

    base_salary = safe_float(data.get("base_salary"), 0)
    commissions = safe_float(data.get("commissions"), 0)
    determining_salary = calculate_determining_salary(base_salary, commissions)

    work_days = safe_int(data.get("work_days_per_week"), 6)
    hours_per_day = safe_float(data.get("hours_per_day"), 8.5 if work_days == 6 else 9)
    monthly_hours = work_days * hours_per_day * 4.33
    hourly_rate = round(determining_salary / monthly_hours, 2) if monthly_hours > 0 else 0
    daily_rate = round(determining_salary / (work_days * 4.33), 2) if work_days > 0 else 0

    results = {
        "duration": duration,
        "determining_salary": determining_salary,
        "hourly_rate": hourly_rate,
        "daily_rate": daily_rate,
        "claims": {},
        "total": 0,
    }

    # Severance (פיצויי פיטורים)
    if data.get("claim_severance"):
        severance = calculate_severance(determining_salary, duration["decimal_years"])
        deposited = safe_float(data.get("severance_deposited"), 0)
        results["claims"]["severance"] = {
            "name": "פיצויי פיטורים",
            "full_amount": severance,
            "deposited": deposited,
            "amount": max(0, round(severance - deposited, 2)),
        }

    # Unpaid salary (שכר שלא שולם)
    if data.get("claim_unpaid_salary"):
        unpaid = safe_float(data.get("unpaid_salary_amount"), 0)
        results["claims"]["unpaid_salary"] = {
            "name": "שכר עבודה שלא שולם",
            "amount": unpaid,
        }

    # Overtime (שעות נוספות)
    if data.get("claim_overtime"):
        ot_hourly_wage = safe_float(data.get("hourly_wage"), 0)
        ot_actual = safe_float(data.get("actual_daily_hours"), 0)
        # Use global OT mode if hourly_wage and actual_daily_hours are filled in
        if ot_hourly_wage > 0 and ot_actual > 0:
            ot_standard = safe_float(data.get("standard_daily_hours"), 8)
            ot_global = safe_float(data.get("global_ot_hours"), 0)
            ot = calculate_overtime_global(
                ot_hourly_wage, ot_standard, ot_actual,
                ot_global, work_days, duration["total_months"],
            )
        else:
            ot = calculate_overtime(
                safe_float(data.get("weekly_overtime_125"), 0),
                safe_float(data.get("weekly_overtime_150"), 0),
                hourly_rate,
                duration["total_months"],
            )
        results["claims"]["overtime"] = {
            "name": "הפרשי שכר – שעות נוספות",
            "amount": ot["total"],
            "details": ot,
        }

    # Pension gaps (הפרשי הפרשות לפנסיה)
    if data.get("claim_pension"):
        pension = calculate_pension_gaps(
            determining_salary,
            duration["total_months"],
            amount_deposited=safe_float(data.get("pension_deposited"), 0),
        )
        results["claims"]["pension"] = {
            "name": "הפרשי הפרשות לפנסיה",
            "amount": pension["gap"],
            "details": pension,
        }

    # Vacation (חופשה)
    if data.get("claim_vacation"):
        vac = calculate_vacation_entitlement(
            duration["decimal_years"], work_days, daily_rate
        )
        paid_days = safe_float(data.get("vacation_days_paid"), 0)
        paid_rate = safe_float(data.get("vacation_rate_paid"), 0)
        paid_value = paid_days * paid_rate
        gap = round(vac["value"] - paid_value, 2)
        results["claims"]["vacation"] = {
            "name": "הפרשי דמי חופשה ופדיון חופשה",
            "entitled_days": vac["total_days"],
            "paid_days": paid_days,
            "amount": max(0, gap),
        }

    # Holiday pay (דמי חגים)
    if data.get("claim_holidays"):
        hol = calculate_holiday_pay(
            duration["decimal_years"],
            daily_rate,
            days_paid=safe_float(data.get("holiday_days_paid"), 0),
            rate_paid=safe_float(data.get("holiday_rate_paid"), 0),
        )
        results["claims"]["holidays"] = {
            "name": "דמי חגים והפרשי דמי חג",
            "amount": hol["difference"],
            "details": hol,
        }

    # Recuperation (הבראה)
    if data.get("claim_recuperation"):
        rec = calculate_recuperation(duration["decimal_years"])
        paid_days = safe_float(data.get("recuperation_days_paid"), 0)
        paid_value = paid_days * RECUPERATION_DAY_VALUE
        gap = round(rec["value"] - paid_value, 2)
        results["claims"]["recuperation"] = {
            "name": "דמי הבראה",
            "entitled_days": rec["total_days"],
            "paid_days": paid_days,
            "amount": max(0, gap),
        }

    # Salary delay damages (פיצויי הלנת שכר)
    if data.get("claim_salary_delay"):
        delay_amount = safe_float(data.get("salary_delay_amount"), 0)
        results["claims"]["salary_delay"] = {
            "name": "פיצויי הלנת שכר",
            "amount": delay_amount,
        }

    # Emotional distress (עוגמת נפש)
    if data.get("claim_emotional"):
        emotional = safe_float(data.get("emotional_amount"), 25000)
        results["claims"]["emotional"] = {
            "name": "פיצוי בגין עוגמת נפש",
            "amount": emotional,
        }

    # Unlawful deductions (ניכויים שלא כדין)
    if data.get("claim_deductions"):
        deductions = safe_float(data.get("deduction_amount"), 0)
        results["claims"]["deductions"] = {
            "name": "ניכויים שלא כדין",
            "amount": deductions,
        }

    # Total
    results["total"] = round(sum(c["amount"] for c in results["claims"].values()), 2)
    return results


def generate_claim_text(data, calculations):
    """Generate the full Hebrew legal claim text based on the firm's template."""

    plaintiff_name = data.get("plaintiff_name", "")
    plaintiff_id = data.get("plaintiff_id", "")
    defendant_name = data.get("defendant_name", "")
    defendant_id = data.get("defendant_id", "")
    defendant_type = data.get("defendant_type", "company")
    defendant_owner = data.get("defendant_owner", "")
    defendant_business = data.get("defendant_business", "")
    job_title = data.get("job_title", "")
    start_date = data.get("start_date", "")
    end_date = data.get("end_date", "")
    termination_type = data.get("termination_type", "fired")
    work_schedule = data.get("work_schedule", "")
    narrative = data.get("narrative", "")

    dur = calculations["duration"]
    det_salary = calculations["determining_salary"]
    hourly = calculations["hourly_rate"]
    daily = calculations["daily_rate"]
    total = calculations["total"]

    # Format dates for Hebrew
    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    start_fmt = start_dt.strftime("%d.%m.%Y")
    end_fmt = end_dt.strftime("%d.%m.%Y")

    # Gender handling
    gender = data.get("gender", "male")
    plaintiff_title = "מר" if gender == "male" else "הגב'"
    pronoun = "התובע" if gender == "male" else "התובעת"
    pronoun_he = "הוא" if gender == "male" else "היא"
    worked = "עבד" if gender == "male" else "עבדה"
    was_forced = "נאלץ" if gender == "male" else "נאלצה"

    # Defendant label
    if defendant_type == "company":
        defendant_label = f"חברת {defendant_name}"
        defendant_desc = f"הינה חברה בבעלותו ותחת ניהולו של {defendant_owner} העוסקת ב{defendant_business}"
    else:
        defendant_label = defendant_name
        defendant_desc = f"העוסק/ת ב{defendant_business}"

    # Termination language
    if termination_type == "fired":
        termination_text = f"עד שפוטר/ה ביום {end_fmt}"
    elif termination_type == "resigned_justified":
        termination_text = f"עד שנאלץ/ה לסיים את העסקתו/ה בדין מפוטר/ת ביום {end_fmt}"
    else:
        termination_text = f"עד שהתפטר/ה ביום {end_fmt}"

    sections = []

    # ── Header ──
    sections.append("כ ת ב    ת ב י ע ה")
    sections.append("")

    # ── General ──
    sections.append("כללי")
    sections.append(f"{pronoun} מיוצג/ת ע\"י ב\"כ, אשר מענה להמצאת כתבי בית דין הוא, כמצוין בכותרת.")
    sections.append(f"{pronoun} מגיש/ה תביעה זו כנגד הנתבע/ת בגין הפרת זכויותיו/ה כעובד/ת וכאדם, הכול כפי שיפורט להלן.")
    sections.append("הטענות שלהלן הינן חלופיות, מצטברות או משלימות - הכול לפי העניין, הקשר הדברים והדבקם.")
    sections.append("")

    # ── Parties ──
    sections.append("הצדדים")
    sections.append(
        f"{pronoun}, {plaintiff_title} {plaintiff_name}, ת.ז. {plaintiff_id}, "
        f"{worked} בנתבע/ת החל מיום {start_fmt} {termination_text}, "
        f"סה\"כ {worked} {pronoun} בנתבע/ת {dur['total_months']} חודשים "
        f"שהם {dur['decimal_years']} שנים (להלן: \"{pronoun}\")."
    )
    sections.append(f"תלושי שכר הנמצאים בידי {pronoun} מצ\"ב ומסומנים כנספח 1.")
    sections.append(
        f"הנתבע/ת, {defendant_label}, ח.פ./ע.מ. {defendant_id}, "
        f"{defendant_desc} "
        f"ומי שהיה/תה מעסיק/תו/ה של {pronoun} בתקופה הרלוונטית לכתב התביעה (להלן: \"הנתבע/ת\")."
    )
    sections.append("")

    # ── Background ──
    sections.append("רקע עובדתי")
    sections.append(
        f"{pronoun} החל/ה את עבודתו/ה בנתבע/ת כ{job_title} החל מיום {start_fmt}."
    )
    if work_schedule:
        sections.append(f"עבודתו/ה של {pronoun} התנהלה {work_schedule}.")

    sections.append(
        f"לכל אורך תקופת העסקתו/ה, {pronoun} היה/תה עובד/ת מצוין/ת ומקצועי/ת "
        f"אשר ביצע/ה את עבודתו/ה נאמנה."
    )

    if narrative:
        sections.append("")
        sections.append(narrative)

    sections.append("")

    # ── Employment Scope and Determining Salary ──
    sections.append("היקף משרה ושכר קובע")
    base = safe_float(data.get("base_salary"), 0)
    comm = safe_float(data.get("commissions"), 0)

    salary_desc = f"שכרו/ה של {pronoun} עמד על סך של {base:,.0f} ₪ ברוטו"
    if comm > 0:
        salary_desc += f" בגין שכר בסיס ובנוסף {comm:,.0f} ₪ בגין עמלות/תוספות חודשיות"
    salary_desc += "."

    sections.append(salary_desc)
    sections.append(
        f"סה\"כ שכרו/ה החודשי הקובע של {pronoun} עמד על {det_salary:,.0f} ₪ ברוטו, "
        f"כך ששכרו/ה השעתי הקובע עמד על סך של {hourly:,.1f} ₪ "
        f"ושכרו/ה היומי הקובע עמד על סך של {daily:,.0f} ₪."
    )
    sections.append("")

    # ── Claim Components ──
    sections.append("רכיבי התביעה")
    sections.append("")

    appendix_num = 2

    claims = calculations["claims"]

    # Unpaid salary
    if "unpaid_salary" in claims:
        c = claims["unpaid_salary"]
        sections.append(f"שכר עבודה שלא שולם")
        sections.append(
            f"כאמור, {pronoun} יטען/תטען כי הנתבע/ת לא שילם/ה לו/ה את שכרו/ה כנדרש על פי דין."
        )
        sections.append(
            f"לפיכך, {pronoun} יבקש/תבקש מבית הדין הנכבד לחייב את הנתבע/ת לשלם ל{pronoun} "
            f"שכר עבודה שלא שולם בסך של {c['amount']:,.0f} ₪ "
            f"בצירוף פיצוי הלנת שכר או הפרשי הצמדה וריבית לפי העניין עד מועד התשלום בפועל."
        )
        sections.append("")

    # Overtime
    if "overtime" in claims:
        c = claims["overtime"]
        d = c["details"]
        sections.append("הפרשי שכר – שעות נוספות")

        if d.get("mode") == "global":
            # Global OT mode - detailed calculation with paid-vs-owed comparison
            sections.append(
                f"כאמור, {pronoun} יטען/תטען כי הנתבע/ת לא שילם/ה לו/ה כנדרש בגין השעות הנוספות הרבות אותן {worked}."
            )
            sections.append(
                f"שכרו/ה השעתי הבסיסי של {pronoun} הינו {d['hourly_wage']:.2f} ₪. "
                f"יום עבודה סטנדרטי: {d['standard_daily_hours']:.1f} שעות. "
                f"שעות עבודה בפועל ביום (ממוצע): {d['actual_daily_hours']:.1f} שעות."
            )
            sections.append(
                f"בהתאם לחוק שעות עבודה ומנוחה, תשי\"א-1951, "
                f"2 השעות הנוספות הראשונות מזכות בתוספת 25% (תעריף {d['rate_125']:.2f} ₪) "
                f"ומעבר לכך בתוספת 50% (תעריף {d['rate_150']:.2f} ₪)."
            )
            sections.append(
                f"תחשיב שעות נוספות שהיה צריך לשלם בכל חודש:"
            )
            sections.append(
                f"שעות נוספות ביום: {d['daily_ot']:.1f} שעות "
                f"({d['daily_ot_125']:.1f} שעות × 125% + {d['daily_ot_150']:.1f} שעות × 150%)"
            )
            sections.append(
                f"ימי עבודה בחודש: {d['work_days_per_month']:.1f} ימים"
            )
            sections.append(
                f"סכום שהיה צריך לשלם בחודש: "
                f"{d['monthly_ot_125_hours']:.1f} שעות × {d['rate_125']:.2f} ₪ + "
                f"{d['monthly_ot_150_hours']:.1f} שעות × {d['rate_150']:.2f} ₪ = "
                f"{d['monthly_should_pay']:,.0f} ₪"
            )
            if d['global_ot_hours'] > 0:
                sections.append(
                    f"שעות נוספות גלובליות ששולמו בפועל: {d['global_ot_hours']:.1f} שעות בחודש "
                    f"(שוויין: {d['monthly_paid']:,.0f} ₪)"
                )
                sections.append(
                    f"הפרש חודשי: {d['monthly_should_pay']:,.0f} ₪ - {d['monthly_paid']:,.0f} ₪ = "
                    f"{d['monthly_difference']:,.0f} ₪"
                )
            sections.append(
                f"הפרש חודשי ({d['monthly_difference']:,.0f} ₪) × {d['months']} חודשי עבודה = "
                f"{c['amount']:,.0f} ₪"
            )
        else:
            # Basic mode - simple weekly OT
            sections.append(
                f"כאמור, {pronoun} יטען/תטען כי הנתבע/ת כלל לא שילם/ה לו/ה בגין השעות הנוספות הרבות אותן {worked}."
            )
            sections.append(
                f"שכרו/ה השעתי של {pronoun} הינו {hourly:.2f} ₪ ומשכך "
                f"תעריף תוספת 25% הינו {d['surcharge_125']:.1f} ₪ "
                f"ותעריף 50% הינו {d['surcharge_150']:.1f} ₪."
            )

        sections.append(
            f"לאור האמור לעיל, בהתאם לתחשיבים, {pronoun} יבקש/תבקש כי בית הדין הנכבד "
            f"יחייב את הנתבע/ת לשלם ל{pronoun} הפרשי שכר שעות נוספות בסך של {c['amount']:,.0f} ₪ "
            f"בצירוף פיצוי הלנת שכר או הפרשי הצמדה וריבית לפי העניין עד מועד התשלום בפועל."
        )
        sections.append("")

    # Pension
    if "pension" in claims:
        c = claims["pension"]
        sections.append("הפרשי הפרשות לפנסיה")
        sections.append(
            f"בהתאם להוראות צו ההרחבה לפנסיה חובה ולצו ההרחבה בדבר הגדלת ההפרשות לביטוח פנסיוני במשק, "
            f"היה על הנתבע/ת להפריש ל{pronoun} בגין רכיב תגמולי המעסיק {PENSION_EMPLOYER_RATE*100}% משכרו/ה המלא בכל חודש."
        )
        sections.append(
            f"בהתאם לתחשיבי {pronoun} על הנתבע/ת לשלם ל{pronoun} הפרשי הפרשות לפנסיה "
            f"בסך {c['amount']:,.0f} ₪."
        )
        sections.append(
            f"לאור האמור לעיל, בהתאם לתחשיבים, {pronoun} יבקש/תבקש כי בית הדין הנכבד "
            f"יחייב את הנתבע/ת לשלם ל{pronoun} הפרשי הפרשות לפנסיה בסך של {c['amount']:,.0f} ₪ "
            f"בצירוף פיצוי הלנת שכר או הפרשי הצמדה וריבית לפי העניין עד מועד התשלום בפועל."
        )
        sections.append("")

    # Severance
    if "severance" in claims:
        c = claims["severance"]
        sections.append("פיצויי פיטורים")
        if data.get("termination_type") == "resigned_justified":
            sections.append(
                f"{pronoun} יטען/תטען, כי לאור ההפרות החמורות והמתמשכות של הנתבע/ת "
                f"והפגיעה בזכויותיו/ה הקוגנטיות נאלץ/ה, בלית ברירה, להודיע על סיום העסקתו/ה."
            )
            sections.append(
                f"משכך, ובהתאם להוראות חוק פיצויי פיטורים, תשכ\"ג-1963 ולפסיקת בתי הדין לעבודה "
                f"{pronoun} הינו/ה זכאי/ת להתפטר בדין מפוטר/ת ולמלוא פיצויי פיטוריו/ה."
            )
        sections.append(
            f"{det_salary:,.0f} ₪ (שכר חודשי קובע) * {dur['decimal_years']} (תקופת העסקה) = {c['full_amount']:,.1f} ₪"
        )
        if c["deposited"] > 0:
            sections.append(f"בניכוי צבירת הפיצויים על שם {pronoun} בקופה בסך {c['deposited']:,.0f} ₪")
            sections.append(f"סה\"כ {pronoun} זכאי/ת להשלמת פיצויי פיטורים בסך {c['amount']:,.0f} ₪")
        sections.append(
            f"לאור האמור לעיל, {pronoun} יבקש/תבקש כי בית הדין הנכבד "
            f"יחייב את הנתבע/ת לשלם ל{pronoun} פיצויי פיטורים בסך של {c['amount']:,.0f} ₪ "
            f"בצירוף פיצוי הלנת שכר או הפרשי הצמדה וריבית לפי העניין עד מועד התשלום בפועל."
        )
        sections.append("")

    # Vacation
    if "vacation" in claims:
        c = claims["vacation"]
        sections.append("הפרשי שכר דמי חופשה ופדיון חופשה")
        sections.append(
            f"בהתאם להוראות חוק חופשה שנתית, תשי\"א-1951 "
            f"{pronoun} היה/תה זכאי/ת לצבירת ימי חופשה "
            f"ובהתאם לוותקו/ה סה\"כ {c['entitled_days']} ימי חופשה לכל אורך התקופה."
        )
        sections.append(
            f"לאור האמור לעיל {pronoun} יבקש/תבקש כי בית הדין הנכבד "
            f"יחייב את הנתבע/ת לשלם ל{pronoun} הפרשי שכר דמי חופשה ופדיון חופשה "
            f"בסך של {c['amount']:,.0f} ₪ "
            f"בצירוף פיצוי הלנת שכר או הפרשי הצמדה וריבית לפי העניין עד מועד התשלום בפועל."
        )
        sections.append("")

    # Holidays
    if "holidays" in claims:
        c = claims["holidays"]
        sections.append("דמי חגים והפרשי דמי חג")
        sections.append(
            f"בהתאם להוראות צו ההרחבה הסכם מסגרת 2000 ולאור העובדה כי {pronoun} הועסק/ה "
            f"כעובד/ת שעתי/ת, לאחר 3 חודשי עבודה בנתבע/ת, {pronoun} היה/תה זכאי/ת לתשלום "
            f"בגין {HOLIDAY_DAYS_PER_YEAR} ימי חג בכל שנת עבודה."
        )
        sections.append(
            f"לאור האמור לעיל {pronoun} יבקש/תבקש כי בית הדין הנכבד "
            f"יחייב את הנתבע/ת לשלם ל{pronoun} דמי חגים והפרשי דמי חג "
            f"בסך של {c['amount']:,.0f} ₪ "
            f"בצירוף פיצוי הלנת שכר או הפרשי הצמדה וריבית לפי העניין עד מועד התשלום בפועל."
        )
        sections.append("")

    # Recuperation
    if "recuperation" in claims:
        c = claims["recuperation"]
        sections.append("דמי הבראה")
        sections.append(
            f"בהתאם להוראות צו ההרחבה בדבר השתתפות המעסיק בהוצאות הבראה ונופש, "
            f"במהלך תקופת העסקתו/ה {pronoun} היה/תה זכאי/ת ל-{c['entitled_days']} ימי הבראה."
        )
        sections.append(
            f"לאור האמור לעיל {pronoun} יבקש/תבקש כי בית הדין הנכבד "
            f"יחייב את הנתבע/ת לשלם ל{pronoun} דמי הבראה "
            f"בסך של {c['amount']:,.0f} ₪ "
            f"בצירוף פיצוי הלנת שכר או הפרשי הצמדה וריבית לפי העניין עד מועד התשלום בפועל."
        )
        sections.append("")

    # Deductions
    if "deductions" in claims:
        c = claims["deductions"]
        sections.append("ניכויים שלא כדין – תגמולי עובד")
        sections.append(
            f"{pronoun} יטען/תטען כי הנתבע/ת ניכה/תה משכרו/ה סכומים שלא כדין ובחוסר תום לב."
        )
        sections.append(
            f"לאור האמור לעיל, {pronoun} יבקש/תבקש מבית הדין הנכבד לחייב את הנתבע/ת "
            f"לשלם ל{pronoun} בגין ניכויים שלא כדין סך של {c['amount']:,.0f} ₪ "
            f"בצירוף פיצוי הלנת שכר או הפרשי הצמדה וריבית לפי העניין עד מועד התשלום בפועל."
        )
        sections.append("")

    # Salary delay
    if "salary_delay" in claims:
        c = claims["salary_delay"]
        sections.append("פיצויי הלנת שכר")
        sections.append(
            f"במרבית תקופת העסקתו/ה הנתבע/ת היה/תה מאחר/ת באופן שיטתי ועקבי "
            f"בתשלום משכורתו/ה החודשית תוך הלנת שכרו/ה שלא כדין."
        )
        sections.append(
            f"לאור האמור לעיל ובהתאם להוראות חוק הגנת השכר, תשי\"ח-1958 "
            f"הרי ש{pronoun} זכאי/ת לפיצוי בגין הלנת שכרו/ה בסך של {c['amount']:,.0f} ₪."
        )
        sections.append("")

    # Emotional distress
    if "emotional" in claims:
        c = claims["emotional"]
        sections.append("פיצוי בגין עוגמת נפש")
        sections.append(
            f"לפיכך, {pronoun} יבקש/תבקש כי בית הדין הנכבד יורה לנתבע/ת לשלם ל{pronoun} "
            f"פיצוי בגין עוגמת נפש בסך של {c['amount']:,.0f} ₪ "
            f"בצירוף הפרשי הצמדה וריבית ממועד קום העילה ועד לתשלום בפועל."
        )
        sections.append("")

    # ── Document delivery ──
    if data.get("claim_documents"):
        sections.append("מסירת מסמכי גמר חשבון")
        sections.append(
            f"{pronoun} יטען/תטען כי חרף העובדה שיחסי העבודה נותקו כבר ביום {end_fmt} "
            f"הנתבע/ת לא מסר/ה ל{pronoun} טופס 161 ומסמכי שחרור והעברת בעלות על הקופה שבבעלותו/ה "
            f"ובכך הלכה למעשה מונע/ת ממנו/ה את הגישה לכספי הפנסיה המגיעים לו/ה על פי דין."
        )
        sections.append(
            f"לאור האמור לעיל {pronoun} יבקש/תבקש כי בית הדין הנכבד יחייב את הנתבע/ת "
            f"למסור לו/ה את מסמכי גמר החשבון ובהם טופס 161 ערוך על פי דין ומסמכי העברת בעלות."
        )
        sections.append("")

    # ── Summary ──
    sections.append("סיכום")
    sections.append("סיכום רכיבי התביעה:")
    sections.append("")

    for key, claim in claims.items():
        sections.append(f"• {claim['name']}: {claim['amount']:,.0f} ₪")

    sections.append("")
    sections.append(f"סה\"כ סכום התביעה: {total:,.0f} ₪ קרן (לא כולל הצמדה וריבית, שכ\"ט עו\"ד והוצאות)")
    sections.append("")

    sections.append(
        f"לאור ההפרות החמורות של זכויותיו/ה של {pronoun} המתוארות בהרחבה בכתב תביעה זה, "
        f"מתבקש בית הדין הנכבד להזמין את הנתבע/ת לדין, ולחייבו/ה במלוא סכום התביעה "
        f"בצירוף הפרשי הצמדה וריבית לפי העניין מקום העילה ועד מועד התשלום בפועל "
        f"כמו גם בסעדים ההצהרתיים המבוקשים."
    )
    sections.append(
        f"בנוסף, מתבקש בית הדין הנכבד לחייב את הנתבע/ת בתשלום הוצאות, שכ\"ט עו\"ד ומע\"מ בגינו."
    )
    sections.append(
        "בית הדין הנכבד מוסמך לדון בתביעה זו לאור מהותה, סכומה, מקום ביצוע העבודה ומענה של הנתבע/ת."
    )

    return "\n".join(sections)


def generate_docx(data, calculations, claim_text):
    """Generate a Word document matching SKILL.md specifications exactly.

    SKILL.md specs:
    - Page: US Letter (12240 × 15840 twips), margins top=709 right=1800 bottom=1276 left=1800
    - Font: David 12pt (24 half-points), RTL bidi, he-IL
    - Numbered paras: ListParagraph, numId, spacing 120/120/360 auto, ind left=-149 right=-709 hanging=425
    - Section headers: bold+underline, NO numbering, ind left=-716 right=-709 firstLine=6
    - Appendix refs: ◄ symbol, bold+underlined, NOT numbered
    - Summary tables: 2-col, bidiVisual BEFORE tblW, last row shaded D9E2F3
    - Signature: 2-col table (spacer 5649 + sig 3377), top border as sig line
    """
    from lxml import etree

    doc = Document()
    WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # ── Page Setup (US Letter, exact twip margins) ───────────────────────
    for section in doc.sections:
        sectPr = section._sectPr
        pgSz = sectPr.find(qn('w:pgSz'))
        if pgSz is None:
            pgSz = etree.SubElement(sectPr, qn('w:pgSz'))
        pgSz.set(qn('w:w'), '12240')
        pgSz.set(qn('w:h'), '15840')

        pgMar = sectPr.find(qn('w:pgMar'))
        if pgMar is None:
            pgMar = etree.SubElement(sectPr, qn('w:pgMar'))
        pgMar.set(qn('w:top'), '709')
        pgMar.set(qn('w:right'), '1800')
        pgMar.set(qn('w:bottom'), '1276')
        pgMar.set(qn('w:left'), '1800')
        pgMar.set(qn('w:header'), '720')
        pgMar.set(qn('w:footer'), '720')

    # ── Configure Default Styles ─────────────────────────────────────────
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'David'
    style_normal.font.size = Pt(12)
    style_normal.font.rtl = True
    # Set cs font on Normal style
    rPr_style = style_normal.element.find(qn('w:rPr'))
    if rPr_style is None:
        rPr_style = etree.SubElement(style_normal.element, qn('w:rPr'))
    rFonts_style = rPr_style.find(qn('w:rFonts'))
    if rFonts_style is None:
        rFonts_style = etree.SubElement(rPr_style, qn('w:rFonts'))
    rFonts_style.set(qn('w:cs'), 'David')
    rFonts_style.set(qn('w:eastAsia'), 'David')
    szCs_style = rPr_style.find(qn('w:szCs'))
    if szCs_style is None:
        szCs_style = etree.SubElement(rPr_style, qn('w:szCs'))
    szCs_style.set(qn('w:val'), '24')

    pf = style_normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style_pPr = style_normal.element.get_or_add_pPr()
    if style_pPr.find(qn('w:bidi')) is None:
        etree.SubElement(style_pPr, qn('w:bidi'))
    # Set spacing on Normal style via XML for exact twip values
    sp = style_pPr.find(qn('w:spacing'))
    if sp is None:
        sp = etree.SubElement(style_pPr, qn('w:spacing'))
    sp.set(qn('w:before'), '120')
    sp.set(qn('w:after'), '120')
    sp.set(qn('w:line'), '360')
    sp.set(qn('w:lineRule'), 'auto')

    # ── Create Numbering ─────────────────────────────────────────────────
    try:
        numbering_part = doc.part.numbering_part
    except Exception:
        dummy = doc.add_paragraph('', style='List Number')
        numbering_part = doc.part.numbering_part
        dummy._element.getparent().remove(dummy._element)
    numbering_elm = numbering_part.element

    # SKILL.md numbering: decimal, "%1.", lvlJc="left", b val="0", bCs val="0", lang bidi="he-IL"
    abstract_num_xml = f'''
    <w:abstractNum w:abstractNumId="0" xmlns:w="{WNS}">
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="360" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
                <w:b w:val="0"/>
                <w:bCs w:val="0"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
                <w:lang w:bidi="he-IL"/>
            </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="1">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1.%2"/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
                <w:b w:val="0"/>
                <w:bCs w:val="0"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
                <w:lang w:bidi="he-IL"/>
            </w:rPr>
        </w:lvl>
    </w:abstractNum>
    '''
    abstract_num = etree.fromstring(abstract_num_xml)
    numbering_elm.insert(0, abstract_num)

    num_xml = f'''
    <w:num w:numId="2" xmlns:w="{WNS}">
        <w:abstractNumId w:val="0"/>
    </w:num>
    '''
    num_elem = etree.fromstring(num_xml)
    numbering_elm.append(num_elem)

    # ── Helper Functions ─────────────────────────────────────────────────

    def _set_rtl_bidi(p):
        """Set RTL and bidi on a paragraph element."""
        pPr = p._element.get_or_add_pPr()
        if pPr.find(qn('w:bidi')) is None:
            etree.SubElement(pPr, qn('w:bidi'))

    def _set_run_font(run, size=12, bold=False, underline=False, font_name='David'):
        """Configure run font properties including complex script."""
        run.font.name = font_name
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.underline = underline
        run.font.rtl = True
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
        rFonts.set(qn('w:cs'), font_name)
        rFonts.set(qn('w:eastAsia'), font_name)
        # bCs for bold complex script
        if bold:
            bCs = rPr.find(qn('w:bCs'))
            if bCs is None:
                etree.SubElement(rPr, qn('w:bCs'))
        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = etree.SubElement(rPr, qn('w:szCs'))
        szCs.set(qn('w:val'), str(size * 2))

    def _set_paragraph_spacing(p):
        """Set SKILL.md spacing: before=120, after=120, line=360, lineRule=auto."""
        pPr = p._element.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = etree.SubElement(pPr, qn('w:spacing'))
        sp.set(qn('w:before'), '120')
        sp.set(qn('w:after'), '120')
        sp.set(qn('w:line'), '360')
        sp.set(qn('w:lineRule'), 'auto')

    def _add_numbering(p, level=0):
        """Add auto-numbering to a paragraph with SKILL.md indentation."""
        pPr = p._element.get_or_add_pPr()
        numPr = etree.SubElement(pPr, qn('w:numPr'))
        ilvl = etree.SubElement(numPr, qn('w:ilvl'))
        ilvl.set(qn('w:val'), str(level))
        numId_el = etree.SubElement(numPr, qn('w:numId'))
        numId_el.set(qn('w:val'), '2')
        # SKILL.md indentation for numbered paras: left=-149, right=-709, hanging=425
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = etree.SubElement(pPr, qn('w:ind'))
        if level == 0:
            ind.set(qn('w:left'), '-149')
            ind.set(qn('w:right'), '-709')
            ind.set(qn('w:hanging'), '425')
        else:
            ind.set(qn('w:left'), '276')
            ind.set(qn('w:right'), '-709')
            ind.set(qn('w:hanging'), '425')

    def add_title(text):
        """Add the main title - centered, bold, large."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_rtl_bidi(p)
        _set_paragraph_spacing(p)
        run = p.add_run(text)
        _set_run_font(run, size=16, bold=True)
        return p

    def add_section_header(text):
        """Add a section header per SKILL.md: bold+underline, NOT numbered, ind left=-716 right=-709 firstLine=6."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_rtl_bidi(p)
        _set_paragraph_spacing(p)
        # SKILL.md section header indentation
        pPr = p._element.get_or_add_pPr()
        ind = etree.SubElement(pPr, qn('w:ind'))
        ind.set(qn('w:left'), '-716')
        ind.set(qn('w:right'), '-709')
        ind.set(qn('w:firstLine'), '6')
        run = p.add_run(text)
        _set_run_font(run, size=12, bold=True, underline=True)
        return p

    def add_numbered_para(text, level=0):
        """Add a numbered body paragraph per SKILL.md."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_rtl_bidi(p)
        _set_paragraph_spacing(p)
        _add_numbering(p, level=level)
        run = p.add_run(text)
        _set_run_font(run, size=12)
        return p

    def add_plain_para(text, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, size=12,
                       bold=False):
        """Add a plain (non-numbered) paragraph with SKILL.md spacing."""
        p = doc.add_paragraph()
        p.alignment = alignment
        _set_rtl_bidi(p)
        _set_paragraph_spacing(p)
        if text:
            run = p.add_run(text)
            _set_run_font(run, size=size, bold=bold)
        return p

    def add_appendix_ref(text):
        """Add appendix reference per SKILL.md: ◄ symbol, bold+underlined, NOT numbered."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_rtl_bidi(p)
        _set_paragraph_spacing(p)
        pPr = p._element.get_or_add_pPr()
        ind = etree.SubElement(pPr, qn('w:ind'))
        ind.set(qn('w:left'), '-149')
        ind.set(qn('w:right'), '-709')
        # ◄ symbol run (bold, not underlined)
        arrow_run = p.add_run('◄  ')
        _set_run_font(arrow_run, size=12, bold=True, underline=False)
        # Text run (bold + underlined)
        text_run = p.add_run(text)
        _set_run_font(text_run, size=12, bold=True, underline=True)
        return p

    def add_calculation_line(text):
        """Add a calculation/formula line - not numbered."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_rtl_bidi(p)
        _set_paragraph_spacing(p)
        pPr = p._element.get_or_add_pPr()
        ind = etree.SubElement(pPr, qn('w:ind'))
        ind.set(qn('w:left'), '-149')
        ind.set(qn('w:right'), '-709')
        run = p.add_run(text)
        _set_run_font(run, size=12)
        return p

    def set_cell_rtl(cell, text, bold=False, size=12, alignment=WD_ALIGN_PARAGRAPH.RIGHT):
        """Set cell text with RTL formatting. No negative indents inside cells."""
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = alignment
        pPr = p._element.get_or_add_pPr()
        if pPr.find(qn('w:bidi')) is None:
            etree.SubElement(pPr, qn('w:bidi'))
        # Ensure no negative indents in cells
        ind = pPr.find(qn('w:ind'))
        if ind is not None:
            pPr.remove(ind)
        # Compact spacing inside table cells
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = etree.SubElement(pPr, qn('w:spacing'))
        sp.set(qn('w:before'), '40')
        sp.set(qn('w:after'), '40')
        sp.set(qn('w:line'), '276')
        sp.set(qn('w:lineRule'), 'auto')
        if text:
            for line_idx, line in enumerate(text.split('\n')):
                if line_idx > 0:
                    run = p.add_run()
                    run.add_break()
                run = p.add_run(line)
                _set_run_font(run, size=size, bold=bold)
        tc = cell._element
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = etree.SubElement(tc, qn('w:tcPr'))
            tc.insert(0, tcPr)

    def _make_table_borderless(table):
        """Remove all borders from a table."""
        tbl = table._element
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, qn('w:tblPr'))
        tblBorders = etree.SubElement(tblPr, qn('w:tblBorders'))
        for bn in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            b = etree.SubElement(tblBorders, qn(f'w:{bn}'))
            b.set(qn('w:val'), 'none')
            b.set(qn('w:sz'), '0')
            b.set(qn('w:space'), '0')
            b.set(qn('w:color'), 'auto')
        return tblPr

    def _set_table_bidi(tblPr):
        """Add bidiVisual BEFORE tblW per SKILL.md."""
        # Remove existing bidiVisual if any
        for existing in tblPr.findall(qn('w:bidiVisual')):
            tblPr.remove(existing)
        bidi = etree.SubElement(tblPr, qn('w:bidiVisual'))
        # Move bidiVisual to be before tblW
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is not None:
            tblPr.remove(bidi)
            tblPr.insert(list(tblPr).index(tblW), bidi)
        else:
            # bidiVisual is already at the end, which is fine if no tblW
            pass

    def add_summary_table(claims_dict, total_amount):
        """Add a 2-column summary table with header row, visible borders, blue header."""
        num_rows = len(claims_dict) + 2  # +1 header row, +1 total row
        tbl = doc.add_table(rows=num_rows, cols=2)

        # Set table properties
        tblEl = tbl._element
        tblPr = tblEl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = etree.SubElement(tblEl, qn('w:tblPr'))

        # bidiVisual BEFORE tblW
        etree.SubElement(tblPr, qn('w:bidiVisual'))
        tblW = etree.SubElement(tblPr, qn('w:tblW'))
        tblW.set(qn('w:type'), 'dxa')
        tblW.set(qn('w:w'), '9026')

        # Grid columns: right wider (5513), left narrower (3513)
        tblGrid = tblEl.find(qn('w:tblGrid'))
        if tblGrid is None:
            tblGrid = etree.SubElement(tblEl, qn('w:tblGrid'))
        else:
            for gc in tblGrid.findall(qn('w:gridCol')):
                tblGrid.remove(gc)
        gc1 = etree.SubElement(tblGrid, qn('w:gridCol'))
        gc1.set(qn('w:w'), '5513')
        gc2 = etree.SubElement(tblGrid, qn('w:gridCol'))
        gc2.set(qn('w:w'), '3513')

        # Table borders (all single, visible)
        tblBorders = etree.SubElement(tblPr, qn('w:tblBorders'))
        for bn in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            b = etree.SubElement(tblBorders, qn(f'w:{bn}'))
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'), '4')
            b.set(qn('w:space'), '0')
            b.set(qn('w:color'), '000000')

        # Row 0: Header row (blue background, white text)
        set_cell_rtl(tbl.rows[0].cells[0], 'רכיב תביעה', bold=True, size=12,
                     alignment=WD_ALIGN_PARAGRAPH.RIGHT)
        set_cell_rtl(tbl.rows[0].cells[1], 'סכום (₪)', bold=True, size=12,
                     alignment=WD_ALIGN_PARAGRAPH.LEFT)

        # Shade header row blue (1A365D)
        def _shade_cell(cell, fill_color, font_color=None):
            tc = cell._element
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = etree.SubElement(tc, qn('w:tcPr'))
                tc.insert(0, tcPr)
            shd = etree.SubElement(tcPr, qn('w:shd'))
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), fill_color)
            if font_color:
                for run in cell.paragraphs[0].runs:
                    rPr = run._element.get_or_add_rPr()
                    color = rPr.find(qn('w:color'))
                    if color is None:
                        color = etree.SubElement(rPr, qn('w:color'))
                    color.set(qn('w:val'), font_color)

        _shade_cell(tbl.rows[0].cells[0], '1A365D', 'FFFFFF')
        _shade_cell(tbl.rows[0].cells[1], '1A365D', 'FFFFFF')

        # Data rows: right col = component name (bold, right-aligned), left col = amount
        for i, (key, claim) in enumerate(claims_dict.items()):
            row_idx = i + 1
            set_cell_rtl(tbl.rows[row_idx].cells[0], claim['name'], bold=True, size=12,
                         alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            set_cell_rtl(tbl.rows[row_idx].cells[1], f"{claim['amount']:,.0f} ₪", size=12,
                         alignment=WD_ALIGN_PARAGRAPH.LEFT)

        # Total row (last) - shaded light blue
        last_row = num_rows - 1
        set_cell_rtl(tbl.rows[last_row].cells[0], 'סה"כ', bold=True, size=12,
                     alignment=WD_ALIGN_PARAGRAPH.RIGHT)
        set_cell_rtl(tbl.rows[last_row].cells[1], f"{total_amount:,.0f} ₪", bold=True, size=12,
                     alignment=WD_ALIGN_PARAGRAPH.LEFT)
        _shade_cell(tbl.rows[last_row].cells[0], 'D9E2F3')
        _shade_cell(tbl.rows[last_row].cells[1], 'D9E2F3')

        return tbl

    # ── Data Extraction ──────────────────────────────────────────────────
    plaintiff_name = data.get("plaintiff_name", "")
    plaintiff_id = data.get("plaintiff_id", "")
    plaintiff_address = data.get("plaintiff_address", "")
    defendant_name = data.get("defendant_name", "")
    defendant_id = data.get("defendant_id", "")
    defendant_address = data.get("defendant_address", "")
    court_name = data.get("court_header", "בית הדין האזורי לעבודה בתל אביב")
    gender = data.get("gender", "male")
    pronoun = "התובע" if gender == "male" else "התובעת"
    total = calculations["total"]
    claims = calculations["claims"]
    attorney_name = data.get("attorney_name", "")
    attorney_id = data.get("attorney_id", "")
    firm_name = data.get("firm_name", "")
    firm_address = data.get("firm_address", "")
    firm_phone = data.get("firm_phone", "")
    firm_fax = data.get("firm_fax", "")
    firm_email = data.get("firm_email", "")

    defendant_label = "הנתבע" if data.get("defendant_type") == "individual" else "הנתבעת"

    # ── Helper: add multiple paragraphs to a cell ────────────────────────
    def set_cell_multiline(cell, lines_spec):
        """Set cell with multiple paragraphs. lines_spec: list of (text, bold, size, alignment) tuples."""
        cell.text = ''
        for idx, (text, bold, size, alignment) in enumerate(lines_spec):
            if idx == 0:
                p = cell.paragraphs[0]
            else:
                p = cell.add_paragraph()
            p.alignment = alignment
            pPr = p._element.get_or_add_pPr()
            if pPr.find(qn('w:bidi')) is None:
                etree.SubElement(pPr, qn('w:bidi'))
            # Remove negative indents
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                pPr.remove(ind)
            # Compact spacing
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = etree.SubElement(pPr, qn('w:spacing'))
            sp.set(qn('w:before'), '20')
            sp.set(qn('w:after'), '20')
            sp.set(qn('w:line'), '240')
            sp.set(qn('w:lineRule'), 'auto')
            if text:
                run = p.add_run(text)
                _set_run_font(run, size=size, bold=bold)

    # ══════════════════════════════════════════════════════════════════════
    # BUILD THE DOCUMENT — Cover Page (Enbar Shachar format)
    # ══════════════════════════════════════════════════════════════════════

    # ── Table 1: Top Header (INVISIBLE borders) ──────────────────────────
    # 2 columns: RIGHT = סע"ש/בפני, LEFT = court name
    # In bidiVisual RTL table: cell[0] renders on RIGHT, cell[1] on LEFT
    hdr_tbl = doc.add_table(rows=1, cols=2)
    hdr_el = hdr_tbl._element
    hdr_tblPr = hdr_el.find(qn('w:tblPr'))
    if hdr_tblPr is None:
        hdr_tblPr = etree.SubElement(hdr_el, qn('w:tblPr'))

    etree.SubElement(hdr_tblPr, qn('w:bidiVisual'))
    hdr_tblW = etree.SubElement(hdr_tblPr, qn('w:tblW'))
    hdr_tblW.set(qn('w:type'), 'dxa')
    hdr_tblW.set(qn('w:w'), '9026')

    # Invisible borders
    _make_table_borderless(hdr_tbl)

    hdr_grid = hdr_el.find(qn('w:tblGrid'))
    if hdr_grid is None:
        hdr_grid = etree.SubElement(hdr_el, qn('w:tblGrid'))
    else:
        for gc in hdr_grid.findall(qn('w:gridCol')):
            hdr_grid.remove(gc)
    for w in ['4513', '4513']:
        gc = etree.SubElement(hdr_grid, qn('w:gridCol'))
        gc.set(qn('w:w'), w)

    # cell[0] = RIGHT side: סע"ש and בפני
    set_cell_multiline(hdr_tbl.rows[0].cells[0], [
        ('סע"ש ________', False, 11, WD_ALIGN_PARAGRAPH.RIGHT),
        ('בפני _________', False, 11, WD_ALIGN_PARAGRAPH.RIGHT),
    ])
    # cell[1] = LEFT side: court name
    set_cell_rtl(hdr_tbl.rows[0].cells[1], court_name, bold=True, size=12,
                 alignment=WD_ALIGN_PARAGRAPH.LEFT)

    # ── Table 2: Parties Section (VISIBLE borders) ───────────────────────
    # Rows: בעניין label, plaintiff, נגד, defendant, מהות/סכום
    # 2 columns: col0=content (wide), col1=label (narrow)
    parties_tbl = doc.add_table(rows=5, cols=2)
    pt_el = parties_tbl._element
    pt_tblPr = pt_el.find(qn('w:tblPr'))
    if pt_tblPr is None:
        pt_tblPr = etree.SubElement(pt_el, qn('w:tblPr'))

    etree.SubElement(pt_tblPr, qn('w:bidiVisual'))
    pt_tblW = etree.SubElement(pt_tblPr, qn('w:tblW'))
    pt_tblW.set(qn('w:type'), 'dxa')
    pt_tblW.set(qn('w:w'), '9026')

    # Visible borders on all sides
    pt_borders = etree.SubElement(pt_tblPr, qn('w:tblBorders'))
    for bn in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = etree.SubElement(pt_borders, qn(f'w:{bn}'))
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')

    pt_grid = pt_el.find(qn('w:tblGrid'))
    if pt_grid is None:
        pt_grid = etree.SubElement(pt_el, qn('w:tblGrid'))
    else:
        for gc in pt_grid.findall(qn('w:gridCol')):
            pt_grid.remove(gc)
    for w in ['7026', '2000']:
        gc = etree.SubElement(pt_grid, qn('w:gridCol'))
        gc.set(qn('w:w'), w)

    # Row 0: "בעניין:" label (right-aligned, bold) — merged across both cols
    set_cell_rtl(parties_tbl.rows[0].cells[0], 'בעניין:', bold=True, size=12,
                 alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    set_cell_rtl(parties_tbl.rows[0].cells[1], '', size=11)

    # Row 1: Plaintiff details (col0) | label (col1)
    plaintiff_lines = []
    name_id = f'{plaintiff_name}, ת.ז. {plaintiff_id}' if plaintiff_id else plaintiff_name
    plaintiff_lines.append((name_id, True, 12, WD_ALIGN_PARAGRAPH.RIGHT))
    if plaintiff_address:
        plaintiff_lines.append((plaintiff_address, False, 11, WD_ALIGN_PARAGRAPH.RIGHT))
    if attorney_name:
        plaintiff_lines.append((f'באמצעות ב"כ עוה"ד {attorney_name}', False, 11, WD_ALIGN_PARAGRAPH.RIGHT))
    if firm_name:
        plaintiff_lines.append((firm_name, False, 11, WD_ALIGN_PARAGRAPH.RIGHT))
    if firm_address:
        plaintiff_lines.append((firm_address, False, 11, WD_ALIGN_PARAGRAPH.RIGHT))
    contact_parts = []
    if firm_phone:
        contact_parts.append(f"טל': {firm_phone}")
    if firm_fax:
        contact_parts.append(f"פקס': {firm_fax}")
    if contact_parts:
        plaintiff_lines.append((' '.join(contact_parts), False, 11, WD_ALIGN_PARAGRAPH.RIGHT))
    if firm_email:
        plaintiff_lines.append((firm_email, False, 11, WD_ALIGN_PARAGRAPH.RIGHT))

    set_cell_multiline(parties_tbl.rows[1].cells[0], plaintiff_lines)
    # Label: bold "התובע/ת" right-aligned
    set_cell_rtl(parties_tbl.rows[1].cells[1], pronoun, bold=True, size=12,
                 alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    # Row 2: "- נגד -" centered across both cols
    set_cell_rtl(parties_tbl.rows[2].cells[0], '- נגד -', bold=True, size=12,
                 alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_rtl(parties_tbl.rows[2].cells[1], '', size=11)

    # Row 3: Defendant details (col0) | label (col1)
    defendant_lines = []
    defendant_lines.append((defendant_name, True, 12, WD_ALIGN_PARAGRAPH.RIGHT))
    if defendant_id:
        defendant_lines.append((f'ח.פ {defendant_id}', False, 11, WD_ALIGN_PARAGRAPH.RIGHT))
    if defendant_address:
        defendant_lines.append((defendant_address, False, 11, WD_ALIGN_PARAGRAPH.RIGHT))

    set_cell_multiline(parties_tbl.rows[3].cells[0], defendant_lines)
    set_cell_rtl(parties_tbl.rows[3].cells[1], defendant_label, bold=True, size=12,
                 alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    # Row 4: מהות התביעה / סכום התביעה in a single row
    amount_str = f'{total:,.0f} ₪'
    set_cell_multiline(parties_tbl.rows[4].cells[0], [
        (f'מהות התביעה: הצהרתית וכספית', True, 11, WD_ALIGN_PARAGRAPH.RIGHT),
        (f'סכום התביעה: {amount_str}', True, 11, WD_ALIGN_PARAGRAPH.RIGHT),
    ])
    set_cell_rtl(parties_tbl.rows[4].cells[1], '', size=11)

    # ── Header Summary Table (financial breakdown) ───────────────────────
    add_plain_para('')
    add_summary_table(claims, total)

    # ── Title ────────────────────────────────────────────────────────────
    add_title('כ ת ב    ת ב י ע ה')

    # ── Body - Parse claim_text and format properly ──────────────────────
    section_headers = {
        "כללי", "הצדדים", "רקע עובדתי", "היקף משרה ושכר קובע",
        "רכיבי התביעה", "סיכום",
        "שכר עבודה שלא שולם", "הפרשי שכר – שעות נוספות",
        "הפרשי הפרשות לפנסיה", "פיצויי פיטורים",
        "הפרשי שכר דמי חופשה ופדיון חופשה",
        "דמי חגים והפרשי דמי חג", "דמי הבראה",
        "ניכויים שלא כדין – תגמולי עובד", "פיצויי הלנת שכר",
        "פיצוי בגין עוגמת נפש", "מסירת מסמכי גמר חשבון",
        "עילות התביעה", "הסעדים המבוקשים",
        "תחשיב שעות נוספות שהיה צריך לשלם בכל חודש:",
    }

    lines = claim_text.split("\n")
    in_summary = False
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        elif stripped == "כ ת ב    ת ב י ע ה":
            continue
        elif stripped == "סיכום רכיבי התביעה:":
            add_section_header(stripped)
            in_summary = True
            continue
        elif in_summary and stripped.startswith("•"):
            continue  # Skip bullet items; we'll use summary table instead
        elif in_summary and not stripped.startswith("•") and "סה\"כ סכום התביעה" not in stripped:
            in_summary = False
            # Fall through to normal processing
        elif "סה\"כ סכום התביעה" in stripped:
            continue  # Skip; shown in summary table
        if stripped in section_headers:
            add_section_header(stripped)
        elif stripped.startswith("תלושי שכר") and "נספח" in stripped:
            add_appendix_ref(stripped)
        elif any(c in stripped for c in ['=', '×']) and '₪' in stripped:
            add_calculation_line(stripped)
        else:
            add_numbered_para(stripped)

    # ── End Summary Table (must match header summary) ────────────────────
    add_section_header("סיכום רכיבי התביעה")
    add_summary_table(claims, total)

    add_plain_para(
        f'סה"כ סכום התביעה: {total:,.0f} ₪ קרן (לא כולל הצמדה וריבית, שכ"ט עו"ד והוצאות)',
        bold=True
    )

    # Final legal paragraphs from claim text (after the summary section)
    found_summary_end = False
    for line in lines:
        stripped = line.strip()
        if "סה\"כ סכום התביעה" in stripped:
            found_summary_end = True
            continue
        if found_summary_end and stripped:
            add_numbered_para(stripped)

    # ── Power of Attorney Note ───────────────────────────────────────────
    add_appendix_ref('ייפוי כוח מצורף לכתב התביעה')

    # ── Signature Table (2-col: spacer 5649 + sig 3377, per SKILL.md) ────
    add_plain_para('')

    sig_table = doc.add_table(rows=1, cols=2)
    sig_tbl_el = sig_table._element
    sig_tblPr = sig_tbl_el.find(qn('w:tblPr'))
    if sig_tblPr is None:
        sig_tblPr = etree.SubElement(sig_tbl_el, qn('w:tblPr'))

    # bidiVisual BEFORE tblW
    sig_bidi = etree.SubElement(sig_tblPr, qn('w:bidiVisual'))
    sig_tblW = etree.SubElement(sig_tblPr, qn('w:tblW'))
    sig_tblW.set(qn('w:type'), 'dxa')
    sig_tblW.set(qn('w:w'), '9026')

    # Remove borders
    sig_borders = etree.SubElement(sig_tblPr, qn('w:tblBorders'))
    for bn in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = etree.SubElement(sig_borders, qn(f'w:{bn}'))
        b.set(qn('w:val'), 'none')
        b.set(qn('w:sz'), '0')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), 'auto')

    # Grid: spacer 5649 + sig 3377
    sig_grid = sig_tbl_el.find(qn('w:tblGrid'))
    if sig_grid is None:
        sig_grid = etree.SubElement(sig_tbl_el, qn('w:tblGrid'))
    else:
        for gc in sig_grid.findall(qn('w:gridCol')):
            sig_grid.remove(gc)
    gc1 = etree.SubElement(sig_grid, qn('w:gridCol'))
    gc1.set(qn('w:w'), '5649')
    gc2 = etree.SubElement(sig_grid, qn('w:gridCol'))
    gc2.set(qn('w:w'), '3377')

    # Spacer cell (empty)
    set_cell_rtl(sig_table.rows[0].cells[0], '', size=12)

    # Signature cell with top border (signature line)
    sig_cell = sig_table.rows[0].cells[1]
    if attorney_name and attorney_id:
        sig_text = f'{attorney_name}, עו"ד\nמ.ר. {attorney_id}\nב"כ {pronoun}'
    else:
        sig_text = f'__________________\nב"כ {pronoun}'
    set_cell_rtl(sig_cell, sig_text, size=12, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    # Add top border to signature cell (serves as signature line)
    sig_tc = sig_cell._element
    sig_tcPr = sig_tc.find(qn('w:tcPr'))
    if sig_tcPr is None:
        sig_tcPr = etree.SubElement(sig_tc, qn('w:tcPr'))
        sig_tc.insert(0, sig_tcPr)
    sig_tcBorders = etree.SubElement(sig_tcPr, qn('w:tcBorders'))
    top_border = etree.SubElement(sig_tcBorders, qn('w:top'))
    top_border.set(qn('w:val'), 'single')
    top_border.set(qn('w:sz'), '4')
    top_border.set(qn('w:space'), '0')
    top_border.set(qn('w:color'), 'auto')

    return doc


# ── Flask Routes ─────────────────────────────────────────────────────────────

@app.before_request
def require_login():
    allowed = ("login", "static", "service_worker", "manifest")
    if request.endpoint not in allowed and not session.get("authenticated"):
        return redirect(url_for("login"))


@app.route("/sw.js")
def service_worker():
    return app.send_static_file("sw.js"), 200, {"Content-Type": "application/javascript", "Service-Worker-Allowed": "/"}


@app.route("/manifest.json")
def manifest():
    return app.send_static_file("manifest.json"), 200, {"Content-Type": "application/manifest+json"}


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        password = request.form.get("password", "")
        if password == APP_PASSWORD:
            session.permanent = True
            session["authenticated"] = True
            return redirect(url_for("index"))
        else:
            error = "סיסמה שגויה, נסה שוב"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/calculate", methods=["POST"])
def calculate():
    data = request.json
    try:
        calculations = calculate_all_claims(data)
        claim_text = generate_claim_text(data, calculations)
        return jsonify({
            "success": True,
            "calculations": calculations,
            "claim_text": claim_text,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route("/generate-docx", methods=["POST"])
def generate_docx_route():
    data = request.json
    try:
        calculations = calculate_all_claims(data)
        claim_text = generate_claim_text(data, calculations)
        doc = generate_docx(data, calculations, claim_text)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        plaintiff = data.get("plaintiff_name", "claim")
        filename = f"כתב_תביעה_{plaintiff}.docx"

        return send_file(
            buffer,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


if __name__ == "__main__":
    import os
    debug = os.environ.get("FLASK_DEBUG", "true").lower() == "true"
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=debug, port=port, host="0.0.0.0")
