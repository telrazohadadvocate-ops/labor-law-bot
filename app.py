"""
Labor Law Claim Generator Bot - Levin Telraz Law Firm
Generates Israeli labor law claims (כתבי תביעה) based on client intake data.
"""

import json
import math
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from flask import Flask, render_template, request, jsonify, send_file
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os
import io

app = Flask(__name__)

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
    """Calculate overtime pay owed (שעות נוספות)."""
    rate_125 = hourly_rate * 1.25
    rate_150 = hourly_rate * 1.50
    surcharge_125 = hourly_rate * OVERTIME_125_RATE
    surcharge_150 = hourly_rate * OVERTIME_150_RATE

    monthly_125 = weekly_overtime_125 * 4 * surcharge_125
    monthly_150 = weekly_overtime_150 * 4 * surcharge_150

    total = round((monthly_125 + monthly_150) * months, 2)
    return {
        "monthly_125": round(monthly_125, 2),
        "monthly_150": round(monthly_150, 2),
        "total": total,
        "rate_125": round(rate_125, 2),
        "rate_150": round(rate_150, 2),
        "surcharge_125": round(surcharge_125, 2),
        "surcharge_150": round(surcharge_150, 2),
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
            "amount": round(severance - deposited, 2),
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
    """Generate a Word document matching the exact format of Levin Telraz firm claims.

    Based on analysis of the real כתב תביעה documents:
    - Font: David 12pt throughout
    - Body paragraphs: Justified, RTL (bidi), 1.5 line spacing, auto-numbered
    - Section headers: Bold + Underline, not numbered
    - Title: Heading 4 style, centered
    - Margins: ~1.8cm left/right, small top/bottom
    - Court header table at top with parties, case type, and amount
    - Signature table at bottom with attorney details
    """
    from docx.shared import Emu
    from copy import deepcopy
    from lxml import etree

    doc = Document()
    WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # ── Page Setup ───────────────────────────────────────────────────────
    for section in doc.sections:
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)
        section.top_margin = Cm(0.7)
        section.bottom_margin = Cm(1.3)

    # ── Configure Default Styles ─────────────────────────────────────────
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'David'
    style_normal.font.size = Pt(12)
    style_normal.font.rtl = True
    pf = style_normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # Set bidi on Normal style
    style_pPr = style_normal.element.get_or_add_pPr()
    etree.SubElement(style_pPr, qn('w:bidi'))

    # ── Create Numbering (auto-numbered paragraphs like the real doc) ────
    # We need to inject numbering definitions into the document.
    # Force creation of numbering part by adding and removing a list paragraph.
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    try:
        numbering_part = doc.part.numbering_part
    except Exception:
        # Create a dummy list item to force numbering part creation
        dummy = doc.add_paragraph('', style='List Number')
        numbering_part = doc.part.numbering_part
        # Remove the dummy paragraph
        dummy._element.getparent().remove(dummy._element)
    numbering_elm = numbering_part.element

    # Abstract numbering definition for Hebrew decimal numbers
    abstract_num_xml = f'''
    <w:abstractNum w:abstractNumId="1" xmlns:w="{WNS}">
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
        </w:lvl>
        <w:lvl w:ilvl="1">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1.%2"/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="1440" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
        </w:lvl>
    </w:abstractNum>
    '''
    abstract_num = etree.fromstring(abstract_num_xml)
    numbering_elm.insert(0, abstract_num)

    num_xml = f'''
    <w:num w:numId="1" xmlns:w="{WNS}">
        <w:abstractNumId w:val="1"/>
    </w:num>
    '''
    num_elem = etree.fromstring(num_xml)
    numbering_elm.append(num_elem)

    # ── Helper Functions ─────────────────────────────────────────────────
    paragraph_counter = [0]  # mutable counter for numbering

    def _set_rtl_bidi(p):
        """Set RTL and bidi on a paragraph element."""
        pPr = p._element.get_or_add_pPr()
        # Add bidi if not present
        if pPr.find(qn('w:bidi')) is None:
            etree.SubElement(pPr, qn('w:bidi'))

    def _set_run_font(run, size=12, bold=False, underline=False, font_name='David'):
        """Configure run font properties."""
        run.font.name = font_name
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.underline = underline
        run.font.rtl = True
        # Set complex script font (for Hebrew)
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
        rFonts.set(qn('w:cs'), font_name)
        # Set cs size
        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = etree.SubElement(rPr, qn('w:szCs'))
        szCs.set(qn('w:val'), str(size * 2))

    def _add_numbering(p, level=0):
        """Add auto-numbering to a paragraph."""
        pPr = p._element.get_or_add_pPr()
        numPr = etree.SubElement(pPr, qn('w:numPr'))
        ilvl = etree.SubElement(numPr, qn('w:ilvl'))
        ilvl.set(qn('w:val'), str(level))
        numId = etree.SubElement(numPr, qn('w:numId'))
        numId.set(qn('w:val'), '1')

    def add_title(text):
        """Add the main title (כתב תביעה) - centered, bold, large."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_rtl_bidi(p)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(text)
        _set_run_font(run, size=16, bold=True)
        return p

    def add_section_header(text):
        """Add a section header - bold, underlined, justified, not numbered."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_rtl_bidi(p)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        _set_run_font(run, size=12, bold=True, underline=True)
        return p

    def add_numbered_para(text, level=0):
        """Add a numbered body paragraph - justified, RTL, 1.5 spacing."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_rtl_bidi(p)
        _add_numbering(p, level=level)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        _set_run_font(run, size=12)
        return p

    def add_plain_para(text, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, size=12,
                       bold=False, spacing_before=6, spacing_after=6):
        """Add a plain (non-numbered) paragraph."""
        p = doc.add_paragraph()
        p.alignment = alignment
        _set_rtl_bidi(p)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(spacing_before)
        p.paragraph_format.space_after = Pt(spacing_after)
        if text:
            run = p.add_run(text)
            _set_run_font(run, size=size, bold=bold)
        return p

    def add_calculation_line(text):
        """Add a calculation/formula line - not numbered, indented."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_rtl_bidi(p)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        # Indent
        pPr = p._element.get_or_add_pPr()
        ind = etree.SubElement(pPr, qn('w:ind'))
        ind.set(qn('w:right'), '720')
        run = p.add_run(text)
        _set_run_font(run, size=12)
        return p

    # ── Helper to build RTL table cell ───────────────────────────────────
    def set_cell_rtl(cell, text, bold=False, size=12, alignment=WD_ALIGN_PARAGRAPH.RIGHT):
        """Set cell text with RTL formatting."""
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = alignment
        _set_rtl_bidi(p)
        if text:
            for line_idx, line in enumerate(text.split('\n')):
                if line_idx > 0:
                    run = p.add_run()
                    run.add_break()
                run = p.add_run(line)
                _set_run_font(run, size=size, bold=bold)
        # Set cell direction
        tc = cell._element
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = etree.SubElement(tc, qn('w:tcPr'))
            tc.insert(0, tcPr)

    # ── Data Extraction ──────────────────────────────────────────────────
    plaintiff_name = data.get("plaintiff_name", "")
    defendant_name = data.get("defendant_name", "")
    court_name = data.get("court_header", "בית הדין האזורי לעבודה")
    gender = data.get("gender", "male")
    pronoun = "התובע" if gender == "male" else "התובעת"
    total = calculations["total"]
    attorney_name = data.get("attorney_name", "")
    attorney_id = data.get("attorney_id", "")

    # ══════════════════════════════════════════════════════════════════════
    # BUILD THE DOCUMENT
    # ══════════════════════════════════════════════════════════════════════

    # ── Court Header Table (top of document) ─────────────────────────────
    # Format: בס"ד | court name on right
    header_table = doc.add_table(rows=4, cols=2)
    header_table.autofit = True

    # Row 0: בס"ד on left, empty right
    set_cell_rtl(header_table.rows[0].cells[0], 'בס"ד', size=11)
    set_cell_rtl(header_table.rows[0].cells[1], '', size=11)

    # Row 1: Court name centered across
    set_cell_rtl(header_table.rows[1].cells[0], court_name, bold=True, size=13,
                 alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_rtl(header_table.rows[1].cells[1], '', size=11)
    # Merge row 1 cells
    header_table.rows[1].cells[0].merge(header_table.rows[1].cells[1])

    # Row 2: Case type
    set_cell_rtl(header_table.rows[2].cells[0], 'מהות התביעה: הצהרתית וכספית',
                 size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_rtl(header_table.rows[2].cells[1], '', size=11)
    header_table.rows[2].cells[0].merge(header_table.rows[2].cells[1])

    # Row 3: Amount
    amount_text = f'סכום התביעה: {total:,.0f} ₪ קרן (לא כולל הצמדה וריבית, שכ"ט עו"ד והוצאות)'
    set_cell_rtl(header_table.rows[3].cells[0], amount_text,
                 bold=True, size=11, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_rtl(header_table.rows[3].cells[1], '', size=11)
    header_table.rows[3].cells[0].merge(header_table.rows[3].cells[1])

    # Remove table borders
    tbl = header_table._element
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = etree.SubElement(tbl, qn('w:tblPr'))
    tblBorders = etree.SubElement(tblPr, qn('w:tblBorders'))
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = etree.SubElement(tblBorders, qn(f'w:{border_name}'))
        border.set(qn('w:val'), 'none')
        border.set(qn('w:sz'), '0')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')

    # Set RTL on table
    bidi_tbl = etree.SubElement(tblPr, qn('w:bidiVisual'))

    add_plain_para('', spacing_before=2, spacing_after=2)

    # ── Parties Block ────────────────────────────────────────────────────
    parties_table = doc.add_table(rows=5, cols=3)
    parties_table.autofit = True

    # Plaintiff
    set_cell_rtl(parties_table.rows[0].cells[2], plaintiff_name, bold=True, size=12)
    set_cell_rtl(parties_table.rows[0].cells[1], '', size=12)
    set_cell_rtl(parties_table.rows[0].cells[0], '', size=12)

    set_cell_rtl(parties_table.rows[1].cells[2], pronoun, size=12)
    set_cell_rtl(parties_table.rows[1].cells[1], '', size=12)
    set_cell_rtl(parties_table.rows[1].cells[0], '', size=12)

    # vs
    set_cell_rtl(parties_table.rows[2].cells[1], '- נ ג ד -', bold=True, size=12,
                 alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_rtl(parties_table.rows[2].cells[0], '', size=12)
    set_cell_rtl(parties_table.rows[2].cells[2], '', size=12)

    # Defendant
    set_cell_rtl(parties_table.rows[3].cells[2], defendant_name, bold=True, size=12)
    set_cell_rtl(parties_table.rows[3].cells[1], '', size=12)
    set_cell_rtl(parties_table.rows[3].cells[0], '', size=12)

    defendant_label = "הנתבע" if data.get("defendant_type") == "individual" else "הנתבעת"
    set_cell_rtl(parties_table.rows[4].cells[2], defendant_label, size=12)
    set_cell_rtl(parties_table.rows[4].cells[1], '', size=12)
    set_cell_rtl(parties_table.rows[4].cells[0], '', size=12)

    # Remove borders from parties table
    ptbl = parties_table._element
    ptblPr = ptbl.find(qn('w:tblPr'))
    if ptblPr is None:
        ptblPr = etree.SubElement(ptbl, qn('w:tblPr'))
    ptblBorders = etree.SubElement(ptblPr, qn('w:tblBorders'))
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = etree.SubElement(ptblBorders, qn(f'w:{border_name}'))
        border.set(qn('w:val'), 'none')
        border.set(qn('w:sz'), '0')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
    etree.SubElement(ptblPr, qn('w:bidiVisual'))

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
    }

    lines = claim_text.split("\n")
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue  # Skip empty lines (spacing handled by paragraph formatting)
        elif stripped == "כ ת ב    ת ב י ע ה":
            continue  # Already added as title
        elif stripped in section_headers or stripped == "סיכום רכיבי התביעה:":
            add_section_header(stripped)
        elif stripped.startswith("•"):
            # Bullet point in summary - use sub-numbering
            add_numbered_para(stripped.lstrip("• "), level=1)
        elif any(c in stripped for c in ['=', '*']) and '₪' in stripped and '(' in stripped:
            # Calculation formula line
            add_calculation_line(stripped)
        else:
            # Regular numbered paragraph
            add_numbered_para(stripped)

    # ── Signature Table ──────────────────────────────────────────────────
    add_plain_para('', spacing_before=6, spacing_after=2)

    # Power of attorney note
    add_plain_para('*ייפוי כוח מצורף לכתב התביעה', bold=True, size=11,
                   spacing_before=12, spacing_after=6)

    sig_table = doc.add_table(rows=1, cols=3)
    sig_table.autofit = True

    if attorney_name and attorney_id:
        sig_text = f'{attorney_name}, עו"ד\nמ.ר. {attorney_id}\nב"כ {pronoun}'
    else:
        sig_text = f'__________________\nב"כ {pronoun}'

    set_cell_rtl(sig_table.rows[0].cells[2], sig_text, size=12)
    set_cell_rtl(sig_table.rows[0].cells[1], '', size=12)
    set_cell_rtl(sig_table.rows[0].cells[0], '', size=12)

    # Remove borders from signature table
    stbl = sig_table._element
    stblPr = stbl.find(qn('w:tblPr'))
    if stblPr is None:
        stblPr = etree.SubElement(stbl, qn('w:tblPr'))
    stblBorders = etree.SubElement(stblPr, qn('w:tblBorders'))
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = etree.SubElement(stblBorders, qn(f'w:{border_name}'))
        border.set(qn('w:val'), 'none')
        border.set(qn('w:sz'), '0')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
    etree.SubElement(stblPr, qn('w:bidiVisual'))

    return doc


# ── Flask Routes ─────────────────────────────────────────────────────────────

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
