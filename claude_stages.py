"""
Single-call AI generation for כתב תביעה documents.

One Claude API call. If it fails, caller falls back to template mode.
"""

import json
import re
import logging
import traceback

import anthropic

MODEL = "claude-haiku-4-5-20251001"
MAX_TOKENS = 3000
API_TIMEOUT = 30.0

SYSTEM_PROMPT_MALE = """אתה עורך דין ישראלי לדיני עבודה. כתוב כתב תביעה מלא בעברית משפטית רשמית.

כללים:
- כתוב הכל בעברית בלבד. אסור אנגלית בשום מקום.
- השתמש בגוף שלישי זכר: התובע, הועסק, פוטר, זכאי, עובד, טוען, יבקש, שכרו, עבודתו, זכויותיו
- השתמש בסכומים המדויקים מהנתונים
- אל תמציא עובדות שלא סופקו
- שלב את העובדות הגולמיות בתוך סעיף רקע עובדתי
- הפנה לחוקים ספציפיים כשרלוונטי

החזר JSON בלבד:
{"sections":[{"title":"כותרת הסעיף","content":"תוכן הסעיף בעברית"}],"appendices":["תיאור נספח"]}"""

SYSTEM_PROMPT_FEMALE = """את עורכת דין ישראלית לדיני עבודה. כתבי כתב תביעה מלא בעברית משפטית רשמית.

כללים:
- כתבי הכל בעברית בלבד. אסור אנגלית בשום מקום.
- השתמשי בגוף שלישי נקבה: התובעת, הועסקה, פוטרה, זכאית, עובדת, טוענת, תבקש, שכרה, עבודתה, זכויותיה
- השתמשי בסכומים המדויקים מהנתונים
- אל תמציאי עובדות שלא סופקו
- שלבי את העובדות הגולמיות בתוך סעיף רקע עובדתי
- הפני לחוקים ספציפיים כשרלוונטי

החזירי JSON בלבד:
{"sections":[{"title":"כותרת הסעיף","content":"תוכן הסעיף בעברית"}],"appendices":["תיאור נספח"]}"""


def generate_claim_single(raw_input, structured_data, calculations,
                          firm_patterns=None, legal_citations=None,
                          api_key=None):
    """Single Claude API call to generate כתב תביעה.

    Returns:
        Dict with sections, appendices, etc. Or None on total failure.
    """
    if not api_key:
        logging.warning("generate_claim_single: no API key")
        return None

    client = anthropic.Anthropic(api_key=api_key, timeout=API_TIMEOUT)

    gender = structured_data.get("gender", "male")
    gender_label = "זכר" if gender == "male" else "נקבה"
    pronoun = "התובע" if gender == "male" else "התובעת"

    calc_lines = []
    for key, claim in calculations.get("claims", {}).items():
        line = f"- {claim['name']}: {claim['amount']:,.0f} ₪"
        if claim.get("formula"):
            line += f" ({claim['formula']})"
        calc_lines.append(line)

    termination_type = structured_data.get("termination_type", "fired")
    if termination_type == "fired":
        termination_he = "פוטר" if gender == "male" else "פוטרה"
    elif termination_type == "resigned_justified":
        termination_he = "התפטר בדין מפוטר" if gender == "male" else "התפטרה בדין מפוטרת"
    else:
        termination_he = "התפטר" if gender == "male" else "התפטרה"

    user_prompt = f"""נתוני התיק:
שם {pronoun}: {structured_data.get('plaintiff_name', '')}
ת.ז.: {structured_data.get('plaintiff_id', '')}
מין: {gender_label}
שם הנתבעת: {structured_data.get('defendant_name', '')}
ח.פ./ע.מ.: {structured_data.get('defendant_id', '')}
תפקיד: {structured_data.get('job_title', '')}
תחילת עבודה: {structured_data.get('start_date', '')}
סיום עבודה: {structured_data.get('end_date', '')}
אופן סיום: {termination_he}
שכר בסיס: {structured_data.get('base_salary', '')} ₪
עמלות/תוספות: {structured_data.get('commissions', '0')} ₪
שכר קובע: {calculations.get('determining_salary', 0):,.0f} ₪
תקופת העסקה: {calculations.get('duration', {}).get('total_months', 0)} חודשים ({calculations.get('duration', {}).get('decimal_years', 0)} שנים)

חישובים (סכומים מחייבים — השתמש בדיוק בסכומים אלה):
{chr(10).join(calc_lines)}
סה"כ: {calculations.get('total', 0):,.0f} ₪

עובדות גולמיות (חובה לשלב בסעיף רקע עובדתי):
{raw_input}

כתוב כתב תביעה מלא בעברית. הסעיפים הנדרשים:
1. כללי — הצהרות פרוצדורליות
2. הצדדים — פרטי {pronoun} והנתבעת
3. רקע עובדתי — שלב כאן את העובדות הגולמיות שלמעלה
4. היקף משרה ושכר קובע
5. רכיבי התביעה — סעיף נפרד לכל רכיב עם חישוב
6. סיכום

החזר JSON בלבד."""

    system = SYSTEM_PROMPT_MALE if gender == "male" else SYSTEM_PROMPT_FEMALE

    # Append firm style hints if available
    if firm_patterns and firm_patterns.get("patterns"):
        style_keys = firm_patterns["patterns"]
        if style_keys.get("opening_phrases"):
            system += f"\n\nדוגמאות פתיחה של המשרד: {json.dumps(style_keys['opening_phrases'][:3], ensure_ascii=False)}"
        if style_keys.get("closing_phrases"):
            system += f"\nדוגמאות סיום: {json.dumps(style_keys['closing_phrases'][:2], ensure_ascii=False)}"

    logging.info(f"Calling Claude API (model={MODEL}, max_tokens={MAX_TOKENS}, timeout={API_TIMEOUT}s)...")
    logging.info(f"User prompt length: {len(user_prompt)} chars, system prompt length: {len(system)} chars")

    try:
        message = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=system,
            messages=[{"role": "user", "content": user_prompt}],
        )
    except Exception as e:
        logging.error(f"Claude API call FAILED: {e}")
        logging.error(traceback.format_exc())
        return None

    raw_text = message.content[0].text.strip()
    logging.info(f"Claude response ({len(raw_text)} chars): {raw_text[:2000]}")
    if len(raw_text) > 2000:
        logging.info(f"Claude response tail: ...{raw_text[-500:]}")

    # Try to parse JSON
    parsed = _safe_parse_json(raw_text)
    if parsed is not None:
        logging.info(f"JSON parsed OK: {len(parsed.get('sections', []))} sections")
        return _normalize(parsed, structured_data, calculations)

    # JSON failed — wrap raw text as single section
    logging.warning("JSON parse failed — wrapping raw text as single section")
    return _wrap_raw_text(raw_text, structured_data, calculations)


def _normalize(parsed, structured_data, calculations):
    """Normalize the parsed response to a consistent format.

    Converts any section format to {"header": ..., "paragraphs": [...]}.
    Filters out English-only text artifacts.
    """
    sections = parsed.get("sections", [])
    gender = structured_data.get("gender", "male")

    normalized_sections = []
    for s in sections:
        # Extract header and body from various formats
        header = s.get("title") or s.get("header") or s.get("name") or ""
        body = s.get("content") or s.get("paragraphs") or s.get("text") or ""

        # Convert body to list of paragraphs
        if isinstance(body, str):
            paragraphs = [p.strip() for p in body.split("\n") if p.strip()]
        elif isinstance(body, list):
            paragraphs = [str(p).strip() for p in body if str(p).strip()]
        else:
            paragraphs = [str(body)]

        # Skip sections with no Hebrew content
        if not header and not paragraphs:
            continue

        # Filter out English-only paragraphs (JSON artifacts, template keys)
        clean_paragraphs = []
        for p in paragraphs:
            # Skip if purely English/ASCII (no Hebrew chars at all)
            if _is_english_only(p):
                logging.warning(f"Filtered English-only paragraph: {p[:100]}")
                continue
            # Fix gender-neutral forms
            p = _fix_gender(p, gender)
            clean_paragraphs.append(p)

        if header:
            header = _fix_gender(header, gender)

        if header or clean_paragraphs:
            normalized_sections.append({"header": header, "paragraphs": clean_paragraphs})

    appendices_raw = parsed.get("appendices", [])
    appendices = []
    for i, a in enumerate(appendices_raw):
        if isinstance(a, str):
            if not _is_english_only(a):
                appendices.append({"number": i + 1, "description": a, "reference_text": a})
        elif isinstance(a, dict):
            appendices.append(a)

    return {
        "gender_form": gender,
        "sections": normalized_sections,
        "appendices": appendices,
        "calculations": parsed.get("calculations", []),
        "legal_citations": parsed.get("legal_citations", []),
        "summary_total": calculations.get("total", 0),
        "verification_notes": [],
        "stage_timing": {"total_seconds": 0, "stages_completed": 1},
    }


def _is_english_only(text):
    """Return True if text contains no Hebrew characters at all."""
    # Hebrew Unicode range: \u0590-\u05FF
    return not bool(re.search(r'[\u0590-\u05FF]', text))


def _fix_gender(text, gender):
    """Replace gender-neutral slashed forms with the correct gender form."""
    if gender == "male":
        replacements = {
            "התובע/ת": "התובע",
            "זכאי/ת": "זכאי",
            "עובד/ת": "עובד",
            "הועסק/ה": "הועסק",
            "פוטר/ה": "פוטר",
            "מיוצג/ת": "מיוצג",
            "מגיש/ה": "מגיש",
            "טוען/ת": "טוען",
            "יבקש/תבקש": "יבקש",
            "שכרו/ה": "שכרו",
            "עבודתו/ה": "עבודתו",
            "זכויותיו/ה": "זכויותיו",
            "העסקתו/ה": "העסקתו",
            "נאלץ/ה": "נאלץ",
            "החל/ה": "החל",
            "ביצע/ה": "ביצע",
            "עבד/ה": "עבד",
            "היה/תה": "היה",
            "מצוין/ת": "מצוין",
            "מקצועי/ת": "מקצועי",
        }
    else:
        replacements = {
            "התובע/ת": "התובעת",
            "זכאי/ת": "זכאית",
            "עובד/ת": "עובדת",
            "הועסק/ה": "הועסקה",
            "פוטר/ה": "פוטרה",
            "מיוצג/ת": "מיוצגת",
            "מגיש/ה": "מגישה",
            "טוען/ת": "טוענת",
            "יבקש/תבקש": "תבקש",
            "שכרו/ה": "שכרה",
            "עבודתו/ה": "עבודתה",
            "זכויותיו/ה": "זכויותיה",
            "העסקתו/ה": "העסקתה",
            "נאלץ/ה": "נאלצה",
            "החל/ה": "החלה",
            "ביצע/ה": "ביצעה",
            "עבד/ה": "עבדה",
            "היה/תה": "היתה",
            "מצוין/ת": "מצוינת",
            "מקצועי/ת": "מקצועית",
        }

    for pattern, replacement in replacements.items():
        text = text.replace(pattern, replacement)

    return text


def _wrap_raw_text(raw_text, structured_data, calculations):
    """Wrap raw text as sections when JSON parsing fails.

    Filters out English-only lines.
    """
    gender = structured_data.get("gender", "male")
    paragraphs = []
    for p in raw_text.split("\n"):
        p = p.strip()
        if p and not _is_english_only(p):
            p = _fix_gender(p, gender)
            paragraphs.append(p)

    return {
        "gender_form": gender,
        "sections": [{"header": "כתב תביעה", "paragraphs": paragraphs[:100]}],
        "appendices": [],
        "calculations": [],
        "legal_citations": [],
        "summary_total": calculations.get("total", 0),
        "verification_notes": ["הטקסט נוצר ללא עיבוד מלא — יש לבדוק ולערוך"],
        "stage_timing": {"total_seconds": 0, "stages_completed": 1},
    }


def _safe_parse_json(text):
    """Try multiple strategies to extract JSON from Claude's response."""
    if not text:
        return None

    # Strategy 1: Direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Strategy 2: Strip markdown fences
    if "```" in text:
        lines = text.split("\n")
        json_lines = []
        in_fence = False
        for line in lines:
            if line.strip().startswith("```") and not in_fence:
                in_fence = True
                continue
            elif line.strip() == "```" and in_fence:
                break
            elif in_fence:
                json_lines.append(line)
        stripped = "\n".join(json_lines)
        if stripped:
            try:
                return json.loads(stripped)
            except json.JSONDecodeError:
                pass

    # Strategy 3: Extract first { to last }
    first = text.find('{')
    last = text.rfind('}')
    if first != -1 and last > first:
        try:
            return json.loads(text[first:last + 1])
        except json.JSONDecodeError:
            pass

    return None
