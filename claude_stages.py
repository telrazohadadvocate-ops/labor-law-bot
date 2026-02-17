"""
Single-call AI generation for כתב תביעה documents.

Returns PLAIN HEBREW TEXT with === TITLE === section delimiters.
No JSON anywhere in the pipeline.
"""

import re
import logging
import traceback

import anthropic

MODEL = "claude-haiku-4-5-20251001"
MAX_TOKENS = 8192
API_TIMEOUT = 60.0


# ── Gender replacement maps ──────────────────────────────────────────────────

GENDER_MALE = {
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
    "שלו/ה": "שלו",
    "הוא/היא": "הוא",
}

GENDER_FEMALE = {
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
    "שלו/ה": "שלה",
    "הוא/היא": "היא",
}


def fix_gender(text, gender):
    """Replace gender-neutral slashed forms with the correct gender form."""
    replacements = GENDER_MALE if gender == "male" else GENDER_FEMALE
    for pattern, replacement in replacements.items():
        text = text.replace(pattern, replacement)
    return text


def parse_plain_text_sections(raw_text):
    """Parse plain text with === TITLE === delimiters into sections.

    Returns:
        List of dicts: [{"title": "...", "lines": ["line1", "line2", ...]}, ...]
    """
    sections = []
    current_title = ""
    current_lines = []

    for line in raw_text.split("\n"):
        stripped = line.strip()

        # Check for section delimiter: === TITLE ===
        match = re.match(r'^===\s*(.+?)\s*===$', stripped)
        if match:
            # Save previous section
            if current_title or current_lines:
                sections.append({
                    "title": current_title,
                    "lines": [l for l in current_lines if l.strip()],
                })
            current_title = match.group(1)
            current_lines = []
        else:
            current_lines.append(stripped)

    # Save last section
    if current_title or current_lines:
        sections.append({
            "title": current_title,
            "lines": [l for l in current_lines if l.strip()],
        })

    return sections


def generate_claim_single(raw_input, structured_data, calculations,
                          firm_patterns=None, legal_citations=None,
                          api_key=None):
    """Single Claude API call to generate כתב תביעה as PLAIN TEXT.

    Returns:
        Dict with 'plain_text' (the raw AI text) and 'sections' (parsed).
        Or None on total failure.
    """
    if not api_key:
        logging.warning("generate_claim_single: no API key")
        return None

    client = anthropic.Anthropic(api_key=api_key, timeout=API_TIMEOUT)

    gender = structured_data.get("gender", "male")
    gender_label = "זכר" if gender == "male" else "נקבה"
    pronoun = "התובע" if gender == "male" else "התובעת"

    # Build calculation lines
    calc_lines = []
    for key, claim in calculations.get("claims", {}).items():
        line = f"- {claim['name']}: {claim['amount']:,.0f} ₪"
        if claim.get("formula"):
            line += f" ({claim['formula']})"
        calc_lines.append(line)

    # Termination type in Hebrew
    termination_type = structured_data.get("termination_type", "fired")
    if termination_type == "fired":
        termination_he = "פוטר" if gender == "male" else "פוטרה"
    elif termination_type == "resigned_justified":
        termination_he = "התפטר בדין מפוטר" if gender == "male" else "התפטרה בדין מפוטרת"
    else:
        termination_he = "התפטר" if gender == "male" else "התפטרה"

    # Build claim components list for the prompt
    claim_components = []
    for key, claim in calculations.get("claims", {}).items():
        claim_components.append(claim['name'])

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

רכיבי התביעה שנבחרו:
{chr(10).join(claim_components)}

חישובים (סכומים מחייבים — השתמש בדיוק בסכומים אלה):
{chr(10).join(calc_lines)}
סה"כ: {calculations.get('total', 0):,.0f} ₪

עובדות גולמיות (חובה לשלב בסעיף רקע עובדתי):
{raw_input}

כתוב כתב תביעה מלא בעברית. החזר טקסט רגיל בלבד."""

    # System prompt
    system = _build_system_prompt(gender, firm_patterns)

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

    # Strip any accidental code blocks or JSON wrapping
    raw_text = _strip_code_blocks(raw_text)

    # Apply gender fixes on the entire text
    raw_text = fix_gender(raw_text, gender)

    # Parse into sections
    sections = parse_plain_text_sections(raw_text)
    logging.info(f"Parsed {len(sections)} sections from plain text")
    for i, s in enumerate(sections):
        logging.info(f"  [{i}] '{s['title']}' — {len(s['lines'])} lines")

    return {
        "plain_text": raw_text,
        "sections": sections,
        "gender": gender,
    }


def _build_system_prompt(gender, firm_patterns=None):
    """Build the system prompt for plain text generation."""
    if gender == "male":
        gender_instruction = "השתמש בגוף שלישי זכר: התובע, הועסק, פוטר, זכאי, עובד, טוען, יבקש, שכרו, עבודתו, זכויותיו"
    else:
        gender_instruction = "השתמשי בגוף שלישי נקבה: התובעת, הועסקה, פוטרה, זכאית, עובדת, טוענת, תבקש, שכרה, עבודתה, זכויותיה"

    system = f"""אתה עורך דין ישראלי לדיני עבודה. כתוב כתב תביעה מלא בעברית משפטית רשמית.

החזר טקסט רגיל בלבד. ללא JSON, ללא קוד, ללא markup. רק טקסט עברי רגיל.

השתמש בפורמט הבא בדיוק:

=== כללי ===
הצהרות פרוצדורליות כלליות

=== הצדדים ===
פרטי התובע/ת והנתבעת

=== רקע עובדתי ===
העובדות כפי שנמסרו

=== היקף משרה ושכר קובע ===
פרטי השכר והמשרה

=== רכיבי התביעה ===
סעיפים נפרדים לכל רכיב תביעה עם חישוב

=== סיכום ===
סיכום וסעד מבוקש

כללים:
- כתוב הכל בעברית בלבד. אסור אנגלית בשום מקום.
- {gender_instruction}
- השתמש בסכומים המדויקים מהנתונים
- אל תמציא עובדות שלא סופקו
- שלב את העובדות הגולמיות בתוך סעיף רקע עובדתי
- מספר את הפסקאות ברצף מתמשך (1, 2, 3...) לאורך כל המסמך
- כל סעיף חייב להכיל 3-5 פסקאות ממוספרות לפחות
- ציין חוקים בשמם המלא כולל שנת חקיקה, לדוגמה: חוק פיצויי פיטורים, התשכ"ג-1963
- בסוף כל רכיב תביעה, הוסף שורת חישוב בפורמט: שכר קובע × מספר חודשים = סכום ₪
- אחרי סעיפים שיש להם מסמך תומך, הוסף שורה: ◄ ראה נספח [מספר] — [תיאור]
- הפרד בין פסקאות בשורה ריקה
- השתמש ב-=== כותרת === כמפריד סעיפים
- הוסף סעיפים נוספים לפי הצורך בהתאם לעובדות התיק (למשל: שימוע ופיטורים, התעמרות, שעות נוספות)"""

    # Append firm style hints if available
    if firm_patterns and firm_patterns.get("patterns"):
        import json
        style_keys = firm_patterns["patterns"]
        if style_keys.get("opening_phrases"):
            system += f"\n\nדוגמאות פתיחה של המשרד: {json.dumps(style_keys['opening_phrases'][:3], ensure_ascii=False)}"
        if style_keys.get("closing_phrases"):
            system += f"\nדוגמאות סיום: {json.dumps(style_keys['closing_phrases'][:2], ensure_ascii=False)}"

    return system


def _strip_code_blocks(text):
    """Strip markdown code blocks if Claude accidentally wraps the response."""
    if "```" in text:
        lines = text.split("\n")
        clean_lines = []
        in_fence = False
        for line in lines:
            if line.strip().startswith("```"):
                in_fence = not in_fence
                continue
            clean_lines.append(line)
        return "\n".join(clean_lines)
    return text
