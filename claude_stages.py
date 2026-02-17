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

SYSTEM_PROMPT = """You are an Israeli labor law attorney. Given case data and facts, draft a כתב תביעה.
Write formal legal Hebrew, third person. Reference laws where relevant.
Use ◄ prefix for appendix references. Show calculation formulas with ₪.
Use EXACT amounts provided. Correct gender. No invented facts. No extra claims.

Return ONLY valid JSON:
{"gender_form":"male or female","sections":[{"title":"section title in Hebrew","content":"section body text"}],"appendices":["appendix description"]}"""


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

    calc_lines = []
    for key, claim in calculations.get("claims", {}).items():
        line = f"- {claim['name']}: {claim['amount']:,.0f} ₪"
        if claim.get("formula"):
            line += f" ({claim['formula']})"
        calc_lines.append(line)

    user_prompt = f"""נתוני התיק:
שם: {structured_data.get('plaintiff_name', '')} | ת.ז.: {structured_data.get('plaintiff_id', '')} | מין: {gender_label}
נתבע: {structured_data.get('defendant_name', '')} | ח.פ.: {structured_data.get('defendant_id', '')}
תפקיד: {structured_data.get('job_title', '')}
תקופה: {structured_data.get('start_date', '')} – {structured_data.get('end_date', '')}
סיום: {structured_data.get('termination_type', 'fired')}
שכר: {structured_data.get('base_salary', '')} ₪ | עמלות: {structured_data.get('commissions', '0')} ₪
שכר קובע: {calculations.get('determining_salary', 0):,.0f} ₪

חישובים (סכומים מחייבים):
{chr(10).join(calc_lines)}
סה"כ: {calculations.get('total', 0):,.0f} ₪

עובדות:
{raw_input}

Generate כתב תביעה as JSON."""

    # Build system with optional cached firm patterns
    system = SYSTEM_PROMPT
    if firm_patterns and firm_patterns.get("patterns"):
        style_keys = firm_patterns["patterns"]
        if style_keys.get("opening_phrases"):
            system += f"\n\nFirm opening style examples: {json.dumps(style_keys['opening_phrases'][:3], ensure_ascii=False)}"
        if style_keys.get("closing_phrases"):
            system += f"\nFirm closing style: {json.dumps(style_keys['closing_phrases'][:2], ensure_ascii=False)}"

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
    """Normalize the parsed response to a consistent format."""
    sections = parsed.get("sections", [])

    # Handle both {"title","content"} and {"header","paragraphs"} formats
    normalized_sections = []
    for s in sections:
        if "header" in s and "paragraphs" in s:
            normalized_sections.append(s)
        elif "title" in s and "content" in s:
            content = s["content"]
            if isinstance(content, str):
                paragraphs = [p.strip() for p in content.split("\n") if p.strip()]
            else:
                paragraphs = content if isinstance(content, list) else [str(content)]
            normalized_sections.append({"header": s["title"], "paragraphs": paragraphs})
        else:
            # Unknown format, try to use whatever keys are there
            header = s.get("title") or s.get("header") or s.get("name") or "סעיף"
            body = s.get("content") or s.get("paragraphs") or s.get("text") or ""
            if isinstance(body, str):
                paragraphs = [p.strip() for p in body.split("\n") if p.strip()]
            elif isinstance(body, list):
                paragraphs = body
            else:
                paragraphs = [str(body)]
            normalized_sections.append({"header": header, "paragraphs": paragraphs})

    appendices_raw = parsed.get("appendices", [])
    appendices = []
    for i, a in enumerate(appendices_raw):
        if isinstance(a, str):
            appendices.append({"number": i + 1, "description": a, "reference_text": a})
        elif isinstance(a, dict):
            appendices.append(a)

    return {
        "gender_form": parsed.get("gender_form", structured_data.get("gender", "male")),
        "sections": normalized_sections,
        "appendices": appendices,
        "calculations": parsed.get("calculations", []),
        "legal_citations": parsed.get("legal_citations", []),
        "summary_total": calculations.get("total", 0),
        "verification_notes": [],
        "stage_timing": {"total_seconds": 0, "stages_completed": 1},
    }


def _wrap_raw_text(raw_text, structured_data, calculations):
    """Wrap raw text as a single section when JSON parsing fails."""
    paragraphs = [p.strip() for p in raw_text.split("\n") if p.strip()]
    return {
        "gender_form": structured_data.get("gender", "male"),
        "sections": [{"header": "כתב תביעה", "paragraphs": paragraphs[:100]}],
        "appendices": [],
        "calculations": [],
        "legal_citations": [],
        "summary_total": calculations.get("total", 0),
        "verification_notes": ["הטקסט נוצר ללא עיבוד JSON — יש לבדוק ולערוך"],
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
