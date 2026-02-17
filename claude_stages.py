"""
Two-stage AI generation pipeline for כתב תביעה documents.

Stage 1 (Analyst): Haiku — fast analysis → sections, laws, appendices
Stage 2 (Drafter): Haiku — full document generation with streaming

RESILIENT: If any stage fails, returns whatever was completed.
Logs FULL Claude responses for debugging.
"""

import json
import re
import time
import logging
import traceback

import anthropic

# ── Model ─────────────────────────────────────────────────────────────────────

MODEL = "claude-haiku-4-5-20251001"

# ── Timeouts ──────────────────────────────────────────────────────────────────

API_TIMEOUT = 120.0    # Anthropic client timeout per request
HARD_TIMEOUT = 150     # Pipeline hard timeout (under gunicorn's 180s)

# ── Stage 1: Analyst ─────────────────────────────────────────────────────────

STAGE1_SYSTEM = """You are a legal analyst for Israeli labor law.
Analyze the case facts. Return ONLY valid JSON:
{
  "sections_required": ["section headers in Hebrew, in order"],
  "applicable_laws": ["law names with relevant sections"],
  "appendices_detected": [{"number": 1, "description": "תלושי שכר", "reference_text": "מצורפים כנספח 1"}],
  "flags": {"harassment": false, "improper_hearing": false, "discriminatory_termination": false, "wage_theft": false, "corporate_veil": false}
}"""

STAGE1_MAX_TOKENS = 1200


def _run_stage1(client, raw_input, structured_data, selected_claims):
    """Stage 1: Analyze facts and determine document structure."""
    user_prompt = f"""נתוני התיק:
{json.dumps(structured_data, ensure_ascii=False, indent=2)}

רכיבי תביעה: {', '.join(selected_claims)}

עובדות:
{raw_input}

Analyze and return JSON."""

    logging.info("Stage 1: Calling Claude API...")
    message = client.messages.create(
        model=MODEL,
        max_tokens=STAGE1_MAX_TOKENS,
        system=STAGE1_SYSTEM,
        messages=[{"role": "user", "content": user_prompt}],
    )
    text = message.content[0].text.strip()
    logging.info(f"Stage 1 RAW RESPONSE ({len(text)} chars): {text[:2000]}")

    parsed = _safe_parse_json(text)
    if parsed is None:
        logging.error(f"Stage 1: JSON parse failed. Full response: {text}")
        raise ValueError(f"Stage 1 JSON parse failed, response starts with: {text[:200]}")
    return parsed


# ── Stage 2: Drafter (streaming) ─────────────────────────────────────────────

STAGE2_MAX_TOKENS = 4000

_STAGE2_BASE = """You are an Israeli labor law attorney drafting a כתב תביעה.
Write formal legal Hebrew, third person. Reference specific law sections.
Use ◄ prefix for appendix references. Show formulas: [salary] × [period] = [total] ₪.

RULES: Use EXACT amounts from input. Correct gender throughout. No invented facts. No extra claims.

Return ONLY valid JSON:
{
  "gender_form": "male"/"female",
  "sections": [{"header": "Hebrew header", "paragraphs": ["..."]}],
  "appendices": [{"number": 1, "description": "...", "reference_text": "..."}],
  "calculations": [{"component": "...", "formula": "...", "amount": 0}],
  "legal_citations": ["..."],
  "summary_total": 0
}"""


def _build_stage2_system(firm_patterns, legal_citations):
    """Build Stage 2 system prompt with prompt caching."""
    static_parts = [_STAGE2_BASE]

    if firm_patterns and firm_patterns.get("patterns"):
        static_parts.append(
            "\n\nFirm style:\n"
            + json.dumps(firm_patterns["patterns"], ensure_ascii=False)
        )

    if legal_citations:
        static_parts.append(
            "\n\nLegal refs:\n"
            + json.dumps(legal_citations, ensure_ascii=False)
        )

    return [
        {
            "type": "text",
            "text": "\n".join(static_parts),
            "cache_control": {"type": "ephemeral"},
        }
    ]


def _run_stage2_streaming(client, raw_input, structured_data, calculations,
                          stage1_analysis, firm_patterns, legal_citations,
                          on_progress=None):
    """Stage 2: Generate full document sections with streaming."""
    gender = structured_data.get("gender", "male")
    gender_label = "זכר" if gender == "male" else "נקבה"

    calc_summary = []
    for key, claim in calculations.get("claims", {}).items():
        entry = f"- {claim['name']}: {claim['amount']:,.0f} ₪"
        if claim.get("formula"):
            entry += f" ({claim['formula']})"
        calc_summary.append(entry)

    user_prompt = f"""נתוני התיק:
- שם: {structured_data.get('plaintiff_name', '')} | ת.ז.: {structured_data.get('plaintiff_id', '')} | מין: {gender_label}
- נתבע: {structured_data.get('defendant_name', '')} | ח.פ.: {structured_data.get('defendant_id', '')}
- תפקיד: {structured_data.get('job_title', '')}
- תקופה: {structured_data.get('start_date', '')} – {structured_data.get('end_date', '')} ({calculations.get('duration', {}).get('total_months', 0)} חודשים)
- סיום: {structured_data.get('termination_type', 'fired')}
- שכר: {structured_data.get('base_salary', '')} ₪ | עמלות: {structured_data.get('commissions', '0')} ₪
- שכר קובע: {calculations.get('determining_salary', 0):,.0f} ₪

חישובים (סכומים מחייבים):
{chr(10).join(calc_summary)}
סה"כ: {calculations.get('total', 0):,.0f} ₪

ניתוח:
{json.dumps(stage1_analysis, ensure_ascii=False)}

עובדות:
{raw_input}

Generate כתב תביעה. Use EXACT amounts above."""

    system = _build_stage2_system(firm_patterns, legal_citations)

    logging.info("Stage 2: Calling Claude API with streaming...")
    collected = []
    with client.messages.stream(
        model=MODEL,
        max_tokens=STAGE2_MAX_TOKENS,
        system=system,
        messages=[{"role": "user", "content": user_prompt}],
    ) as stream:
        for text in stream.text_stream:
            collected.append(text)
            if on_progress:
                on_progress(sum(len(c) for c in collected))

    full_text = "".join(collected).strip()
    logging.info(f"Stage 2 RAW RESPONSE ({len(full_text)} chars): {full_text[:3000]}")
    if len(full_text) > 3000:
        logging.info(f"Stage 2 RAW RESPONSE (tail): ...{full_text[-500:]}")

    parsed = _safe_parse_json(full_text)
    if parsed is None:
        logging.error(f"Stage 2: JSON parse failed! Full response logged above.")
        # Return the raw text as a fallback single-section document
        raise ValueError(f"Stage 2 JSON parse failed. Response length: {len(full_text)}, starts with: {full_text[:300]}")
    return parsed


# ── JSON Parsing Helpers ─────────────────────────────────────────────────────

def _safe_parse_json(text):
    """Try multiple strategies to parse JSON from Claude's response.

    Returns parsed dict or None if all strategies fail.
    """
    if not text:
        return None

    # Strategy 1: Direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Strategy 2: Strip markdown fences
    if "```" in text:
        stripped = _strip_markdown_fences(text)
        if stripped:
            try:
                return json.loads(stripped)
            except json.JSONDecodeError:
                pass

    # Strategy 3: Find JSON object with regex (first { to last })
    match = re.search(r'\{', text)
    if match:
        # Find the matching closing brace
        start = match.start()
        brace_count = 0
        end = start
        for i in range(start, len(text)):
            if text[i] == '{':
                brace_count += 1
            elif text[i] == '}':
                brace_count -= 1
                if brace_count == 0:
                    end = i + 1
                    break
        if end > start:
            try:
                return json.loads(text[start:end])
            except json.JSONDecodeError:
                pass

    # Strategy 4: Try from first { to last }
    first_brace = text.find('{')
    last_brace = text.rfind('}')
    if first_brace != -1 and last_brace > first_brace:
        try:
            return json.loads(text[first_brace:last_brace + 1])
        except json.JSONDecodeError:
            pass

    logging.error(f"All JSON parse strategies failed for text ({len(text)} chars)")
    return None


def _build_fallback_from_text(raw_text, structured_data, calculations):
    """Build a minimal valid response from raw text when JSON parsing fails."""
    # Split text into paragraphs and create a simple sections structure
    paragraphs = [p.strip() for p in raw_text.split('\n') if p.strip()]

    sections = []
    current_section = {"header": "כתב תביעה", "paragraphs": []}

    for para in paragraphs:
        # Detect headers (short lines, often bold or numbered)
        if len(para) < 80 and not para.endswith('.') and not para.endswith('₪'):
            if current_section["paragraphs"]:
                sections.append(current_section)
            current_section = {"header": para, "paragraphs": []}
        else:
            current_section["paragraphs"].append(para)

    if current_section["paragraphs"]:
        sections.append(current_section)

    if not sections:
        sections = [{"header": "כתב תביעה", "paragraphs": paragraphs[:50]}]

    return {
        "gender_form": structured_data.get("gender", "male"),
        "sections": sections,
        "appendices": [],
        "calculations": [],
        "legal_citations": [],
        "summary_total": calculations.get("total", 0),
    }


# ── Orchestrator ──────────────────────────────────────────────────────────────

def generate_claim_multistage(raw_input, structured_data, calculations,
                              firm_patterns=None, legal_citations=None,
                              api_key=None, on_stage=None):
    """Run the 2-stage AI generation pipeline.

    RESILIENT: If Stage 1 fails, skips to Stage 2 with empty analysis.
    If Stage 2 JSON fails, tries to extract usable text.
    Never returns None unless there's no API key or a total catastrophic failure.

    Returns:
        Combined dict with sections, appendices, calculations, citations,
        summary_total. Or None only on catastrophic failure.

    Raises:
        TimeoutError: If total elapsed time exceeds HARD_TIMEOUT.
    """
    if not api_key:
        logging.warning("generate_claim_multistage: no API key")
        return None

    client = anthropic.Anthropic(api_key=api_key, timeout=API_TIMEOUT)
    start_time = time.time()

    def _check_timeout(stage_name):
        elapsed = time.time() - start_time
        if elapsed > HARD_TIMEOUT:
            raise TimeoutError(
                f"חריגה ממגבלת הזמן ({HARD_TIMEOUT} שניות) בשלב {stage_name}. "
                "נסו שוב או השתמשו בתבנית."
            )

    # Collect selected claims
    claim_keys = {
        "claim_severance": "פיצויי פיטורים",
        "claim_prior_notice": "חלף הודעה מוקדמת",
        "claim_unpaid_salary": "שכר עבודה שלא שולם",
        "claim_overtime": "הפרשי שכר – שעות נוספות",
        "claim_pension": "הפרשי הפרשות לפנסיה",
        "claim_vacation": "הפרשי דמי חופשה ופדיון חופשה",
        "claim_holidays": "דמי חגים והפרשי דמי חג",
        "claim_recuperation": "דמי הבראה",
        "claim_salary_delay": "פיצויי הלנת שכר",
        "claim_emotional": "פיצוי בגין עוגמת נפש",
        "claim_deductions": "ניכויים שלא כדין",
        "claim_documents": "מסירת מסמכי גמר חשבון",
    }
    selected_claims = [name for key, name in claim_keys.items() if structured_data.get(key)]

    # ── Stage 1: Analyst (non-fatal) ──
    stage1 = {"sections_required": [], "applicable_laws": [], "appendices_detected": [], "flags": {}}

    if on_stage:
        on_stage("analyzing", "מנתח עובדות...")
    logging.info("Stage 1 (Analyst) starting...")
    try:
        stage1 = _run_stage1(client, raw_input, structured_data, selected_claims)
        elapsed = time.time() - start_time
        logging.info(f"Stage 1 completed in {elapsed:.1f}s — sections: {stage1.get('sections_required', [])}")
    except Exception as e:
        elapsed = time.time() - start_time
        logging.error(f"Stage 1 FAILED after {elapsed:.1f}s (non-fatal, continuing with empty analysis): {e}")
        logging.error(f"Stage 1 traceback: {traceback.format_exc()}")
        # Continue with empty stage1 — Stage 2 can still work

    _check_timeout("ניתוח")

    # ── Stage 2: Drafter (streaming) ──
    if on_stage:
        on_stage("drafting", "מנסח סעיפים...")
    logging.info("Stage 2 (Drafter+streaming) starting...")

    def _on_stream_progress(chars):
        if on_stage and chars % 500 < 50:
            on_stage("drafting_progress", f"מנסח... ({chars} תווים)")

    stage2 = None
    stage2_raw_text = None

    try:
        stage2 = _run_stage2_streaming(
            client, raw_input, structured_data, calculations,
            stage1, firm_patterns, legal_citations,
            on_progress=_on_stream_progress,
        )
        elapsed = time.time() - start_time
        logging.info(f"Stage 2 completed in {elapsed:.1f}s: {len(stage2.get('sections', []))} sections")
    except ValueError as e:
        # JSON parse failed but we might have raw text
        elapsed = time.time() - start_time
        logging.error(f"Stage 2 JSON PARSE FAILED after {elapsed:.1f}s: {e}")
        logging.error(f"Stage 2 traceback: {traceback.format_exc()}")
        # Try to build fallback from whatever text we got
        error_msg = str(e)
        # The raw text was already logged in _run_stage2_streaming
    except Exception as e:
        elapsed = time.time() - start_time
        logging.error(f"Stage 2 FAILED after {elapsed:.1f}s: {e}")
        logging.error(f"Stage 2 traceback: {traceback.format_exc()}")

    # If Stage 2 failed completely, try a simpler single-call fallback
    if stage2 is None:
        logging.warning("Stage 2 produced no result. Attempting simple fallback call...")
        if on_stage:
            on_stage("drafting", "ניסיון נוסף...")
        try:
            stage2 = _run_simple_fallback(client, raw_input, structured_data, calculations, selected_claims)
            elapsed = time.time() - start_time
            logging.info(f"Fallback completed in {elapsed:.1f}s")
        except Exception as e2:
            logging.error(f"Fallback ALSO FAILED: {e2}")
            logging.error(f"Fallback traceback: {traceback.format_exc()}")
            return None

    total_elapsed = time.time() - start_time
    logging.info(f"Pipeline completed in {total_elapsed:.1f}s total")

    if on_stage:
        on_stage("done", f"הושלם ב-{total_elapsed:.0f} שניות")

    return {
        "gender_form": stage2.get("gender_form", structured_data.get("gender", "male")),
        "sections": stage2.get("sections", []),
        "appendices": stage2.get("appendices", []) + stage1.get("appendices_detected", []),
        "calculations": stage2.get("calculations", []),
        "legal_citations": stage2.get("legal_citations", []),
        "summary_total": calculations.get("total", 0),
        "verification_notes": [],
        "stage_timing": {
            "total_seconds": round(total_elapsed, 1),
            "stages_completed": 2,
        },
    }


def _run_simple_fallback(client, raw_input, structured_data, calculations, selected_claims):
    """Ultra-simple single-call fallback if Stage 2 fails.

    Uses a much shorter prompt and lower max_tokens to maximize chance of success.
    """
    gender = structured_data.get("gender", "male")
    gender_he = "התובעת" if gender == "female" else "התובע"

    calc_lines = []
    for key, claim in calculations.get("claims", {}).items():
        calc_lines.append(f"{claim['name']}: {claim['amount']:,.0f} ₪")

    prompt = f"""Write a כתב תביעה in Hebrew for {gender_he}.
Claims: {', '.join(selected_claims)}
Amounts: {'; '.join(calc_lines)}
Total: {calculations.get('total', 0):,.0f} ₪
Facts: {raw_input[:2000]}

Return JSON: {{"gender_form":"{gender}","sections":[{{"header":"...","paragraphs":["..."]}}],"appendices":[],"calculations":[],"legal_citations":[],"summary_total":{calculations.get('total', 0)}}}"""

    logging.info("Fallback: Calling Claude API (simple, max_tokens=3000)...")
    message = client.messages.create(
        model=MODEL,
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}],
    )
    text = message.content[0].text.strip()
    logging.info(f"Fallback RAW RESPONSE ({len(text)} chars): {text[:2000]}")

    parsed = _safe_parse_json(text)
    if parsed is not None:
        return parsed

    # Last resort: wrap the raw text as a single section
    logging.warning("Fallback JSON also failed — wrapping raw text as single section")
    return _build_fallback_from_text(text, structured_data, calculations)


def _strip_markdown_fences(text):
    """Strip markdown code fences from response."""
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
    return "\n".join(json_lines)
