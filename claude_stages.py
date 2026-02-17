"""
Two-stage AI generation pipeline for כתב תביעה documents.

Stage 1 (Analyst): Haiku — fast analysis → sections, laws, appendices
Stage 2 (Drafter): Haiku — full document generation with streaming

All calls use Haiku for speed (under Render's 60s limit).
API timeout: 55s. Hard timeout: 55s.
"""

import json
import time
import logging

import anthropic

# ── Model ─────────────────────────────────────────────────────────────────────

MODEL = "claude-haiku-4-5-20251001"

# ── Timeouts ──────────────────────────────────────────────────────────────────

API_TIMEOUT = 55.0     # Anthropic client timeout (under Render's 60s)
HARD_TIMEOUT = 55      # Pipeline hard timeout in seconds

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

    message = client.messages.create(
        model=MODEL,
        max_tokens=STAGE1_MAX_TOKENS,
        system=STAGE1_SYSTEM,
        messages=[{"role": "user", "content": user_prompt}],
    )
    text = message.content[0].text.strip()
    if text.startswith("```"):
        text = _strip_markdown_fences(text)
    return json.loads(text)


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
    if full_text.startswith("```"):
        full_text = _strip_markdown_fences(full_text)
    return json.loads(full_text)


# ── Orchestrator ──────────────────────────────────────────────────────────────

def generate_claim_multistage(raw_input, structured_data, calculations,
                              firm_patterns=None, legal_citations=None,
                              api_key=None, on_stage=None):
    """Run the 2-stage AI generation pipeline.

    Returns:
        Combined dict with sections, appendices, calculations, citations,
        summary_total. Or None on failure.

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

    # ── Stage 1: Analyst ──
    if on_stage:
        on_stage("analyzing", "מנתח עובדות...")
    logging.info("Stage 1 (Analyst) starting...")
    try:
        stage1 = _run_stage1(client, raw_input, structured_data, selected_claims)
        elapsed = time.time() - start_time
        logging.info(f"Stage 1 completed in {elapsed:.1f}s")
    except Exception as e:
        logging.error(f"Stage 1 failed: {e}")
        return None

    _check_timeout("ניתוח")

    # ── Stage 2: Drafter (streaming) ──
    if on_stage:
        on_stage("drafting", "מנסח סעיפים...")
    logging.info("Stage 2 (Drafter+streaming) starting...")

    def _on_stream_progress(chars):
        if on_stage and chars % 500 < 50:
            on_stage("drafting_progress", f"מנסח... ({chars} תווים)")

    try:
        stage2 = _run_stage2_streaming(
            client, raw_input, structured_data, calculations,
            stage1, firm_patterns, legal_citations,
            on_progress=_on_stream_progress,
        )
        elapsed = time.time() - start_time
        logging.info(f"Stage 2 completed in {elapsed:.1f}s: {len(stage2.get('sections', []))} sections")
    except json.JSONDecodeError as e:
        logging.error(f"Stage 2 JSON parse failed: {e}")
        return None
    except Exception as e:
        logging.error(f"Stage 2 failed: {e}")
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
