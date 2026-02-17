"""
Multi-stage AI generation pipeline for כתב תביעה documents.

Stage 1 (Analyst): Analyzes raw facts → identifies sections, laws, appendices
Stage 2 (Drafter): Generates full document sections using firm patterns + citations
Stage 3 (Verifier): Validates calculations and section completeness

Each stage has independent error handling. Stage 3 is non-fatal.
"""

import json
import time
import logging

import anthropic

# ── Stage 1: Analyst ──────────────────────────────────────────────────────────

STAGE1_SYSTEM = """You are a legal analyst specializing in Israeli labor law.
Analyze the provided case facts and structured data. Identify:
1. Which document sections are needed (beyond the standard כללי/הצדדים/רקע/סיכום)
2. Which specific laws and regulations apply
3. Whether appendices are referenced or implied in the facts
4. Special flags: harassment, improper hearing, discriminatory termination, etc.

Return ONLY valid JSON with this structure:
{
  "sections_required": ["list of section headers in Hebrew, in order"],
  "applicable_laws": ["list of law names with sections that apply"],
  "appendices_detected": [
    {"number": 1, "description": "תלושי שכר", "reference_text": "תלושי שכר מצורפים כנספח 1"}
  ],
  "flags": {
    "harassment": false,
    "improper_hearing": false,
    "discriminatory_termination": false,
    "wage_theft": false,
    "corporate_veil": false
  }
}"""

STAGE1_MAX_TOKENS = 1500


def _run_stage1(client, raw_input, structured_data, selected_claims):
    """Stage 1: Analyze facts and determine document structure."""
    user_prompt = f"""נתוני התיק המובנים:
{json.dumps(structured_data, ensure_ascii=False, indent=2)}

רכיבי תביעה שנבחרו: {', '.join(selected_claims)}

עובדות גולמיות:
{raw_input}

Analyze the above case and return the JSON analysis."""

    message = client.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=STAGE1_MAX_TOKENS,
        system=STAGE1_SYSTEM,
        messages=[{"role": "user", "content": user_prompt}],
    )
    text = message.content[0].text.strip()
    if text.startswith("```"):
        text = _strip_markdown_fences(text)
    return json.loads(text)


# ── Stage 2: Drafter ──────────────────────────────────────────────────────────

STAGE2_MAX_TOKENS = 6000


def _build_stage2_system(firm_patterns, legal_citations):
    """Build Stage 2 system prompt with firm patterns and legal citations.

    Uses cache_control for prompt caching on the large reference blocks.
    """
    base = """You are an Israeli labor law attorney drafting a complete כתב תביעה for בית הדין לעבודה.

Write each section with formal legal Hebrew, third person, proper clause structure.
Reference specific law sections where applicable.
Use the firm's writing style patterns provided below.
Reference appendices with ◄ prefix where relevant.
Include calculation formulas using the provided amounts (use × and = symbols with ₪).

CRITICAL RULES:
- Do NOT invent facts not in the input
- Do NOT add claims not in the selected list
- Use EXACT calculated amounts from the structured data
- Use correct gender throughout (based on gender_form)
- For each claim component, show the formula: [salary] × [period] = [total]
- Return ONLY valid JSON, no markdown fences

Return JSON with this structure:
{
  "gender_form": "male" or "female",
  "sections": [
    {"header": "section header in Hebrew", "paragraphs": ["paragraph 1", "paragraph 2"]}
  ],
  "appendices": [
    {"number": 1, "description": "תלושי שכר", "reference_text": "מצורפים כנספח 1"}
  ],
  "calculations": [
    {"component": "פיצויי פיטורים", "formula": "10,000 ₪ × 2.5 שנים = 25,000 ₪", "amount": 25000}
  ],
  "legal_citations": ["חוק פיצויי פיטורים, תשכ\\"ג-1963"],
  "summary_total": 123456
}"""

    system_parts = [{"type": "text", "text": base}]

    if firm_patterns and firm_patterns.get("patterns"):
        patterns_text = "\n\n## Firm Writing Patterns (FOLLOW THIS STYLE):\n" + json.dumps(
            firm_patterns["patterns"], ensure_ascii=False, indent=2
        )
        system_parts.append({
            "type": "text",
            "text": patterns_text,
            "cache_control": {"type": "ephemeral"},
        })

    if legal_citations:
        citations_text = "\n\n## Legal Citations Reference:\n" + json.dumps(
            legal_citations, ensure_ascii=False, indent=2
        )
        system_parts.append({
            "type": "text",
            "text": citations_text,
            "cache_control": {"type": "ephemeral"},
        })

    return system_parts


def _run_stage2(client, raw_input, structured_data, calculations, stage1_analysis,
                firm_patterns, legal_citations):
    """Stage 2: Generate full document sections."""
    gender = structured_data.get("gender", "male")
    gender_label = "זכר" if gender == "male" else "נקבה"

    # Build calculation summary for the prompt
    calc_summary = []
    for key, claim in calculations.get("claims", {}).items():
        entry = f"- {claim['name']}: {claim['amount']:,.0f} ₪"
        if claim.get("formula"):
            entry += f" (נוסחה: {claim['formula']})"
        calc_summary.append(entry)

    user_prompt = f"""נתוני התיק:
- שם התובע/ת: {structured_data.get('plaintiff_name', '')}
- ת.ז.: {structured_data.get('plaintiff_id', '')}
- מין: {gender_label}
- שם הנתבע/ת: {structured_data.get('defendant_name', '')}
- ח.פ./ע.מ.: {structured_data.get('defendant_id', '')}
- תפקיד: {structured_data.get('job_title', '')}
- תאריך תחילת עבודה: {structured_data.get('start_date', '')}
- תאריך סיום עבודה: {structured_data.get('end_date', '')}
- סוג סיום העסקה: {structured_data.get('termination_type', 'fired')}
- שכר בסיס: {structured_data.get('base_salary', '')} ₪
- עמלות/תוספות: {structured_data.get('commissions', '0')} ₪
- ימי עבודה בשבוע: {structured_data.get('work_days_per_week', '6')}
- תקופת העסקה: {calculations.get('duration', {}).get('total_months', 0)} חודשים ({calculations.get('duration', {}).get('decimal_years', 0)} שנים)
- שכר קובע: {calculations.get('determining_salary', 0):,.0f} ₪

תוצאות חישוב (סכומים מחייבים – השתמש בדיוק בסכומים אלה):
{chr(10).join(calc_summary)}
סה"כ: {calculations.get('total', 0):,.0f} ₪

ניתוח שלב 1:
{json.dumps(stage1_analysis, ensure_ascii=False, indent=2)}

עובדות גולמיות ותיאור הנסיבות:
{raw_input}

Generate the full כתב תביעה sections. Use the EXACT amounts from the calculation results above."""

    system = _build_stage2_system(firm_patterns, legal_citations)

    message = client.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=STAGE2_MAX_TOKENS,
        system=system,
        messages=[{"role": "user", "content": user_prompt}],
    )
    text = message.content[0].text.strip()
    if text.startswith("```"):
        text = _strip_markdown_fences(text)
    return json.loads(text)


# ── Stage 3: Verifier ────────────────────────────────────────────────────────

STAGE3_SYSTEM = """You are a quality reviewer for Israeli labor court claim documents.
Review the drafted sections against the authoritative calculations.

Check:
1. All selected claims have a corresponding section
2. Amounts in the text match the authoritative calculated amounts exactly
3. No claims were added that weren't selected
4. Gender consistency throughout

If corrections are needed, return the corrected sections.
If no corrections needed, return the sections unchanged.

Return ONLY valid JSON:
{
  "verified_sections": [{"header": "...", "paragraphs": ["..."]}],
  "verification_notes": ["list of changes made or 'אושר ללא שינויים'"],
  "amounts_verified": true
}"""

STAGE3_MAX_TOKENS = 800


def _run_stage3(client, stage2_output, calculations, selected_claims):
    """Stage 3: Verify and correct the drafted sections."""
    calc_amounts = {}
    for key, claim in calculations.get("claims", {}).items():
        calc_amounts[claim["name"]] = claim["amount"]

    user_prompt = f"""סעיפים שנוצרו (שלב 2):
{json.dumps(stage2_output.get('sections', []), ensure_ascii=False, indent=2)}

סכומים מחייבים:
{json.dumps(calc_amounts, ensure_ascii=False, indent=2)}

סה"כ מחייב: {calculations.get('total', 0):,.0f} ₪

רכיבי תביעה שנבחרו: {', '.join(selected_claims)}

Verify the sections and return corrections if needed."""

    message = client.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=STAGE3_MAX_TOKENS,
        system=STAGE3_SYSTEM,
        messages=[{"role": "user", "content": user_prompt}],
    )
    text = message.content[0].text.strip()
    if text.startswith("```"):
        text = _strip_markdown_fences(text)
    return json.loads(text)


# ── Orchestrator ──────────────────────────────────────────────────────────────

# Total budget: 120s (Render timeout) minus some safety margin
TOTAL_TIMEOUT = 110  # seconds
STAGE3_MIN_REMAINING = 20  # skip Stage 3 if less than this many seconds remain


def generate_claim_multistage(raw_input, structured_data, calculations,
                              firm_patterns=None, legal_citations=None,
                              api_key=None):
    """Run the 3-stage AI generation pipeline.

    Args:
        raw_input: Raw facts text from user.
        structured_data: Dict with all form data.
        calculations: Results from calculate_all_claims().
        firm_patterns: Loaded firm_patterns.json dict (or None).
        legal_citations: Loaded legal_citations.json dict (or None).
        api_key: Anthropic API key.

    Returns:
        Combined dict with sections, appendices, calculations, citations,
        summary_total, verification_notes. Or None on failure.
    """
    if not api_key:
        logging.warning("generate_claim_multistage: no API key")
        return None

    client = anthropic.Anthropic(api_key=api_key, timeout=90.0)
    start_time = time.time()

    # Collect selected claims for prompts
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
    logging.info("Stage 1 (Analyst) starting...")
    try:
        stage1 = _run_stage1(client, raw_input, structured_data, selected_claims)
        elapsed = time.time() - start_time
        logging.info(f"Stage 1 completed in {elapsed:.1f}s: {len(stage1.get('sections_required', []))} sections identified")
    except Exception as e:
        logging.error(f"Stage 1 failed: {e}")
        return None

    # ── Stage 2: Drafter ──
    logging.info("Stage 2 (Drafter) starting...")
    try:
        stage2 = _run_stage2(client, raw_input, structured_data, calculations,
                             stage1, firm_patterns, legal_citations)
        elapsed = time.time() - start_time
        logging.info(f"Stage 2 completed in {elapsed:.1f}s: {len(stage2.get('sections', []))} sections generated")
    except Exception as e:
        logging.error(f"Stage 2 failed: {e}")
        return None

    # ── Stage 3: Verifier (non-fatal, skipped if time is short) ──
    remaining = TOTAL_TIMEOUT - (time.time() - start_time)
    verification_notes = []
    final_sections = stage2.get("sections", [])

    if remaining >= STAGE3_MIN_REMAINING:
        logging.info(f"Stage 3 (Verifier) starting... ({remaining:.0f}s remaining)")
        try:
            stage3 = _run_stage3(client, stage2, calculations, selected_claims)
            elapsed = time.time() - start_time
            logging.info(f"Stage 3 completed in {elapsed:.1f}s")
            if stage3.get("verified_sections"):
                final_sections = stage3["verified_sections"]
            verification_notes = stage3.get("verification_notes", [])
        except Exception as e:
            logging.warning(f"Stage 3 failed (non-fatal, using Stage 2 output): {e}")
            verification_notes = [f"אימות לא הושלם: {str(e)[:100]}"]
    else:
        logging.info(f"Skipping Stage 3 — only {remaining:.0f}s remaining")
        verification_notes = ["שלב האימות דולג עקב מגבלת זמן"]

    total_elapsed = time.time() - start_time
    logging.info(f"Multi-stage pipeline completed in {total_elapsed:.1f}s total")

    # ── Combine results ──
    return {
        "gender_form": stage2.get("gender_form", structured_data.get("gender", "male")),
        "sections": final_sections,
        "appendices": stage2.get("appendices", []) + _appendices_from_stage1(stage1),
        "calculations": stage2.get("calculations", []),
        "legal_citations": stage2.get("legal_citations", []),
        "summary_total": calculations.get("total", 0),
        "verification_notes": verification_notes,
        "stage_timing": {
            "total_seconds": round(total_elapsed, 1),
            "stages_completed": 3 if remaining >= STAGE3_MIN_REMAINING else 2,
        },
    }


def _appendices_from_stage1(stage1):
    """Extract any appendices detected by Stage 1 that aren't already in Stage 2."""
    return stage1.get("appendices_detected", [])


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
