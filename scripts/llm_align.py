# llm_align.py
# Uses a local LLM (via OpenAI-compatible API) to map course Learning Outcomes
# to Program Outcomes, then writes alignment.json for use by generate_report.py
#
# Copy .env.example to .env and fill in your LLM details before running.
# Run: python scripts/llm_align.py

import argparse
import json
import os
import sys
from datetime import datetime
from openai import OpenAI
from dotenv import load_dotenv

# Force UTF-8 output on Windows consoles
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── Configuration (loaded from .env) ─────────────────────────────────────────
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))

LLM_BASE_URL = os.getenv("LLM_BASE_URL")
LLM_MODEL    = os.getenv("LLM_MODEL",   "gpt-4o")
LLM_API_KEY  = os.getenv("LLM_API_KEY", "none")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ── Cloud-leakage guardrail ───────────────────────────────────────────────────
def check_cloud_safety(allow_cloud: bool):
    """Warn and confirm before sending data to a cloud API."""
    if LLM_BASE_URL:
        return   # local endpoint configured — safe to proceed
    if allow_cloud:
        print("WARNING: LLM_BASE_URL not set. Sending data to OpenAI cloud API (--allow-cloud flag set).")
        return
    print("\n" + "!" * 60)
    print("  WARNING: LLM_BASE_URL is not set in your .env file.")
    print("  Curriculum data will be sent to the OpenAI cloud API.")
    print("  This may violate your institution's data handling policy.")
    print("!" * 60)
    print("\n  To use a local model, set LLM_BASE_URL in your .env file.")
    print("  See .env.example for configuration options.")
    print("\n  To proceed with the cloud API anyway, re-run with --allow-cloud")
    sys.exit(1)

# ── LLM client ────────────────────────────────────────────────────────────────
client_kwargs = {"api_key": LLM_API_KEY}
if LLM_BASE_URL:
    client_kwargs["base_url"] = LLM_BASE_URL
client = OpenAI(**client_kwargs)


def align_outcome(course_code, course_name, outcome_text, objectives, po_list_text):
    """Ask the LLM to map a single course LO to Program Outcomes."""
    objectives_text = "\n".join(f"  - {o}" for o in objectives if o)

    prompt = f"""You are a curriculum alignment expert for a college diploma program.

A course learning outcome and its supporting objectives are listed below.
Map this outcome to the Program Outcomes it best supports.

Classify each relevant Program Outcome as:
  - "primary"   : the LO directly and substantially addresses this PO
  - "supporting": the LO partially or indirectly supports this PO
  - omit entirely if there is no meaningful connection

Course: {course_code} — {course_name}
Learning Outcome: {outcome_text}
Objectives:
{objectives_text}

Program Outcomes to consider:
{po_list_text}

Respond ONLY with a valid JSON object in this exact format:
{{
  "primary": ["PO1", "PO3"],
  "supporting": ["PO5", "PO6"]
}}
Use only the PO IDs listed above. Both arrays may be empty but must be present."""

    response = client.chat.completions.create(
        model=LLM_MODEL,
        messages=[
            {"role": "system", "content": "You are a curriculum alignment expert. Respond only with valid JSON."},
            {"role": "user",   "content": prompt},
        ],
        temperature=0.1,
        max_tokens=300,
    )

    raw = response.choices[0].message.content.strip()
    # Strip markdown code fences if the model adds them
    raw = raw.replace("```json", "").replace("```", "").strip()

    try:
        result = json.loads(raw)
        return {
            "primary":    result.get("primary", []),
            "supporting": result.get("supporting", []),
        }
    except json.JSONDecodeError:
        print(f"  WARNING: Could not parse JSON for {course_code} — '{outcome_text[:60]}...'")
        print(f"  Raw response: {raw}")
        return {"primary": [], "supporting": []}


def main():
    parser = argparse.ArgumentParser(
        description="Map course Learning Outcomes to Program Outcomes using a local LLM."
    )
    parser.add_argument(
        "--program", required=True,
        help="Program folder name under data/raw/ and data/processed/ (e.g. ITSM)"
    )
    parser.add_argument(
        "--allow-cloud", action="store_true",
        help="Allow sending data to a cloud API if LLM_BASE_URL is not set. "
             "Use only if you have verified this is permitted by your data handling policy."
    )
    args = parser.parse_args()

    # ── Safety check ─────────────────────────────────────────────────────────
    check_cloud_safety(args.allow_cloud)

    run_ts   = datetime.now().strftime("%Y%m%d_%H%M")
    proc_dir = os.path.join(BASE_DIR, "data", "processed", args.program)

    curriculum_path = os.path.join(proc_dir, "curriculum_extracted.json")
    po_path         = os.path.join(proc_dir, "program_outcomes.json")
    output_path     = os.path.join(proc_dir, "alignment.json")
    archive_path    = os.path.join(proc_dir, f"alignment_{run_ts}.json")

    # ── Load Program Outcomes ─────────────────────────────────────────────────
    print(f"Loading Program Outcomes from {po_path}")
    with open(po_path, "r", encoding="utf-8") as f:
        po_data = json.load(f)
    po_list_text = "\n".join(f"{po['id']}: {po['text']}" for po in po_data)

    # ── Load curriculum ───────────────────────────────────────────────────────
    print(f"Loading curriculum from {curriculum_path}")
    with open(curriculum_path, "r", encoding="utf-8") as f:
        courses = json.load(f)

    total_outcomes = sum(len(c.get("outcomes", [])) for c in courses)
    processed      = 0
    results        = []

    for course in courses:
        code = course["course_code"]
        name = course["course_name"]
        print(f"\n{'─'*60}\nCourse: {code} — {name}")

        course_result = {"course_code": code, "course_name": name, "outcomes": []}

        for outcome in course.get("outcomes", []):
            processed += 1
            text       = outcome["outcome_text"]
            objectives = outcome.get("objectives", [])
            print(f"  [{processed}/{total_outcomes}] {text[:70]}...")

            alignment = align_outcome(code, name, text, objectives, po_list_text)
            print(f"    → Primary: {alignment['primary']}  Supporting: {alignment['supporting']}")

            course_result["outcomes"].append({
                "outcome_text": text,
                "primary":      alignment["primary"],
                "supporting":   alignment["supporting"],
            })

        results.append(course_result)

    # ── Wrap results with run metadata ───────────────────────────────────────
    output = {
        "_meta": {
            "program":      args.program,
            "run_timestamp": run_ts,
            "model":        LLM_MODEL,
            "endpoint":     LLM_BASE_URL or "OpenAI cloud (default)",
            "disclaimer":   "AI-generated alignment. Requires expert human review before official use.",
            "total_courses": len(results),
            "total_outcomes": sum(len(c["outcomes"]) for c in results),
        },
        "alignment": results,
    }

    os.makedirs(proc_dir, exist_ok=True)
    for path in (output_path, archive_path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(output, f, indent=2)

    print(f"\n{'='*60}")
    print(f"Done.")
    print(f"  Current:  {output_path}")
    print(f"  Archived: {archive_path}")


if __name__ == "__main__":
    main()
