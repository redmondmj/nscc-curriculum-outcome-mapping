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

# ── LLM client ────────────────────────────────────────────────────────────────
# base_url is optional — omitting it uses the OpenAI default (cloud)
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
    args = parser.parse_args()

    proc_dir        = os.path.join(BASE_DIR, "data", "processed", args.program)
    curriculum_path = os.path.join(proc_dir, "curriculum_extracted.json")
    po_path         = os.path.join(proc_dir, "program_outcomes.json")
    output_path     = os.path.join(proc_dir, "alignment.json")

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

    os.makedirs(proc_dir, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    print(f"\n{'='*60}\nDone. Alignment written to {output_path}")


if __name__ == "__main__":
    main()
