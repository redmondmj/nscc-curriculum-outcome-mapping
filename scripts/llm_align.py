# llm_align.py
# Uses a local LLM (via OpenAI-compatible API) to map course Learning Outcomes
# to Program Outcomes, then writes alignment.json for use by generate_report.py
#
# Copy .env.example to .env and fill in your LLM details before running.
# Run: python scripts/llm_align.py

import json
import os
from openai import OpenAI
from dotenv import load_dotenv

# ── Configuration (loaded from .env) ─────────────────────────────────────────
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))

LLM_BASE_URL = os.getenv("LLM_BASE_URL")    # None = use OpenAI default (cloud)
LLM_MODEL    = os.getenv("LLM_MODEL",    "gpt-4o")
LLM_API_KEY  = os.getenv("LLM_API_KEY",  "none")

BASE_DIR        = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CURRICULUM_PATH = os.path.join(BASE_DIR, "data", "processed", "curriculum_extracted.json")
OUTPUT_PATH     = os.path.join(BASE_DIR, "data", "processed", "alignment.json")

# ── Program Outcomes ──────────────────────────────────────────────────────────
PROGRAM_OUTCOMES = [
    "PO1: Analyze and document business systems, requirements and problems using standard methodologies and notation.",
    "PO2: Design, implement, and maintain a secure networked environment.",
    "PO3: Configure, secure and administer network operating systems, software and hardware to support business systems.",
    "PO4: Provide technical training to support business systems.",
    "PO5: Research, learn and integrate innovations in systems management and security.",
    "PO6: Integrate professional practices and skills into all projects, activities and communications in the context of an IT industry environment.",
    "PO7: Demonstrate continuous professional improvement through reflection and modification of processes and approaches in relation to the IT industries.",
    "PO8: Blend service and learning in ways that use program-related skills, knowledge and behaviours to serve others at the campus, within the College and in the community.",
    "PO9: Apply a Portfolio approach to the personal management of learning and career planning relating to the learner's occupational readiness.",
    "PO10: Apply the Essential & Employability Skills needed to enter, stay in, and progress in the world of work, productively contributing to the economy and the community.",
    "PO11: Apply sustainable practices that support economic, social, cultural and environmental stewardship.",
    "PO12: Demonstrate the principles of quality and safety as per the 5S+S standard, complete WHMIS and OH&S requirements.",
]

PO_LIST_TEXT = "\n".join(PROGRAM_OUTCOMES)

# ── LLM client ────────────────────────────────────────────────────────────────
# base_url is optional — omitting it uses the OpenAI default (cloud)
client_kwargs = {"api_key": LLM_API_KEY}
if LLM_BASE_URL:
    client_kwargs["base_url"] = LLM_BASE_URL
client = OpenAI(**client_kwargs)


def align_outcome(course_code, course_name, outcome_text, objectives):
    """Ask the LLM to map a single course LO to Program Outcomes."""
    objectives_text = "\n".join(f"  - {o}" for o in objectives if o)

    prompt = f"""You are a curriculum alignment expert for a college IT diploma program.

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
{PO_LIST_TEXT}

Respond ONLY with a valid JSON object in this exact format:
{{
  "primary": ["PO1", "PO3"],
  "supporting": ["PO5", "PO6"]
}}
Use only PO1 through PO12. Both arrays may be empty but must be present."""

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
        # Normalise keys to lowercase just in case
        return {
            "primary":    result.get("primary", []),
            "supporting": result.get("supporting", []),
        }
    except json.JSONDecodeError:
        print(f"  WARNING: Could not parse JSON for {course_code} — '{outcome_text[:60]}...'")
        print(f"  Raw response: {raw}")
        return {"primary": [], "supporting": []}


def main():
    print(f"Loading curriculum from {CURRICULUM_PATH}")
    with open(CURRICULUM_PATH, "r", encoding="utf-8") as f:
        courses = json.load(f)

    results = []
    total_outcomes = sum(len(c.get("outcomes", [])) for c in courses)
    processed = 0

    for course in courses:
        code = course["course_code"]
        name = course["course_name"]
        print(f"\n{'─'*60}")
        print(f"Course: {code} — {name}")

        course_result = {
            "course_code": code,
            "course_name": name,
            "outcomes": []
        }

        for outcome in course.get("outcomes", []):
            processed += 1
            text = outcome["outcome_text"]
            objectives = outcome.get("objectives", [])
            print(f"  [{processed}/{total_outcomes}] {text[:70]}...")

            alignment = align_outcome(code, name, text, objectives)
            print(f"    → Primary: {alignment['primary']}  Supporting: {alignment['supporting']}")

            course_result["outcomes"].append({
                "outcome_text": text,
                "primary":      alignment["primary"],
                "supporting":   alignment["supporting"],
            })

        results.append(course_result)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    print(f"\n{'='*60}")
    print(f"Done. Alignment written to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
