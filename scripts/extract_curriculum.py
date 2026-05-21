# extract_curriculum.py
# Extracts course Learning Outcomes from per-course RTF files (which are HTML)
# and writes curriculum_extracted.json for use by llm_align.py
#
# Usage: python scripts/extract_curriculum.py --program ITSM
#
# Source docs go in: data/raw/<PROGRAM>/courses/
# Output written to: data/processed/<PROGRAM>/curriculum_extracted.json
#
# --format flag (default: per-course-rtf)
# TODO: add --format unified-pdf when a single combined course doc is available.
#       Slot new parser in as a separate function below extract_course_from_rtf()
#       and dispatch from main() based on the flag.

import argparse
import glob
import json
import os
from bs4 import BeautifulSoup


# ── Parsers ───────────────────────────────────────────────────────────────────

def extract_course_from_rtf(filepath: str) -> dict:
    """Parse a single RTF (HTML) course outline file."""
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        soup = BeautifulSoup(f, "html.parser")

    title_text  = soup.title.string if soup.title else ""
    course_code = ""
    course_name = ""

    if ":" in title_text:
        course_code, course_name = [x.strip() for x in title_text.split(":", 1)]
    else:
        course_code = os.path.basename(filepath).split("_")[0].strip()
        course_name = title_text

    lines = soup.get_text(separator="\n", strip=True).split("\n")

    outcomes          = []
    current_outcome   = None
    parsing_objectives = False

    for i, line in enumerate(lines):
        line = line.strip()

        if line.startswith("Learning Outcomes Display") or line.startswith("Other Course Notes"):
            break

        if line == "Outcome" and i + 1 < len(lines):
            if current_outcome:
                outcomes.append(current_outcome)
            current_outcome    = {"outcome_text": lines[i + 1].strip(), "objectives": []}
            parsing_objectives = False

        elif line == "Objectives":
            parsing_objectives = True

        elif parsing_objectives and current_outcome:
            if line and line not in ("Outcome", "Objectives"):
                current_outcome["objectives"].append(line)

    if current_outcome:
        outcomes.append(current_outcome)

    return {"course_code": course_code, "course_name": course_name, "outcomes": outcomes}


# TODO: def extract_courses_from_unified_pdf(filepath: str) -> list[dict]:
#     """Parse a single PDF containing all courses for the program."""
#     pass


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Extract course Learning Outcomes from curriculum source documents."
    )
    parser.add_argument(
        "--program", required=True,
        help="Program folder name under data/raw/ (e.g. ITSM)"
    )
    parser.add_argument(
        "--format", default="per-course-rtf",
        choices=["per-course-rtf"],   # extend here when unified-pdf is ready
        help="Source document format (default: per-course-rtf)"
    )
    args = parser.parse_args()

    base_dir    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    courses_dir = os.path.join(base_dir, "data", "raw",       args.program, "courses")
    out_dir     = os.path.join(base_dir, "data", "processed", args.program)
    output_path = os.path.join(out_dir, "curriculum_extracted.json")

    if not os.path.isdir(courses_dir):
        raise FileNotFoundError(
            f"Courses folder not found: {courses_dir}\n"
            f"Create it and drop your course documents inside."
        )

    all_courses = []

    if args.format == "per-course-rtf":
        rtf_files = sorted(glob.glob(os.path.join(courses_dir, "*.rtf")))
        print(f"Found {len(rtf_files)} .rtf course files in {courses_dir}")

        for filepath in rtf_files:
            try:
                data = extract_course_from_rtf(filepath)
                all_courses.append(data)
                print(f"  Extracted: {data['course_code']} — {data['course_name']} "
                      f"({len(data['outcomes'])} outcomes)")
            except Exception as e:
                print(f"  ERROR parsing {os.path.basename(filepath)}: {e}")

    os.makedirs(out_dir, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(all_courses, f, indent=4)

    print(f"\nSaved {len(all_courses)} courses to {output_path}")


if __name__ == "__main__":
    main()
