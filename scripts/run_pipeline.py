# run_pipeline.py
# Runs the full LO-to-PO alignment pipeline for a given program.
#
# Usage: python scripts/run_pipeline.py --program ITSM
#
# Expected input structure:
#   data/raw/<PROGRAM>/                  ← program outline PDF goes here
#   data/raw/<PROGRAM>/courses/          ← per-course RTF files go here
#
# Output written to:
#   data/processed/<PROGRAM>/program_outcomes.json
#   data/processed/<PROGRAM>/curriculum_extracted.json
#   data/processed/<PROGRAM>/alignment.json
#   data/processed/<PROGRAM>/LO_PO_Alignment_Report.docx

import argparse
import subprocess
import sys
import os

SCRIPTS = os.path.dirname(os.path.abspath(__file__))

STEPS = [
    ("Extract Program Outcomes", "extract_program_outcomes.py"),
    ("Extract Curriculum",       "extract_curriculum.py"),
    ("LLM Alignment",            "llm_align.py"),
    ("Generate Report",          "generate_report.py"),
]


def run_step(label: str, script: str, program: str):
    print(f"\n{'='*60}")
    print(f"  STEP: {label}")
    print(f"{'='*60}")
    result = subprocess.run(
        [sys.executable, os.path.join(SCRIPTS, script), "--program", program],
        check=False
    )
    if result.returncode != 0:
        print(f"\n  ERROR: '{label}' failed (exit code {result.returncode}). Pipeline stopped.")
        sys.exit(result.returncode)


def main():
    parser = argparse.ArgumentParser(
        description="Run the full LO-to-PO alignment pipeline for a program."
    )
    parser.add_argument(
        "--program", required=True,
        help="Program folder name under data/raw/ (e.g. ITSM)"
    )
    parser.add_argument(
        "--from-step", default=1, type=int, choices=[1, 2, 3, 4],
        help="Start from a specific step (1=extract POs, 2=extract curriculum, "
             "3=LLM align, 4=report). Useful for resuming after a failure."
    )
    args = parser.parse_args()

    print(f"\nPipeline starting for program: {args.program}")
    if args.from_step > 1:
        print(f"Resuming from step {args.from_step}.")

    for i, (label, script) in enumerate(STEPS, start=1):
        if i < args.from_step:
            print(f"  Skipping step {i}: {label}")
            continue
        run_step(label, script, args.program)

    print(f"\n{'='*60}")
    print(f"  Pipeline complete for: {args.program}")
    print(f"  Report: data/processed/{args.program}/LO_PO_Alignment_Report.docx")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
