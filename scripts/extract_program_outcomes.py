# extract_program_outcomes.py
# Extracts the numbered Program Outcomes list from a program outline PDF
# and writes program_outcomes.json for use by llm_align.py and generate_report.py
#
# Usage: python scripts/extract_program_outcomes.py --program ITSM

import argparse
import json
import os
import re
import pdfplumber


def extract_pos_from_pdf(pdf_path: str) -> list[dict]:
    """
    Parse a numbered Program Outcomes list from an NSCC program outline PDF.
    Expects a section headed 'Program Outcomes' containing numbered items.
    Returns a list of dicts: [{"id": "PO1", "text": "..."}, ...]
    """
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"

    lines = [l.strip() for l in full_text.split("\n")]

    # ── Find the Program Outcomes section ────────────────────────────────────
    SECTION_START = re.compile(r'^program outcomes$', re.IGNORECASE)
    # Only stop at top-level structural headings — NOT inline explanatory text
    # like "Essential Skills:" or "Employability Skills:" which appear mid-section
    SECTION_END   = re.compile(
        r'^(program admission|graduation requirements|employment opportunities'
        r'|additional information|accreditation|advisement|plan of study|milestones)',
        re.IGNORECASE
    )
    # Also filter out PDF header/footer bleed lines (e.g. "ITSM: IT Systems... Program Outline 1")
    FOOTER_LINE   = re.compile(r'program outline\s+\d+$', re.IGNORECASE)
    PO_LINE       = re.compile(r'^(\d{1,2})\.\s+(.+)')

    in_section  = False
    outcomes    = []
    current_num = None
    current_txt = ""

    for line in lines:
        if not in_section:
            if SECTION_START.match(line):
                in_section = True
            continue

        if SECTION_END.match(line):
            break

        # Skip PDF header/footer bleed lines
        if FOOTER_LINE.search(line):
            continue

        m = PO_LINE.match(line)
        if m:
            if current_num is not None:
                outcomes.append({
                    "id":   f"PO{current_num}",
                    "text": current_txt.strip()
                })
            current_num = int(m.group(1))
            current_txt = m.group(2)
        elif current_num is not None and line:
            # Continuation of the previous PO text (line-wrapped in PDF)
            current_txt += " " + line

    if current_num is not None and current_txt:
        outcomes.append({
            "id":   f"PO{current_num}",
            "text": current_txt.strip()
        })

    return outcomes


def main():
    parser = argparse.ArgumentParser(
        description="Extract Program Outcomes from a program outline PDF."
    )
    parser.add_argument(
        "--program", required=True,
        help="Program folder name under data/raw/ (e.g. ITSM)"
    )
    args = parser.parse_args()

    base_dir    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    raw_dir     = os.path.join(base_dir, "data", "raw",       args.program)
    out_dir     = os.path.join(base_dir, "data", "processed", args.program)
    output_path = os.path.join(out_dir, "program_outcomes.json")

    # Find the program outline PDF
    pdf_candidates = [
        f for f in os.listdir(raw_dir)
        if f.lower().endswith(".pdf")
    ]
    if not pdf_candidates:
        raise FileNotFoundError(
            f"No PDF found in {raw_dir}. "
            f"Drop the program outline PDF there and try again."
        )
    if len(pdf_candidates) > 1:
        print(f"Multiple PDFs found — using: {pdf_candidates[0]}")
    pdf_path = os.path.join(raw_dir, pdf_candidates[0])

    print(f"Extracting Program Outcomes from: {pdf_path}")
    outcomes = extract_pos_from_pdf(pdf_path)

    if not outcomes:
        raise ValueError(
            "No Program Outcomes found. Check that the PDF contains a "
            "'Program Outcomes' section with a numbered list."
        )

    os.makedirs(out_dir, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(outcomes, f, indent=2)

    print(f"Extracted {len(outcomes)} Program Outcomes:")
    for po in outcomes:
        print(f"  {po['id']}: {po['text'][:80]}{'...' if len(po['text']) > 80 else ''}")
    print(f"\nSaved to: {output_path}")


if __name__ == "__main__":
    main()
