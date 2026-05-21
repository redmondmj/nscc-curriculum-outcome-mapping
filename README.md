# NSCC Curriculum Outcome Mapping

A fully local pipeline for mapping course Learning Outcomes (LOs) to Program Outcomes (POs) using a locally-hosted LLM, with automated Word report generation.

## ⚠ Important

All pipeline outputs are AI-generated and require expert human review before use in any official
capacity. See [DISCLAIMER.md](DISCLAIMER.md) for full details on risks, data handling, and
intended use.

## Pipeline Overview

```
PDF curriculum docs
        ↓
extract_curriculum.py     → data/processed/curriculum_extracted.json
        ↓
llm_align.py              → data/processed/alignment.json
(calls local LLM via OpenAI-compatible API)
        ↓
generate_report.py        → data/processed/LO_PO_Alignment_Report.docx
(python-docx + matplotlib, no LLM required)
```

No cloud services required. All processing runs locally.

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

## Configuration

Copy `.env.example` to `.env` and uncomment the block that matches your setup:

```bash
cp .env.example .env
```

`.env.example` includes ready-to-use configurations for:
- **NSCC Truro Campus AI Cluster** — local vLLM running Qwen 32B (default)
- **Ollama** — any locally-pulled model
- **LM Studio** — any locally-loaded model
- **OpenAI** — cloud fallback (note: data leaves your network)

> **Note on Claude / Anthropic:** The Anthropic API is not OpenAI-compatible.
> The easiest path is to front it with [LiteLLM](https://github.com/BerriAI/litellm)
> and point `LLM_BASE_URL` at the LiteLLM proxy.

## Input folder structure

Create a folder under `data/raw/` named after your program, then drop your source documents in:

```
data/raw/
  ITSM/
    program_outline.pdf     ← program outline PDF (Program Outcomes extracted from here)
    courses/
      NETW1027.rtf          ← one RTF file per course
      OSYS1000.rtf
      ...
```

## Usage

### Run the full pipeline (recommended)
```bash
python scripts/run_pipeline.py --program ITSM
```
Runs all four steps in order. Use `--from-step 3` to resume from a specific step if something fails.

### Run steps individually
```bash
python scripts/extract_program_outcomes.py --program ITSM
python scripts/extract_curriculum.py       --program ITSM
python scripts/llm_align.py               --program ITSM
python scripts/generate_report.py         --program ITSM
```

All outputs are written to `data/processed/<PROGRAM>/`.

### Adding a second program
Just create `data/raw/CSD/` with the same structure and run:
```bash
python scripts/run_pipeline.py --program CSD
```

## Adding support for a unified course document (future)

When a single PDF containing all courses becomes available, add a parser function in
`extract_curriculum.py` and extend the `--format` flag choices. The `--format per-course-rtf`
default behaviour is unchanged.

## Requirements

- Python 3.10+
- A locally-hosted LLM serving an OpenAI-compatible API (e.g. Ollama, vLLM, LM Studio)
- Tested with Qwen 2.5 32B
