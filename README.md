# NSCC Curriculum Outcome Mapping

A fully local pipeline for mapping course Learning Outcomes (LOs) to Program Outcomes (POs) using a locally-hosted LLM, with automated Word report generation.

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

Before running `llm_align.py`, edit the configuration block at the top of the script:

```python
LLM_BASE_URL = "http://YOUR-CLUSTER-IP:PORT/v1"
LLM_MODEL    = "your-model-name"
LLM_API_KEY  = "none"
```

## Usage

### 1. Extract curriculum from PDFs
```bash
python scripts/extract_curriculum.py
```
Place source PDF(s) in `data/raw/` first.

### 2. Run LLM alignment
```bash
python scripts/llm_align.py
```
Reads `data/processed/curriculum_extracted.json`, writes `data/processed/alignment.json`.

### 3. Generate report
```bash
python scripts/generate_report.py
```
Reads `data/processed/alignment.json`, writes `data/processed/LO_PO_Alignment_Report.docx`.

## Requirements

- Python 3.10+
- A locally-hosted LLM serving an OpenAI-compatible API (e.g. Ollama, vLLM, LM Studio)
- Tested with Qwen 2.5 32B
