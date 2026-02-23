> [!WARNING]
> This project is unmaintained and no longer actively developed.

# X21 Agent Testing and Benchmarking

Automated evaluation harness for testing the X21 Excel agent with test
specifications and against benchmark datasets.

## Supported Datasets

- **x21-demo** — Our custom curated test cases for X21-demo-specific
  functionality
- **SpreadsheetBench** — The
  [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench)
  verified 400 dataset
- Custom test cases

## Prerequisites

- **Windows** with Excel installed
- **Python 3.14+**
- **Deno server** running (`cd X21/deno-server && deno task start`)

## Setup

### Install dependencies

```bash
# Install pipenv (Python virtual environment manager) if you don't have it
pip install pipenv
# or
python -m pip install pipenv


# Install dependencies
cd X21/evaluation
pipenv install
#or
python -m pipenv install

```

## Test Cases

When releasing features, **always create or update test cases** so we can
regress the new behavior quickly. Use YAML cases for feature work because they
are fast, focused, and easy to review in PRs.

### How to run YAML cases

```bash
# Activate the pipenv virtual environment for the project
pipenv shell
# or
python -m pipenv shell

# Run all YAML cases in evaluation/cases (default)
python test.py

# Run a single YAML case file
python test.py --cases cases/x21-demo.yaml --limit 5

# Run a YAML case by ID
python test.py --cases cases/copy_paste_tool.yaml --id copy_paste_revert

# Run the echo batch with an absolute path
pipenv run python test.py --cases D:\Working\x21\X21\evaluation\cases\echo_batch.yaml
```

### Simple Feature Tests

When developing a feature, provide **either** assertions **or** a golden file
in your YAML case (or both). Tests without assertions/golden will fail.

### Assertions-based (fast, focused checks)

```yaml
cases:
  - id: "copy_paste_optimistic"
    name: "Copy paste optimistic"
    inputWorkbook: "test-cases/CopyPaste.xlsx"
    prompt: "Use copy_paste to copy Source!A1:B2 into Target!D4:E5."
    selectedTools:
      - copy_paste
    assertions:
      - sheet: "Target"
        cell: "D4"
        equalsText: "S1"
```

### Golden-file based (diff across a range)

```yaml
cases:
  - name: "fill_inputs"
    inputWorkbook: >-
      ../data/x21-samples/spreadsheet/fill_inputs/1_fill_inputs_init.xlsx
    goldenFile: >-
      ../data/x21-samples/spreadsheet/fill_inputs/1_fill_inputs_golden.xlsx
    answerPosition: "G28:G51"
    prompt: >-
      Fill the FY 24 column G28:G46 with the actuals from Sheet "Source Actuals"
      by linking with formulas
```

## Results

The `runs/` directory contains experiment runs:

### Experiment Runs (Gitignored)

All evaluation runs generate a timestamped JSON file inside the run directory:

```text
runs/20251209_141159/evaluation_results_20251209_141159.json
```

These files contain:

- Run metadata (date, duration, model)
- Summary statistics (pass rate, token usage)
- Individual test results with timing and tokens

**Experiment JSONs are gitignored** for day-to-day development and iteration.

### Mismatch Reports

For each failing test case, a detailed mismatch report is generated under:

```text
runs/{run_id}/misclassifications/{case}_mismatches.txt
```

Outputs are saved alongside in:

```text
runs/{run_id}/outputs/{case}_output.xlsx
```

Example:

```text
Test ID: 17-35
Answer Position: Sheet1!A1:B10
Total Mismatches: 3
============================================================

Cell: Sheet1!A3
  Expected: 150.25
  Actual:   150.0

Cell: Sheet1!B5
  Expected: 'Total'
  Actual:   None
```

This makes it easy to see exactly which cells differ and debug agent
behavior.

## How It Works

1. Opens `*_init.xlsx` in Excel via COM automation
2. Sends the prompt to the agent via WebSocket
3. Auto-approves all tool permission requests
4. Saves the result as `*_output.xlsx`
5. Compares output against `*_golden.xlsx` at specified cell ranges

## Directory Structure

<!-- markdownlint-disable MD013 -->
```text
evaluation/
├── README.md                          # This file
├── Pipfile / Pipfile.lock            # Python dependencies
│
├── test.py                            # Main orchestrator (entrypoint)
├── update_benchmarks.py               # Add results to BENCHMARKS.md
│
├── src/                               # Core evaluation code
│   ├── agent_client.py               # WebSocket client for deno server
│   ├── excel_controller.py           # Excel COM automation
│   └── evaluation_utils.py           # Workbook comparison logic
│
├── BENCHMARKS.md                      # Official benchmark tracking (committed)
│
├── data/                              # Test datasets (mostly gitignored)
│   ├── x21-samples/                  # Custom X21 test cases
│   │   ├── dataset.json
│   │   └── spreadsheet/
    │   │       └── {id}/
    │   │           ├── 1_{id}_init.xlsx       # Input
    │   │           ├── 1_{id}_golden.xlsx     # Expected output
    │   │           ├── 1_{id}_output.xlsx     # Agent output (gitignored)
    │   │           ├── 1_{id}_mismatches.txt  # Cell diff report (gitignored)
    │   │           └── *.pdf, *.png, ...      # Optional attachments (YAML prompt lives in cases/)
│   │
│   └── spreadsheetbench_verified_400/     # SpreadsheetBench dataset
│       ├── dataset.json
│       └── spreadsheet/
│
└── runs/                              # Evaluation run artifacts (gitignored)
```
<!-- markdownlint-enable MD013 -->

## Dataset Format

### YAML Cases (x21-samples and custom cases)

YAML cases live in `evaluation/cases/` and are the source of truth for
x21-samples prompts. Each case supports an `id` (used with `--id`), prompts,
tool selection, optional golden files, and assertions.

### SpreadsheetBench (dataset.json)

Each `dataset.json` entry requires `id` and `answer_position`. Optional
`attachments` array lists files (PDFs, images) to send with the prompt:

```json
{
    "id": "pdf_extraction",
    "answer_position": "A1:G35",
    "attachments": ["income_statement.pdf"]
}
```

Place attachment files in the test's spreadsheet directory alongside
`prompt.txt`.

## Benchmarking

Only for benchmarking do we use the SpreadsheetBench dataset.

### 1. Download the spreadsheet dataset

The `x21-samples` dataset is included in the repository.

**SpreadsheetBench**: Download the `spreadsheetbench_verified_400` dataset from
[Google Drive][spreadsheetbench-drive] and place it into `evaluation/data/`:

```text
evaluation/
└── data/
    ├── x21-samples/                         # Included in repo
    │   ├── dataset.json
    │   └── spreadsheet/
    └── spreadsheetbench_verified_400/       # Download from Google Drive
        ├── dataset.json
        └── spreadsheet/
```

### 2. Run evaluations on SpreadsheetBench

```bash
# Run SpreadsheetBench only
python test.py --dataset-dir data/spreadsheetbench_verified_400 --limit 5

# Run a SpreadsheetBench case by ID
python test.py --dataset-dir data/spreadsheetbench_verified_400 --id 10452
```

### Benchmark Evaluation (Global)

For official benchmarks, results are tracked in `BENCHMARKS.md`:

```bash
# Run a benchmark
python test.py --dataset-dir data/spreadsheetbench_verified_400

# Add result to benchmark table
python update_benchmarks.py \
  runs/20251209_141159/evaluation_results_20251209_141159.json \
  --notes "Fixed prompt handling"
```

**BENCHMARKS.md is committed** to maintain a history of performance over time.

## Attribution

`src/evaluation_utils.py` is adapted from the SpreadsheetBench evaluation code:
[SpreadsheetBench evaluation code][spreadsheetbench-eval]

If the upstream changes, you can copy the updated comparison functions from
there.

<!-- markdownlint-disable MD013 -->
[spreadsheetbench-drive]: https://drive.google.com/drive/folders/1MPZF6YH5qtV_rVthduO-NmpTJC-TKAey?usp=drive_link
[spreadsheetbench-eval]: https://github.com/RUCKBReasoning/SpreadsheetBench/blob/main/evaluation/evaluation.py
<!-- markdownlint-enable MD013 -->
