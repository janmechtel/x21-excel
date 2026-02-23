# Evaluation Benchmarks

Performance tracking for X21 agent on benchmark datasets.

---

## X21-Samples Benchmark

Small curated test set for quick validation.

<!-- markdownlint-disable MD013 -->
| Date | Model | Tests | Pass Rate | Avg Tokens/Test | Total Duration | Notes |
|------|-------|-------|-----------|-----------------|----------------|-------|
| 2025-12-09 | claude-sonnet-4-20250514 | 1/4 | 25.0% | 131,862 | 143.8s | baseline |
| 2025-12-09 | gpt-5 | 0/4 | 0.0% | 8,048 | 464.9s | GPT baseline |
| | | | | | | |
| 2025-12-10 | claude-sonnet-4-20250514 | 4/4 | 100.0% | 186,658 | 265.8s | Improve prompting: read before write; follow templates; format only on request; ask a question |
<!-- markdownlint-enable MD013 -->

---

## SpreadsheetBench Benchmark

Full SpreadsheetBench verified dataset (400 tests).

<!-- markdownlint-disable MD013 -->
| Date | Model | Tests | Pass Rate | Avg Tokens/Test | Total Duration | Notes |
|------|-------|-------|-----------|-----------------|----------------|-------|
| 2025-12-09 | claude-sonnet-4-20250514 | 2/5 | 40.0% | 79,471 | 449.1s | baseline |
| 2025-12-09 | gpt-5 | 0/5 | 0.0% | 7,451 | 537.8s | GPT baseline |
| | | | | | | |
<!-- markdownlint-enable MD013 -->

---

## How to Update

After running a benchmark evaluation:

```bash
# Run evaluation
python test.py --dataset-dir data/x21-samples

# Add result to table
python update_benchmarks.py \
  runs/20251209_141159/evaluation_results_20251209_141159.json \
  --notes "Your notes"
```

---

## Column Definitions

- **Date**: When the benchmark was run
- **Model**: LLM model used (e.g., `claude-sonnet-4-20250514`)
- **Tests**: Number passed / total tests
- **Pass Rate**: Percentage where output matches golden reference
- **Avg Tokens/Test**: Average total tokens (input + output) per test
- **Total Duration**: Wall-clock time for entire benchmark run
- **Notes**: Changes, fixes, or context for this run

Individual test results are saved in `runs/` (gitignored).
