#!/usr/bin/env python3
"""
Helper script to add a benchmark result to BENCHMARKS.md

Usage:
    From evaluation/ directory:
    python update_benchmarks.py
        runs/20251209_141159/evaluation_results_20251209_141159.json
    python update_benchmarks.py
        runs/20251209_141159/evaluation_results_20251209_141159.json
        --notes "Fixed prompt handling"
"""

import argparse
import json
import os
from datetime import datetime


def add_benchmark_result(results_file: str, notes: str = ""):
    """Add a benchmark result row to BENCHMARKS.md"""

    # Read the results file
    with open(results_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Extract key metrics
    run_info = data["run_info"]
    summary = data["summary"]

    # Determine which benchmark (based on dataset_dir or cases_file)
    dataset_dir = run_info.get("dataset_dir", "")
    dataset_name = run_info.get("dataset_name", "") or ""
    cases_file = run_info.get("cases_file", "")

    if (
        "x21-samples" in dataset_dir
        or "x21-samples" in dataset_name.lower()
        or "x21-samples" in cases_file
    ):
        benchmark_name = "X21-Samples"
    elif "spreadsheetbench" in dataset_dir.lower():
        benchmark_name = "SpreadsheetBench"
    else:
        benchmark_name = "Unknown"

    # Format the data
    date = datetime.fromisoformat(run_info["start_time"]).strftime("%Y-%m-%d")
    model = run_info.get("model", "unknown")
    tests = f"{summary['passed']}/{summary['total_tests']}"
    pass_rate = f"{summary['pass_rate']:.1f}%"
    avg_tokens = f"{summary['average_tokens_per_test']:,.0f}"
    duration = f"{summary['total_duration_seconds']:.1f}s"

    # Create the table row
    row = (
        f"| {date} | {model} | {tests} | {pass_rate} | {avg_tokens} |"
        f" {duration} | {notes} |"
    )

    print(f"\nBenchmark: {benchmark_name}")
    print("Add this row to BENCHMARKS.md:\n")
    print(row)
    print()

    # Read BENCHMARKS.md (from evaluation/ directory)
    eval_dir = os.path.dirname(os.path.abspath(__file__))
    benchmarks_file = os.path.join(eval_dir, "BENCHMARKS.md")
    with open(benchmarks_file, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # Find the appropriate table and add the row
    marker = f"## {benchmark_name} Benchmark"
    found_section = False
    insert_index = None

    for i, line in enumerate(lines):
        if marker in line:
            found_section = True
        elif found_section and line.startswith("| -") and not insert_index:
            # This is a placeholder row or we're after the header
            insert_index = i + 1
        elif found_section and line.startswith("---"):
            # End of this section
            if not insert_index:
                insert_index = i
                # Skip back over any blank lines before the section separator
                while insert_index > 0 and lines[insert_index - 1].strip() == "":
                    insert_index -= 1
            break

    if insert_index:
        # Remove placeholder row if it exists
        if insert_index < len(lines) and "| -" in lines[insert_index - 1]:
            lines.pop(insert_index - 1)
            insert_index -= 1

        # Insert the new row (no extra blank line needed)
        lines.insert(insert_index, row + "\n")

        # Write back
        with open(benchmarks_file, "w", encoding="utf-8") as f:
            f.writelines(lines)

        print(f"✓ Added result to {benchmark_name} section in BENCHMARKS.md")
    else:
        print(f"✗ Could not find {benchmark_name} section in BENCHMARKS.md")
        print("Please add manually:")
        print(row)


def main():
    """Parse CLI args and update benchmarks."""
    parser = argparse.ArgumentParser(
        description="Add benchmark result to BENCHMARKS.md"
    )
    parser.add_argument("results_file", help="Path to evaluation results JSON file")
    parser.add_argument("--notes", default="", help="Optional notes about this run")

    args = parser.parse_args()

    if not os.path.exists(args.results_file):
        print(f"Error: File not found: {args.results_file}")
        return

    add_benchmark_result(args.results_file, args.notes)


if __name__ == "__main__":
    main()
