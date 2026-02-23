"""
Main evaluation orchestrator for X21 agent.

Usage:
    From evaluation/ directory:
    pipenv run python test.py --limit 5
    pipenv run python test.py --id 10452
"""

import argparse
import asyncio
import base64
import json
import mimetypes
import os
import shutil
import time
from datetime import datetime

import yaml
from dotenv import load_dotenv
from src.agent_client import AgentClient, test_connection
from src.evaluation_utils import compare_workbooks, evaluate_assertions
from src.excel_controller import ExcelController
from tqdm import tqdm

# Load environment variables
load_dotenv()

# Configuration
WS_PORT = os.getenv("WS_PORT", "8000")
WS_URL = f"ws://localhost:{WS_PORT}/ws"
TIMEOUT_SECONDS = 120

EVAL_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(EVAL_DIR, "data")
DATASET_DIR = os.path.join(DATA_DIR, "x21-samples")
RUNS_DIR = os.path.join(EVAL_DIR, "runs")


def load_dataset(dataset_dir: str) -> list:
    """Load the dataset.json file from the dataset directory."""
    dataset_path = os.path.join(dataset_dir, "dataset.json")
    with open(dataset_path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_prompt(test_id: str, dataset_dir: str) -> str:
    """Load the prompt.txt for a dataset.json test case."""
    prompt_path = os.path.join(dataset_dir, "spreadsheet", test_id, "prompt.txt")
    with open(prompt_path, "r", encoding="utf-8") as f:
        return f.read().strip()


def load_yaml_cases(cases_file: str) -> dict:
    """Load YAML test cases and normalize file paths."""
    with open(cases_file, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}

    cases = data.get("cases", data if isinstance(data, list) else [])
    if not isinstance(cases, list):
        raise ValueError("YAML cases must be a list or contain a 'cases' list")

    base_dir = os.path.dirname(os.path.abspath(cases_file))
    normalized = []
    for case in cases:
        if not isinstance(case, dict):
            continue
        normalized_case = dict(case)
        for key in ["inputWorkbook", "goldenFile", "outputWorkbook"]:
            if case.get(key):
                normalized_case[key] = os.path.abspath(
                    os.path.join(base_dir, case[key])
                )
        if case.get("attachments"):
            normalized_case["attachments"] = [
                os.path.abspath(os.path.join(base_dir, p)) for p in case["attachments"]
            ]
        normalized.append(normalized_case)

    return {
        "cases": normalized,
        "dataset_name": data.get("datasetName"),
        "cases_file": os.path.abspath(cases_file),
    }


def load_yaml_cases_dir(cases_dir: str) -> dict:
    """Load all YAML cases in a directory and combine them."""
    if not os.path.isdir(cases_dir):
        raise FileNotFoundError(f"Cases directory not found: {cases_dir}")

    entries = sorted(f for f in os.listdir(cases_dir) if f.lower().endswith(".yaml"))
    combined = []
    files = []
    for entry in entries:
        path = os.path.join(cases_dir, entry)
        data = load_yaml_cases(path)
        combined.extend(data.get("cases", []))
        files.append(path)

    return {
        "cases": combined,
        "dataset_name": "cases-dir",
        "cases_files": files,
        "cases_dir": os.path.abspath(cases_dir),
    }


def load_attachments_from_paths(paths: list) -> list:
    """Load and encode attachments from absolute paths."""
    attachments = []
    for path in paths:
        if not os.path.exists(path):
            raise FileNotFoundError(f"Attachment not found: {path}")

        with open(path, "rb") as f:
            content = f.read()

        filename = os.path.basename(path)
        mime_type, _ = mimetypes.guess_type(filename)
        if not mime_type:
            mime_type = "application/octet-stream"

        attachments.append(
            {
                "name": filename,
                "type": mime_type,
                "size": len(content),
                "base64": base64.b64encode(content).decode("utf-8"),
            }
        )
    return attachments


def build_output_path(input_path: str, output_path: str = None) -> str:
    """Build output workbook path from input workbook."""
    if output_path:
        return output_path
    root, ext = os.path.splitext(input_path)
    return f"{root}_output{ext or '.xlsx'}"


def sanitize_case_name(name: str) -> str:
    """Create a filesystem-safe case name."""
    cleaned = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in name)
    return cleaned.strip("_") or "case"


def write_mismatch_report(
    mismatches: list, mismatch_path: str, case_label: str, answer_position: str
) -> str:
    """Write a mismatch report to disk."""
    os.makedirs(os.path.dirname(mismatch_path), exist_ok=True)
    with open(mismatch_path, "w", encoding="utf-8") as f:
        f.write(f"Case: {case_label}\n")
        f.write(f"Answer Position: {answer_position}\n")
        f.write(f"Total Mismatches: {len(mismatches)}\n")
        f.write("=" * 60 + "\n\n")
        for m in mismatches:
            f.write(f"Cell: {m['sheet']}!{m['cell']}\n")
            f.write(f"  Expected: {repr(m['expected'])}\n")
            f.write(f"  Actual:   {repr(m['actual'])}\n")
            f.write("\n")
    return mismatch_path


def save_results_to_file(
    results: list, run_dir: str, run_info: dict, run_id: str = None
) -> str:
    """
    Save evaluation results to a JSON file.

    Args:
        results: List of test results
        run_dir: Directory for this run's artifacts
        run_info: Dictionary with run metadata (start_time, end_time, etc.)

    Returns:
        Path to the saved results file
    """
    # Create run directory if it doesn't exist
    os.makedirs(run_dir, exist_ok=True)

    # Generate filename with timestamp
    timestamp = run_id or datetime.now().strftime("%Y%m%d_%H%M%S")
    results_file = os.path.join(run_dir, f"evaluation_results_{timestamp}.json")

    # Calculate summary statistics
    total_tests = len(results)
    passed_tests = sum(1 for r in results if r.get("comparison_passed", False))
    failed_tests = total_tests - passed_tests

    total_input_tokens = sum(
        r.get("token_usage", {}).get("input_tokens", 0) for r in results
    )
    total_output_tokens = sum(
        r.get("token_usage", {}).get("output_tokens", 0) for r in results
    )
    total_tokens = total_input_tokens + total_output_tokens

    total_duration = sum(r.get("duration_seconds", 0) for r in results)
    avg_duration = total_duration / total_tests if total_tests > 0 else 0

    # Determine if all tests used the same model
    models = [r.get("model") for r in results if r.get("model")]
    if models and all(m == models[0] for m in models):
        # All tests used the same model - add to run_info
        run_info["model"] = models[0]
        # Remove model from individual test results
        for r in results:
            r.pop("model", None)

    # Build output structure
    output = {
        "run_info": run_info,
        "summary": {
            "total_tests": total_tests,
            "passed": passed_tests,
            "failed": failed_tests,
            "pass_rate": round(100 * passed_tests / total_tests, 2)
            if total_tests > 0
            else 0,
            "total_duration_seconds": round(total_duration, 2),
            "average_duration_seconds": round(avg_duration, 2),
            "total_tokens": total_tokens,
            "total_input_tokens": total_input_tokens,
            "total_output_tokens": total_output_tokens,
            "average_tokens_per_test": round(total_tokens / total_tests, 2)
            if total_tests > 0
            else 0,
        },
        "test_results": results,
    }

    # Write to file
    with open(results_file, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2)

    return results_file


def load_attachments(test_id: str, dataset_dir: str, attachment_names: list) -> list:
    """
    Load and encode attachment files for a test case.

    Args:
        test_id: The test case ID
        dataset_dir: Path to the dataset directory
        attachment_names: List of filenames to load from the test's spreadsheet
            directory

    Returns:
        List of attachment dicts with {name, type, size, base64}
    """
    attachments = []
    spreadsheet_dir = os.path.join(dataset_dir, "spreadsheet", test_id)

    for filename in attachment_names:
        file_path = os.path.join(spreadsheet_dir, filename)
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Attachment not found: {file_path}")

        # Read and encode
        with open(file_path, "rb") as f:
            content = f.read()

        # Detect MIME type
        mime_type, _ = mimetypes.guess_type(filename)
        if not mime_type:
            mime_type = "application/octet-stream"

        attachments.append(
            {
                "name": filename,
                "type": mime_type,
                "size": len(content),
                "base64": base64.b64encode(content).decode("utf-8"),
            }
        )

    return attachments


async def run_single_test(
    test_data: dict,
    dataset_dir: str,
    outputs_dir: str,
    misclassifications_dir: str,
    excel: ExcelController,
    agent: AgentClient,
    verbose: bool = False,
) -> dict:
    """
    Run a single test case.

    Returns a result dict with success/failure info.
    """
    test_id = str(test_data["id"])
    start_time = time.time()

    result = {
        "id": test_id,
        "instruction_type": test_data.get("instruction_type", ""),
        "success": False,
        "error": None,
        "comparison_passed": False,
        "start_time": datetime.fromtimestamp(start_time).isoformat(),
        "duration_seconds": 0,
        "token_usage": {"input_tokens": 0, "output_tokens": 0, "total_tokens": 0},
        "model": None,
    }

    # Paths
    spreadsheet_dir = os.path.join(dataset_dir, "spreadsheet", test_id)
    init_path = os.path.join(spreadsheet_dir, f"1_{test_id}_init.xlsx")
    golden_path = os.path.join(spreadsheet_dir, f"1_{test_id}_golden.xlsx")
    output_path = os.path.join(outputs_dir, f"{test_id}_output.xlsx")

    if not os.path.exists(init_path):
        result["error"] = "Init file not found"
        return result

    try:
        # 1. Copy init to output path
        shutil.copy2(init_path, output_path)

        # 2. Open output workbook in Excel
        if not excel.open_workbook(output_path):
            result["error"] = "Failed to open workbook"
            return result

        workbook_name = excel.get_workbook_name()
        workbook_path = excel.get_workbook_path()

        # 3. Load prompt
        prompt = load_prompt(test_id, dataset_dir)
        if verbose:
            print(f"\n  Prompt: {prompt[:100]}...")

        # 4. Load attachments if specified
        attachments = None
        attachment_names = test_data.get("attachments", [])
        if attachment_names:
            attachments = load_attachments(test_id, dataset_dir, attachment_names)
            if verbose:
                print(f"  Attachments: {[a['name'] for a in attachments]}")

        # 5. Run agent
        def on_message(msg):
            if verbose:
                msg_type = msg.get("type", "")
                # Skip streaming messages - they're handled by agent_client
                # inline printing
                if msg_type in ["content_block_delta", "text", "stream:delta"]:
                    pass
                else:
                    print(f"    [{msg_type}]")

        agent_result = await agent.run_prompt(
            prompt=prompt,
            workbook_name=workbook_name,
            workbook_path=workbook_path,
            attachments=attachments,
            on_message=on_message if verbose else None,
        )

        # Extract usage information from agent result
        if isinstance(agent_result, dict):
            success = agent_result.get("success", False)
            if "usage" in agent_result:
                usage = agent_result["usage"]
                result["token_usage"] = {
                    "input_tokens": usage.get("input_tokens", 0),
                    "output_tokens": usage.get("output_tokens", 0),
                    "total_tokens": usage.get("input_tokens", 0)
                    + usage.get("output_tokens", 0),
                }
            if "model" in agent_result:
                result["model"] = agent_result["model"]
        else:
            success = agent_result

        if not success:
            result["error"] = "Agent execution failed or timed out"
            result["duration_seconds"] = time.time() - start_time
            excel.close_workbook(save=False)
            return result

        # 6. Save and close
        excel.close_workbook(save=True)
        result["success"] = True

        # 7. Compare with golden
        answer_position = test_data.get("answer_position", "A1:Z100")
        comparison_passed, msg, mismatches = compare_workbooks(
            golden_path,
            output_path,
            test_data.get("instruction_type", ""),
            answer_position,
        )
        result["comparison_passed"] = comparison_passed
        if not comparison_passed and msg:
            result["comparison_error"] = msg

        # 8. Write mismatch report if there are mismatches
        if mismatches:
            mismatch_path = os.path.join(
                misclassifications_dir, f"{test_id}_mismatches.txt"
            )
            result["mismatch_report"] = write_mismatch_report(
                mismatches, mismatch_path, test_id, answer_position
            )

    except Exception as e:
        result["error"] = str(e)
        result["duration_seconds"] = time.time() - start_time
        try:
            excel.close_workbook(save=False)
        except Exception:
            pass

    # Record final duration
    result["duration_seconds"] = time.time() - start_time
    result["end_time"] = datetime.now().isoformat()

    return result


async def run_yaml_case(
    case_data: dict,
    outputs_dir: str,
    misclassifications_dir: str,
    excel: ExcelController,
    agent: AgentClient,
    verbose: bool = False,
) -> dict:
    """Run a single YAML-defined test case."""
    case_id = case_data.get("id") or case_data.get("name") or "unknown"
    case_name = case_data.get("name") or case_id
    start_time = time.time()

    result = {
        "id": case_id,
        "case_name": case_name,
        "success": False,
        "error": None,
        "comparison_passed": False,
        "start_time": datetime.fromtimestamp(start_time).isoformat(),
        "duration_seconds": 0,
        "token_usage": {"input_tokens": 0, "output_tokens": 0, "total_tokens": 0},
        "model": None,
    }

    input_path = case_data.get("inputWorkbook")
    golden_path = case_data.get("goldenFile")
    if not input_path or not os.path.exists(input_path):
        result["error"] = "Input workbook not found"
        return result

    safe_name = sanitize_case_name(case_id)
    output_path = os.path.join(outputs_dir, f"{safe_name}_output.xlsx")

    try:
        shutil.copy2(input_path, output_path)

        if not excel.open_workbook(output_path):
            result["error"] = "Failed to open workbook"
            return result

        workbook_name = excel.get_workbook_name()
        workbook_path = excel.get_workbook_path()

        prompt = (case_data.get("prompt") or "").strip()
        if not prompt:
            result["error"] = "Prompt not provided"
            excel.close_workbook(save=False)
            return result
        if verbose:
            print(f"\n  Prompt: {prompt[:100]}...")

        attachments = None
        if case_data.get("attachments"):
            attachments = load_attachments_from_paths(case_data["attachments"])
            if verbose:
                print(f"  Attachments: {[a['name'] for a in attachments]}")

        def on_message(msg):
            if verbose:
                msg_type = msg.get("type", "")
                if msg_type in ["content_block_delta", "text", "stream:delta"]:
                    pass
                else:
                    print(f"    [{msg_type}]")

        active_tools = case_data.get("selectedTools")
        agent_result = await agent.run_prompt(
            prompt=prompt,
            workbook_name=workbook_name,
            workbook_path=workbook_path,
            active_tools=active_tools,
            attachments=attachments,
            on_message=on_message if verbose else None,
        )

        if isinstance(agent_result, dict):
            success = agent_result.get("success", False)
            if "usage" in agent_result:
                usage = agent_result["usage"]
                result["token_usage"] = {
                    "input_tokens": usage.get("input_tokens", 0),
                    "output_tokens": usage.get("output_tokens", 0),
                    "total_tokens": usage.get("input_tokens", 0)
                    + usage.get("output_tokens", 0),
                }
            if "model" in agent_result:
                result["model"] = agent_result["model"]
        else:
            success = agent_result

        if not success:
            result["error"] = "Agent execution failed or timed out"
            result["duration_seconds"] = time.time() - start_time
            excel.close_workbook(save=False)
            return result

        tool_use_ids = []
        if isinstance(agent_result, dict):
            tool_use_ids = agent_result.get("tool_use_ids", [])

        if case_data.get("revertAfterTool") and tool_use_ids:
            revert_ok = await agent.revert_tools(workbook_name, tool_use_ids)
            result["revert_success"] = revert_ok
            if not revert_ok:
                result["error"] = "Failed to revert tool changes"
                excel.close_workbook(save=False)
                return result

        excel.close_workbook(save=True)
        result["success"] = True

        assertion_passed = True
        assertion_msg = ""
        assertion_mismatches = []
        if case_data.get("assertions"):
            assertion_passed, assertion_msg, assertion_mismatches = evaluate_assertions(
                output_path, case_data.get("assertions", [])
            )
            result["assertions_passed"] = assertion_passed
            if assertion_mismatches:
                result["assertion_mismatches"] = assertion_mismatches
                assertion_report_path = os.path.join(
                    misclassifications_dir, f"{safe_name}_assertions_mismatches.txt"
                )
                result["assertion_mismatch_report"] = write_mismatch_report(
                    assertion_mismatches, assertion_report_path, case_id, "assertions"
                )

        golden_passed = True
        golden_msg = ""
        golden_mismatches = []
        if golden_path:
            if not os.path.exists(golden_path):
                result["comparison_passed"] = False
                result["comparison_error"] = "Golden file not found"
                result["duration_seconds"] = time.time() - start_time
                return result
            answer_position = case_data.get("answerPosition", "A1:Z100")
            golden_passed, golden_msg, golden_mismatches = compare_workbooks(
                golden_path, output_path, "", answer_position
            )
            result["golden_passed"] = golden_passed
            if golden_mismatches:
                result["golden_mismatches"] = golden_mismatches
                mismatch_path = os.path.join(
                    misclassifications_dir, f"{safe_name}_mismatches.txt"
                )
                result["mismatch_report"] = write_mismatch_report(
                    golden_mismatches, mismatch_path, case_id, answer_position
                )

        if case_data.get("assertions") and golden_path:
            result["comparison_passed"] = assertion_passed and golden_passed
        elif case_data.get("assertions"):
            result["comparison_passed"] = assertion_passed
        elif golden_path:
            result["comparison_passed"] = golden_passed
        else:
            result["comparison_passed"] = False
            result["comparison_error"] = "No assertions or golden file provided"

        if not result["comparison_passed"]:
            combined = ", ".join([m for m in [assertion_msg, golden_msg] if m])
            if combined:
                result["comparison_error"] = combined

    except Exception as e:
        result["error"] = str(e)
        result["duration_seconds"] = time.time() - start_time
        try:
            excel.close_workbook(save=False)
        except Exception as close_err:
            # Ignore errors during cleanup, but record them for debugging.
            result["close_error"] = str(close_err)

    result["duration_seconds"] = time.time() - start_time
    result["end_time"] = datetime.now().isoformat()

    return result


async def run_evaluation(
    dataset_dir: str,
    limit: int = None,
    test_id: str = None,
    verbose: bool = False,
    ws_url: str = WS_URL,
    cases_file: str = None,
) -> list:
    """Run evaluation on dataset."""

    # Track run start time
    run_start_time = time.time()
    run_start_timestamp = datetime.now().isoformat()
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = os.path.join(RUNS_DIR, run_id)
    outputs_dir = os.path.join(run_dir, "outputs")
    misclassifications_dir = os.path.join(run_dir, "misclassifications")
    os.makedirs(outputs_dir, exist_ok=True)
    os.makedirs(misclassifications_dir, exist_ok=True)

    dataset = None
    cases_metadata = None
    cases_dir = os.path.join(EVAL_DIR, "cases")
    use_cases = False
    total_cases_in_dataset = 0

    if cases_file:
        cases_metadata = load_yaml_cases(cases_file)
        dataset = cases_metadata["cases"]
        total_cases_in_dataset = len(dataset)
        use_cases = True
    elif "spreadsheetbench" in (dataset_dir or "").lower():
        dataset = load_dataset(dataset_dir)
        total_cases_in_dataset = len(dataset)
    else:
        cases_metadata = load_yaml_cases_dir(cases_dir)
        dataset = cases_metadata["cases"]
        total_cases_in_dataset = len(dataset)
        use_cases = True

    print(f"Loaded {len(dataset)} test cases")

    # Filter by ID if specified
    if test_id:
        if use_cases:
            dataset = [
                d for d in dataset if str(d.get("id") or d.get("name")) == test_id
            ]
        else:
            dataset = [d for d in dataset if str(d["id"]) == test_id]
        if not dataset:
            print(f"Test ID '{test_id}' not found")
            return []

    # Apply limit
    if limit:
        dataset = dataset[:limit]

    print(f"Running {len(dataset)} test cases")

    # Check deno server connection
    print(f"\nChecking deno server at {ws_url}...")
    if not await test_connection(ws_url):
        print("ERROR: Cannot connect to deno server. Make sure it's running:")
        print("  cd X21/deno-server && deno task dev")
        return []
    print("✓ Deno server connected")

    # Connect to Excel
    excel = ExcelController()
    if not excel.connect():
        print("ERROR: Cannot connect to Excel. Make sure Excel is running.")
        return []
    print("✓ Excel connected")

    # Connect agent client
    agent = AgentClient(ws_url=ws_url, timeout=TIMEOUT_SECONDS)
    await agent.connect()
    print("✓ Agent client connected")

    # Run tests
    results = []
    passed = 0
    failed = 0

    try:
        pbar = tqdm(dataset, desc="Evaluating")
        for test_data in pbar:
            if use_cases:
                result = await run_yaml_case(
                    test_data,
                    outputs_dir,
                    misclassifications_dir,
                    excel,
                    agent,
                    verbose,
                )
            else:
                result = await run_single_test(
                    test_data,
                    dataset_dir,
                    outputs_dir,
                    misclassifications_dir,
                    excel,
                    agent,
                    verbose,
                )
            results.append(result)

            if result["comparison_passed"]:
                passed += 1
                tqdm.write(f"  ✓ {result['id']}: passed")
            else:
                failed += 1
                error_detail = result.get("error") or result.get(
                    "comparison_error", "mismatch"
                )
                tqdm.write(f"  ✗ {result['id']}: {error_detail}")

            # Update progress bar with success rate
            total = passed + failed
            rate = 100 * passed / total if total > 0 else 0
            pbar.set_postfix_str(f"{passed}/{total} passed ({rate:.0f}%)")

    finally:
        await agent.close()

    # Calculate run statistics
    run_end_time = time.time()
    run_end_timestamp = datetime.now().isoformat()
    run_duration = run_end_time - run_start_time

    # Print summary
    print(f"\n{'=' * 50}")
    print(
        f"Results: {passed}/{len(results)} passed ({100 * passed / len(results):.1f}%)"
    )
    print(f"Total Duration: {run_duration:.2f} seconds")

    # Calculate token usage
    total_input_tokens = sum(
        r.get("token_usage", {}).get("input_tokens", 0) for r in results
    )
    total_output_tokens = sum(
        r.get("token_usage", {}).get("output_tokens", 0) for r in results
    )
    total_tokens = total_input_tokens + total_output_tokens

    if total_tokens > 0:
        total_tokens_display = (
            f"Total Tokens: {total_tokens:,} (Input: {total_input_tokens:,}, "
            f"Output: {total_output_tokens:,})"
        )
        print(total_tokens_display)
        print(f"Average Tokens per Test: {total_tokens / len(results):.0f}")

    print(f"{'=' * 50}")

    # Save results to file
    run_info = {
        "start_time": run_start_timestamp,
        "end_time": run_end_timestamp,
        "duration_seconds": round(run_duration, 2),
        "dataset_dir": dataset_dir,
        "limit": limit,
        "test_id": test_id,
        "total_tests_in_dataset": total_cases_in_dataset,
        "tests_run": len(results),
        "run_id": run_id,
        "outputs_dir": outputs_dir,
        "misclassifications_dir": misclassifications_dir,
    }
    if use_cases:
        if cases_file:
            run_info["cases_file"] = cases_file
        if cases_metadata:
            if cases_metadata.get("dataset_name"):
                run_info["dataset_name"] = cases_metadata["dataset_name"]
            if cases_metadata.get("cases_dir"):
                run_info["cases_dir"] = cases_metadata["cases_dir"]
            if cases_metadata.get("cases_files"):
                run_info["cases_files"] = cases_metadata["cases_files"]

    results_file = save_results_to_file(results, run_dir, run_info, run_id=run_id)
    print(f"\n✓ Results saved to: {results_file}")

    return results


def main():
    """Parse CLI args and run evaluation."""
    parser = argparse.ArgumentParser(description="Run X21 agent evaluation")
    parser.add_argument("--limit", type=int, help="Limit number of test cases")
    parser.add_argument("--id", type=str, help="Run specific test case by ID")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    parser.add_argument("--ws-url", type=str, default=WS_URL, help="WebSocket URL")
    parser.add_argument(
        "--dataset-dir",
        type=str,
        default=DATASET_DIR,
        help="SpreadsheetBench dataset directory",
    )
    parser.add_argument("--cases", type=str, help="YAML test cases file")

    args = parser.parse_args()

    asyncio.run(
        run_evaluation(
            dataset_dir=args.dataset_dir,
            limit=args.limit,
            test_id=args.id,
            verbose=args.verbose,
            ws_url=args.ws_url,
            cases_file=args.cases,
        )
    )


if __name__ == "__main__":
    main()
