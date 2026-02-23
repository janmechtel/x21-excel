import { assert, assertEquals } from "std/assert";
import {
  analyzeRangeOverlaps,
  columnToNumber,
  expandRange,
  expandRangeByIndex,
  findAllRectangles,
  numberToColumn,
  parseRange,
} from "../src/utils/excel-range.ts";

/**
 * Check if two ranges overlap
 */
function rangesOverlap(range1: string, range2: string): boolean {
  const cells1 = getCellsFromRange(range1);
  const cells2 = getCellsFromRange(range2);

  for (const cell of cells1) {
    if (cells2.has(cell)) {
      return true;
    }
  }

  return false;
}

/**
 * Get all individual cells from a range string (using the functions from main module)
 */
function getCellsFromRange(range: string): Set<string> {
  const cells = new Set<string>();
  const parsed = parseRange(range);

  const startCol = columnToNumber(parsed.startColumn);
  const endCol = columnToNumber(parsed.endColumn);

  for (let row = parsed.startRow; row <= parsed.endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      const colLetter = numberToColumn(col);
      cells.add(`${colLetter}${row}`);
    }
  }

  return cells;
}

/**
 * Get all cells from a list of ranges
 */
function getCellsFromRanges(ranges: string[]): Set<string> {
  const allCells = new Set<string>();

  for (const range of ranges) {
    const rangeCells = getCellsFromRange(range);
    rangeCells.forEach((cell) => allCells.add(cell));
  }

  return allCells;
}

/**
 * Check if all result ranges are non-overlapping
 */
function validateNoOverlaps(ranges: string[]): boolean {
  for (let i = 0; i < ranges.length; i++) {
    for (let j = i + 1; j < ranges.length; j++) {
      if (rangesOverlap(ranges[i], ranges[j])) {
        console.log(`Overlap found between ${ranges[i]} and ${ranges[j]}`);
        return false;
      }
    }
  }
  return true;
}

/**
 * Check if two sets of cells are equal
 */
function cellSetsEqual(set1: Set<string>, set2: Set<string>): boolean {
  if (set1.size !== set2.size) return false;

  for (const cell of set1) {
    if (!set2.has(cell)) return false;
  }

  return true;
}

Deno.test("findAllRectangles - basic functionality", () => {
  const matrix = [
    [1, 1],
    [1, 1],
  ];

  const result = findAllRectangles(matrix);

  // Should find one rectangle
  assertEquals(result.length, 1);
  assertEquals(result[0].value, 1);
  assertEquals(result[0].topLeft, [0, 0]);
  assertEquals(result[0].bottomRight, [1, 1]);
});

Deno.test("findAllRectangles - empty matrix", () => {
  const matrix: number[][] = [];
  const result = findAllRectangles(matrix);
  assertEquals(result, []);
});

Deno.test("findAllRectangles - single cell", () => {
  const matrix = [[5]];
  const result = findAllRectangles(matrix);

  assertEquals(result.length, 1);
  assertEquals(result[0].value, 5);
  assertEquals(result[0].topLeft, [0, 0]);
  assertEquals(result[0].bottomRight, [0, 0]);
});

Deno.test("findAllRectangles - horizontal line", () => {
  const matrix = [[1, 1, 1]];

  const result = findAllRectangles(matrix);

  assertEquals(result.length, 1);
  assertEquals(result[0].value, 1);
  assertEquals(result[0].topLeft, [0, 0]);
  assertEquals(result[0].bottomRight, [0, 2]);
});

Deno.test("findAllRectangles - vertical line", () => {
  const matrix = [
    [1],
    [1],
    [1],
  ];

  const result = findAllRectangles(matrix);

  assertEquals(result.length, 1);
  assertEquals(result[0].value, 1);
  assertEquals(result[0].topLeft, [0, 0]);
  assertEquals(result[0].bottomRight, [2, 0]);
});

Deno.test("findAllRectangles - returns valid rectangles", () => {
  const matrix = [
    [1, 0, 2],
    [1, 0, 2],
  ];

  const result = findAllRectangles(matrix);

  // Should return some rectangles
  assert(result.length > 0);

  // All rectangles should have valid coordinates
  result.forEach((rect) => {
    assert(rect.topLeft[0] >= 0, "Top row should be non-negative");
    assert(rect.topLeft[1] >= 0, "Left column should be non-negative");
    assert(
      rect.bottomRight[0] >= rect.topLeft[0],
      "Bottom row should be >= top row",
    );
    assert(
      rect.bottomRight[1] >= rect.topLeft[1],
      "Right column should be >= left column",
    );
    assert(
      rect.bottomRight[0] < matrix.length,
      "Bottom row should be within matrix bounds",
    );
    assert(
      rect.bottomRight[1] < matrix[0].length,
      "Right column should be within matrix bounds",
    );
  });
});

Deno.test("findAllRectangles - preserves values", () => {
  const matrix = [
    [1, 2],
    [3, 4],
  ];

  const result = findAllRectangles(matrix);

  // Should find rectangles for all unique values
  const foundValues = result.map((r) => r.value);
  const uniqueValues = [...new Set(foundValues)];

  // Should have found rectangles for values 1, 2, 3, 4
  assert(uniqueValues.includes(1), "Should find rectangle with value 1");
  assert(uniqueValues.includes(2), "Should find rectangle with value 2");
  assert(uniqueValues.includes(3), "Should find rectangle with value 3");
  assert(uniqueValues.includes(4), "Should find rectangle with value 4");
});

Deno.test("findAllRectangles - area calculation", () => {
  const matrix = [
    [1, 1, 1],
    [1, 1, 1],
  ];

  const result = findAllRectangles(matrix);

  assertEquals(result.length, 1);

  const rect = result[0];
  const area = (rect.bottomRight[0] - rect.topLeft[0] + 1) *
    (rect.bottomRight[1] - rect.topLeft[1] + 1);

  assertEquals(area, 6); // 2 rows × 3 columns = 6
});

Deno.test("findAllRectangles - handles negative values", () => {
  const matrix = [
    [-1, -1],
    [-1, -1],
  ];

  const result = findAllRectangles(matrix);

  assertEquals(result.length, 1);
  assertEquals(result[0].value, -1);
});

Deno.test("findAllRectangles - format index simulation", () => {
  // Simple case: all cells have same format (index 0)
  const matrix = [
    [0, 0],
    [0, 0],
  ];

  const result = findAllRectangles(matrix);

  assertEquals(result.length, 1);
  assertEquals(result[0].value, 0);
  assertEquals(result[0].topLeft, [0, 0]);
  assertEquals(result[0].bottomRight, [1, 1]);
});

Deno.test("findAllRectangles - result structure", () => {
  const matrix = [[1]];
  const result = findAllRectangles(matrix);

  // Verify the structure of returned rectangles
  assertEquals(result.length, 1);

  const rect = result[0];
  assert(Array.isArray(rect.topLeft), "topLeft should be an array");
  assert(Array.isArray(rect.bottomRight), "bottomRight should be an array");
  assert(typeof rect.value === "number", "value should be a number");
  assertEquals(
    rect.topLeft.length,
    2,
    "topLeft should have 2 elements [row, col]",
  );
  assertEquals(
    rect.bottomRight.length,
    2,
    "bottomRight should have 2 elements [row, col]",
  );
});

Deno.test("findAllRectangles - real case example", () => {
  const matrix = [
    [0, 1, 1, 1, 1, 1, 2],
    [3, 4, 4, 4, 4, 4, 5],
    [3, 4, 4, 4, 4, 4, 5],
    [3, 4, 4, 4, 4, 4, 5],
    [3, 4, 4, 4, 4, 4, 5],
    [3, 4, 4, 4, 4, 4, 5],
    [3, 4, 4, 4, 4, 4, 5],
    [3, 4, 4, 4, 4, 4, 5],
    [6, 7, 7, 7, 7, 7, 8],
  ];
  const result = findAllRectangles(matrix);
  console.log("Rectangles:", result);

  // Should find one rectangle for each distinct region:
  // - 0: single cell [0,0]
  // - 1: horizontal rectangle [0,1] to [0,5] (1x5)
  // - 2: single cell [0,6]
  // - 3: vertical rectangle [1,0] to [7,0] (7x1)
  // - 4: large rectangle [1,1] to [7,5] (7x5)
  // - 5: vertical rectangle [1,6] to [7,6] (7x1)
  // - 6: single cell [8,0]
  // - 7: horizontal rectangle [8,1] to [8,5] (1x5)
  // - 8: single cell [8,6]
  assertEquals(result.length, 9);

  // Verify all rectangles have valid coordinates
  result.forEach((rect) => {
    assert(rect.topLeft[0] >= 0, "Top row should be non-negative");
    assert(rect.topLeft[1] >= 0, "Left column should be non-negative");
    assert(
      rect.bottomRight[0] >= rect.topLeft[0],
      "Bottom row should be >= top row",
    );
    assert(
      rect.bottomRight[1] >= rect.topLeft[1],
      "Right column should be >= left column",
    );
    assert(
      rect.bottomRight[0] < matrix.length,
      "Bottom row should be within matrix bounds",
    );
    assert(
      rect.bottomRight[1] < matrix[0].length,
      "Right column should be within matrix bounds",
    );
  });

  // Check that we found rectangles for all unique values
  const foundValues = result.map((r) => r.value).sort();
  const expectedValues = [0, 1, 2, 3, 4, 5, 6, 7, 8];
  assertEquals(foundValues, expectedValues);

  // Find and verify the largest rectangle (should be the 4's - 7x5 = 35 area)
  const largestRect = result[0]; // Should be sorted by area descending
  assertEquals(largestRect.value, 4);

  const area = (largestRect.bottomRight[0] - largestRect.topLeft[0] + 1) *
    (largestRect.bottomRight[1] - largestRect.topLeft[1] + 1);
  assertEquals(area, 35); // 7 rows × 5 columns = 35
});

// Tests for analyzeRangeOverlaps function
Deno.test("analyzeRangeOverlaps - empty array", () => {
  const result = analyzeRangeOverlaps([]);
  assertEquals(result.ranges, []);
});

Deno.test("analyzeRangeOverlaps - single range", () => {
  const result = analyzeRangeOverlaps(["A1:C3"]);
  assertEquals(result.ranges, ["A1:C3"]);
});

Deno.test("analyzeRangeOverlaps - no overlaps", () => {
  const result = analyzeRangeOverlaps(["A1:B2", "D1:E2", "A4:B5"]);

  // Should return all original ranges since no overlaps
  assert(result.ranges.includes("A1:B2"));
  assert(result.ranges.includes("D1:E2"));
  assert(result.ranges.includes("A4:B5"));
  assertEquals(result.ranges.length, 3);
});

Deno.test("analyzeRangeOverlaps - simple overlap", () => {
  const result = analyzeRangeOverlaps(["A1:C3", "B2:D4"]);

  // Optimization may reduce the number of ranges through merging
  assert(result.ranges.length >= 1);

  // Check that all ranges are valid
  result.ranges.forEach((range) => {
    assert(
      range.match(/^[A-Z]+\d+(?::[A-Z]+\d+)?$/),
      `Invalid range format: ${range}`,
    );
  });

  // Validate cell preservation
  const inputCells = getCellsFromRanges(["A1:C3", "B2:D4"]);
  const resultCells = getCellsFromRanges(result.ranges);
  assert(cellSetsEqual(inputCells, resultCells), "Cell preservation failed");
});

Deno.test("analyzeRangeOverlaps - user example: A1:B10 with A2:H8", () => {
  const result = analyzeRangeOverlaps(["A1:B10", "A2:H8"]);

  // Should include all cells from both ranges
  const inputCells = getCellsFromRanges(["A1:B10", "A2:H8"]);
  const resultCells = getCellsFromRanges(result.ranges);
  assert(cellSetsEqual(inputCells, resultCells), "Cell preservation failed");

  // Should have reasonable number of ranges (optimization may reduce count)
  assert(result.ranges.length >= 1, "Should have at least one range");

  console.log("User example result:", result.ranges);
});

Deno.test("analyzeRangeOverlaps - multiple overlaps", () => {
  const result = analyzeRangeOverlaps(["A1:D4", "B2:E5", "C3:F6"]);

  // Should have multiple ranges
  assert(result.ranges.length > 0);

  // All ranges should be valid
  result.ranges.forEach((range) => {
    assert(
      range.match(/^[A-Z]+\d+(?::[A-Z]+\d+)?$/),
      `Invalid range format: ${range}`,
    );
  });

  // No duplicate ranges
  const uniqueRanges = Array.from(new Set(result.ranges));
  assertEquals(result.ranges.length, uniqueRanges.length);
});

Deno.test("analyzeRangeOverlaps - complete overlap", () => {
  const result = analyzeRangeOverlaps(["A1:C3", "B2:B2"]);

  // B2:B2 is completely inside A1:C3, optimization should merge them
  assert(result.ranges.length >= 1);

  // Should preserve all cells
  const inputCells = getCellsFromRanges(["A1:C3", "B2:B2"]);
  const resultCells = getCellsFromRanges(result.ranges);
  assert(cellSetsEqual(inputCells, resultCells), "Cell preservation failed");
});

Deno.test("analyzeRangeOverlaps - identical ranges", () => {
  const result = analyzeRangeOverlaps(["A1:C3", "A1:C3"]);

  // Should have just the overlapping range
  assertEquals(result.ranges.length, 1);
  assertEquals(result.ranges[0], "A1:C3");
});

Deno.test("analyzeRangeOverlaps - adjacent ranges", () => {
  const result = analyzeRangeOverlaps(["A1:B2", "C1:D2"]);

  // Adjacent ranges can be merged by optimization
  console.log("Adjacent ranges result:", result.ranges);

  // Should preserve all cells
  const inputCells = getCellsFromRanges(["A1:B2", "C1:D2"]);
  const resultCells = getCellsFromRanges(result.ranges);
  assert(cellSetsEqual(inputCells, resultCells), "Cell preservation failed");
});

Deno.test("analyzeRangeOverlaps - single cell ranges", () => {
  const result = analyzeRangeOverlaps(["A1", "A1", "B1"]);

  // Optimization can merge adjacent single cells
  assert(result.ranges.length >= 1);

  // Should preserve all unique cells
  const inputCells = getCellsFromRanges(["A1", "A1", "B1"]);
  const resultCells = getCellsFromRanges(result.ranges);
  assert(cellSetsEqual(inputCells, resultCells), "Cell preservation failed");

  console.log("Single cell ranges result:", result.ranges);
});

Deno.test("analyzeRangeOverlaps - complex scenario", () => {
  const ranges = ["A1:C5", "B3:E7", "D1:F3", "E5:G8"];
  const result = analyzeRangeOverlaps(ranges);

  // Should handle complex overlaps
  assert(result.ranges.length > 0);

  // Verify all returned ranges are valid Excel ranges
  result.ranges.forEach((range) => {
    assert(
      range.match(/^[A-Z]+\d+(?::[A-Z]+\d+)?$/),
      `Invalid range: ${range}`,
    );
  });

  console.log("Complex scenario result:", result.ranges);
});

Deno.test("analyzeRangeOverlaps - L-shaped overlap", () => {
  const result = analyzeRangeOverlaps(["A1:C3", "C1:E3"]);

  // Optimization can merge into one large rectangle
  assert(result.ranges.length >= 1);

  // Should preserve all cells
  const inputCells = getCellsFromRanges(["A1:C3", "C1:E3"]);
  const resultCells = getCellsFromRanges(result.ranges);
  assert(cellSetsEqual(inputCells, resultCells), "Cell preservation failed");

  console.log("L-shaped result:", result.ranges);
});

Deno.test("analyzeRangeOverlaps - range validation", () => {
  const result = analyzeRangeOverlaps(["A1:B2", "A3:B4", "C1:D4"]);

  // All results should be valid range strings
  result.ranges.forEach((range) => {
    // Should match Excel range pattern
    assert(
      range.match(/^[A-Z]+\d+(?::[A-Z]+\d+)?$/),
      `Invalid range: ${range}`,
    );

    // If it's a range (contains :), start should be <= end
    if (range.includes(":")) {
      const [start, end] = range.split(":");
      const startMatch = start.match(/([A-Z]+)(\d+)/);
      const endMatch = end.match(/([A-Z]+)(\d+)/);

      if (startMatch && endMatch) {
        const startCol = startMatch[1];
        const startRow = parseInt(startMatch[2]);
        const endCol = endMatch[1];
        const endRow = parseInt(endMatch[2]);

        assert(startRow <= endRow, `Invalid row order in range ${range}`);
        assert(startCol <= endCol, `Invalid column order in range ${range}`);
      }
    }
  });
});

// Comprehensive validation tests
Deno.test("analyzeRangeOverlaps - comprehensive validation: no overlaps in result", () => {
  const testCases = [
    ["A1:B10", "A2:H8"],
    ["A1:C5", "B3:E7", "D1:F3"],
    ["A1:C3", "C1:E3"],
    ["A1:E5", "B2:D4"],
    ["A1", "A1", "B1"],
    ["A1:B2", "C3:D4", "E5:F6"],
  ];

  for (const inputRanges of testCases) {
    const result = analyzeRangeOverlaps(inputRanges);

    // Validate no overlaps in result
    assert(
      validateNoOverlaps(result.ranges),
      `Result ranges should not overlap for input: ${
        JSON.stringify(inputRanges)
      }`,
    );
  }
});

Deno.test("analyzeRangeOverlaps - comprehensive validation: reasonable result count", () => {
  const testCases = [
    { input: ["A1:B10", "A2:H8"], maxExpected: 10 }, // Should be much less than this
    { input: ["A1:C5", "B3:E7", "D1:F3"], maxExpected: 20 },
    { input: ["A1:C3", "C1:E3"], maxExpected: 5 },
    { input: ["A1:E5", "B2:D4"], maxExpected: 10 },
    { input: ["A1", "A1", "B1"], maxExpected: 3 }, // Optimization can merge to 1
    { input: ["A1:B2", "C3:D4", "E5:F6"], maxExpected: 3 }, // No overlaps, same count or optimized
    { input: ["A1:A1", "B1:B1", "C1:C1", "D1:D1"], maxExpected: 4 }, // Optimization can merge to 1
  ];

  for (const testCase of testCases) {
    const result = analyzeRangeOverlaps(testCase.input);

    // Result should be reasonable (not explosive growth)
    assert(
      result.ranges.length <= testCase.maxExpected,
      `Result count (${result.ranges.length}) should be reasonable (≤ ${testCase.maxExpected}) for: ${
        JSON.stringify(testCase.input)
      }`,
    );

    // Cell preservation
    const inputCells = getCellsFromRanges(testCase.input);
    const resultCells = getCellsFromRanges(result.ranges);
    assert(cellSetsEqual(inputCells, resultCells), "Cell preservation failed");

    console.log(
      `✓ Reasonable count for ${
        JSON.stringify(testCase.input)
      }: ${testCase.input.length} → ${result.ranges.length}`,
    );
  }
});

Deno.test("analyzeRangeOverlaps - comprehensive validation: cell preservation", () => {
  const testCases = [
    ["A1:B2", "A2:B3"], // Simple overlap
    ["A1:C3", "B2:D4"], // Partial overlap
    ["A1:B10", "A2:H8"], // User's example
    ["A1:E5", "B2:D4"], // Complete containment
    ["A1", "A1", "B1"], // Duplicate single cells
    ["A1:B2", "C1:D2"], // No overlap
    ["A1:C5", "B3:E7", "D1:F3"], // Multiple overlaps
  ];

  for (const inputRanges of testCases) {
    const result = analyzeRangeOverlaps(inputRanges);

    // Get all cells from input and result
    const inputCells = getCellsFromRanges(inputRanges);
    const resultCells = getCellsFromRanges(result.ranges);

    // All input cells should be in result
    for (const cell of inputCells) {
      assert(
        resultCells.has(cell),
        `Input cell ${cell} missing from result for: ${
          JSON.stringify(inputRanges)
        }`,
      );
    }

    // No extra cells in result
    for (const cell of resultCells) {
      assert(
        inputCells.has(cell),
        `Extra cell ${cell} found in result for: ${
          JSON.stringify(inputRanges)
        }`,
      );
    }

    // Sets should be exactly equal
    assert(
      cellSetsEqual(inputCells, resultCells),
      `Cell sets should be equal for: ${JSON.stringify(inputRanges)}`,
    );

    console.log(
      `✓ Validated cell preservation for: ${JSON.stringify(inputRanges)}`,
    );
  }
});

Deno.test("analyzeRangeOverlaps - comprehensive validation: all properties combined", () => {
  const complexTestCases = [
    // Complex multi-range scenarios
    ["A1:D5", "B2:F6", "C3:G7", "E1:H4"],
    ["A1:B10", "C5:D15", "B8:E12"],
    ["A1:A5", "B1:B5", "C1:C5", "A3:C3"], // Grid pattern
    ["A1:E1", "A2:E2", "A3:E3", "C1:C3"], // Row and column intersections
  ];

  for (const inputRanges of complexTestCases) {
    const result = analyzeRangeOverlaps(inputRanges);

    // 1. No overlaps in result
    assert(
      validateNoOverlaps(result.ranges),
      `❌ Overlaps found in result for: ${JSON.stringify(inputRanges)}`,
    );

    // 2. Cell preservation
    const inputCells = getCellsFromRanges(inputRanges);
    const resultCells = getCellsFromRanges(result.ranges);

    assert(
      cellSetsEqual(inputCells, resultCells),
      `❌ Cell preservation failed for: ${JSON.stringify(inputRanges)}`,
    );

    // 3. Reasonable result count (should not explode)
    const totalInputCells = inputCells.size;
    assert(
      result.ranges.length <= totalInputCells, // Very generous upper bound
      `❌ Result count too high for: ${JSON.stringify(inputRanges)}`,
    );

    console.log(`✓ All validations passed for: ${JSON.stringify(inputRanges)}`);
    console.log(
      `  Input ranges: ${inputRanges.length}, Result ranges: ${result.ranges.length}`,
    );
    console.log(`  Total cells: ${inputCells.size}`);
  }
});

Deno.test("analyzeRangeOverlaps - edge case validation", () => {
  // Edge cases that could break the algorithm
  const edgeCases = [
    [], // Empty input
    ["A1"], // Single range
    ["A1", "A1"], // Identical ranges
    ["A1:A1"], // Single cell as range
    ["A1:B1", "B1:C1"], // Adjacent ranges sharing one cell
    ["A1:Z1", "A2:Z2", "A1:A26", "Z1:Z26"], // Large ranges with small overlaps
  ];

  for (const inputRanges of edgeCases) {
    const result = analyzeRangeOverlaps(inputRanges);

    if (inputRanges.length === 0) {
      assertEquals(result.ranges.length, 0);
      continue;
    }

    // Apply all validations
    assert(
      validateNoOverlaps(result.ranges),
      `Edge case overlap validation failed: ${JSON.stringify(inputRanges)}`,
    );

    const inputCells = getCellsFromRanges(inputRanges);
    const resultCells = getCellsFromRanges(result.ranges);

    assert(
      cellSetsEqual(inputCells, resultCells),
      `Edge case cell preservation failed: ${JSON.stringify(inputRanges)}`,
    );

    // For edge cases, result should be reasonable
    assert(
      result.ranges.length <= inputCells.size,
      `Edge case result count too high: ${JSON.stringify(inputRanges)}`,
    );

    console.log(
      `✓ Edge case validated: ${
        JSON.stringify(inputRanges)
      } → ${result.ranges.length} ranges`,
    );
  }
});

Deno.test("analyzeRangeOverlaps - edge case validation: Complex case", () => {
  const complexCase = [
    "C9:N10",
    "L5:M14",
    "K11",
    "F13:P13",
    "H6:I13",
    "F6:J12",
    "G8:J11",
  ];
  const result = analyzeRangeOverlaps(complexCase);
  console.log("Result ranges:", result.ranges);
});

Deno.test("analyzeRangeOverlaps - optimization: merges adjacent ranges", () => {
  // Test cases where adjacent ranges should be merged
  const testCases = [
    {
      input: ["A1:A2", "A3:A4"],
      expectedMerged: ["A1:A4"],
      description: "Vertical adjacent ranges",
    },
    {
      input: ["A1:B1", "C1:D1"],
      expectedMerged: ["A1:D1"],
      description: "Horizontal adjacent ranges",
    },
    {
      input: ["A1:B2", "A3:B4", "A5:B6"],
      expectedMerged: ["A1:B6"],
      description: "Multiple vertical adjacent ranges",
    },
    {
      input: ["A1:A3", "B1:B3", "C1:C3"],
      expectedMerged: ["A1:C3"],
      description: "Multiple horizontal adjacent ranges",
    },
    {
      input: ["A1:B1", "A2:B2", "C1:D1", "C2:D2"],
      expectedMerged: ["A1:D2"],
      description: "All ranges merge into one large rectangle",
    },
  ];

  for (const testCase of testCases) {
    const result = analyzeRangeOverlaps(testCase.input);

    // Check that expected merged ranges are present
    for (const expectedRange of testCase.expectedMerged) {
      assert(
        result.ranges.includes(expectedRange),
        `Expected merged range ${expectedRange} not found in result: ${
          JSON.stringify(result.ranges)
        } for ${testCase.description}`,
      );
    }

    // Should have fewer or equal ranges after optimization
    assert(
      result.ranges.length <= testCase.input.length,
      `Optimization should not increase range count for ${testCase.description}`,
    );

    // Cell preservation should still work
    const inputCells = getCellsFromRanges(testCase.input);
    const resultCells = getCellsFromRanges(result.ranges);
    assert(
      cellSetsEqual(inputCells, resultCells),
      `Cell preservation failed after optimization for ${testCase.description}`,
    );

    console.log(
      `✓ ${testCase.description}: ${JSON.stringify(testCase.input)} → ${
        JSON.stringify(result.ranges)
      }`,
    );
  }
});

Deno.test("analyzeRangeOverlaps - optimization: prefers larger ranges", () => {
  // Test that optimization results in fewer, larger ranges
  const testCases = [
    {
      input: ["A1:A1", "A2:A2", "A3:A3", "A4:A4", "A5:A5"],
      maxExpectedCount: 1, // Should merge into A1:A5
      description: "Chain of single cells",
    },
    {
      input: ["A1:B1", "A2:B2", "A3:B3"],
      maxExpectedCount: 1, // Should merge into A1:B3
      description: "Chain of 2-column ranges",
    },
    {
      input: ["A1:C1", "D1:F1", "G1:I1"],
      maxExpectedCount: 1, // Should merge into A1:I1
      description: "Chain of 3-column ranges",
    },
  ];

  for (const testCase of testCases) {
    const result = analyzeRangeOverlaps(testCase.input);

    // Should have fewer ranges after optimization
    assert(
      result.ranges.length <= testCase.maxExpectedCount,
      `Expected at most ${testCase.maxExpectedCount} ranges, got ${result.ranges.length} for ${testCase.description}`,
    );

    // Cell preservation
    const inputCells = getCellsFromRanges(testCase.input);
    const resultCells = getCellsFromRanges(result.ranges);
    assert(
      cellSetsEqual(inputCells, resultCells),
      `Cell preservation failed for ${testCase.description}`,
    );

    // No overlaps
    assert(
      validateNoOverlaps(result.ranges),
      `Overlaps found after optimization for ${testCase.description}`,
    );

    console.log(
      `✓ ${testCase.description}: ${testCase.input.length} → ${result.ranges.length} ranges`,
    );
    console.log(`  Result: ${JSON.stringify(result.ranges)}`);
  }
});

Deno.test("analyzeRangeOverlaps - optimization: handles complex scenarios", () => {
  // Test optimization with overlapping ranges that create merge opportunities
  const result1 = analyzeRangeOverlaps(["A1:A5", "A6:A10"]);
  console.log("Adjacent ranges merge test:", result1.ranges);

  const result2 = analyzeRangeOverlaps(["A1:B3", "A4:B6", "C1:D3", "C4:D6"]);
  console.log("Multiple adjacent groups:", result2.ranges);

  const result3 = analyzeRangeOverlaps(["A1:C1", "A2:C2", "D1:F2"]);
  console.log("Partial adjacency:", result3.ranges);

  // Verify all maintain our core properties
  const testCases = [
    ["A1:A5", "A6:A10"],
    ["A1:B3", "A4:B6", "C1:D3", "C4:D6"],
    ["A1:C1", "A2:C2", "D1:F2"],
  ];

  for (const inputRanges of testCases) {
    const result = analyzeRangeOverlaps(inputRanges);

    // Core validations still apply
    assert(validateNoOverlaps(result.ranges));

    const inputCells = getCellsFromRanges(inputRanges);
    const resultCells = getCellsFromRanges(result.ranges);
    assert(cellSetsEqual(inputCells, resultCells));

    console.log(
      `✓ Complex scenario validated: ${
        JSON.stringify(inputRanges)
      } → ${result.ranges.length} ranges`,
    );
  }
});

// Tests for expandRangeByIndex function
Deno.test("expandRangeByIndex - basic functionality", () => {
  // Test expanding a simple range
  const result = expandRangeByIndex("B2:C3", 2);
  assertEquals(result, "A1:E5");

  // The original range B2:C3 (2x2) should expand by 2 cells in each direction
  // making it A1:E5 (5x5, but clamped to A1 on the left/top)
});

Deno.test("expandRangeByIndex - single cell expansion", () => {
  // Test expanding a single cell
  const result = expandRangeByIndex("B2:B2", 2);
  assertEquals(result, "A1:D4");

  // Single cell B2 should expand by 2 in each direction (clamped to A1)
});

Deno.test("expandRangeByIndex - zero expansion", () => {
  // Test with no expansion
  const result = expandRangeByIndex("B2:C3", 0);
  assertEquals(result, "B2:C3");

  // Should return the original range
});

Deno.test("expandRangeByIndex - odd expansion values", () => {
  // Test with odd expansion values (no longer dividing by 2)
  const result = expandRangeByIndex("C3:D4", 3);
  // Expand by 3 in each direction
  assertEquals(result, "A1:G7");
});

Deno.test("expandRangeByIndex - large expansion", () => {
  // Test with larger expansion
  const result = expandRangeByIndex("D4:F6", 4);
  // Expand by 4 in each direction
  assertEquals(result, "A1:J10");
});

// Edge cases with top left corner (A1)
Deno.test("expandRangeByIndex - edge case: A1 single cell", () => {
  // Test expanding A1 (top-left corner)
  const result = expandRangeByIndex("A1:A1", 2);
  assertEquals(result, "A1:C3");

  // Cannot expand beyond A1 (column A, row 1), so it stays at A1
});

Deno.test("expandRangeByIndex - edge case: A1 range", () => {
  // Test expanding range starting at A1
  const result = expandRangeByIndex("A1:B2", 2);
  assertEquals(result, "A1:D4");

  // Cannot expand beyond A1, but can expand right and down
});

Deno.test("expandRangeByIndex - edge case: first column", () => {
  // Test expanding range in first column
  const result = expandRangeByIndex("A3:A5", 2);
  assertEquals(result, "A1:C7");

  // Cannot expand left of column A
});

Deno.test("expandRangeByIndex - edge case: first row", () => {
  // Test expanding range in first row
  const result = expandRangeByIndex("C1:E1", 2);
  assertEquals(result, "A1:G3");

  // Cannot expand above row 1
});

Deno.test("expandRangeByIndex - edge case: corner range A1:B1", () => {
  // Test expanding range that includes top-left corner
  const result = expandRangeByIndex("A1:B1", 4);
  assertEquals(result, "A1:F5");

  // Can only expand right and down from A1
});

Deno.test("expandRangeByIndex - edge case: corner range A1:A2", () => {
  // Test expanding range that includes top-left corner vertically
  const result = expandRangeByIndex("A1:A2", 2);
  assertEquals(result, "A1:C4");

  // Can only expand right and down from A1
});

// Tests for expandRange function (supporting function)
Deno.test("expandRange - basic functionality", () => {
  // Test expanding with different row and column values
  const result = expandRange("C3:D4", 2, 4);
  // Expand by 2 rows and 4 columns in each direction
  assertEquals(result, "A1:H6");
});

Deno.test("expandRange - asymmetric expansion", () => {
  // Test expanding with different values for rows and columns
  const result = expandRange("C3:D4", 4, 2);
  // Expand by 4 rows and 2 columns in each direction
  assertEquals(result, "A1:F8");
});

Deno.test("expandRange - edge case with A1", () => {
  // Test expanding range that would go beyond A1
  const result = expandRange("A1:B2", 2, 2);
  assertEquals(result, "A1:D4");

  // Cannot expand beyond A1
});

Deno.test("expandRange - large asymmetric expansion", () => {
  // Test with very different row and column expansion
  const result = expandRange("E5:F6", 0, 6);
  // No row expansion, 6 column expansion each way
  assertEquals(result, "A5:L6");
});

// Comprehensive edge case testing
Deno.test("expandRangeByIndex - comprehensive edge cases", () => {
  const testCases = [
    {
      input: { range: "A1:A1", expand: 1 },
      expected: "A1:B2",
      description: "Single A1 cell with minimal expansion",
    },
    {
      input: { range: "A1:A1", expand: 4 },
      expected: "A1:E5",
      description: "Single A1 cell with larger expansion",
    },
    {
      input: { range: "A2:A2", expand: 2 },
      expected: "A1:C4",
      description: "Single cell in first column, can expand up",
    },
    {
      input: { range: "B1:B1", expand: 2 },
      expected: "A1:D3",
      description: "Single cell in first row, can expand left",
    },
    {
      input: { range: "A1:C1", expand: 2 },
      expected: "A1:E3",
      description: "Horizontal range starting at A1",
    },
    {
      input: { range: "A1:A3", expand: 2 },
      expected: "A1:C5",
      description: "Vertical range starting at A1",
    },
    {
      input: { range: "A1:B2", expand: 6 },
      expected: "A1:H8",
      description: "Rectangle starting at A1 with large expansion",
    },
  ];

  for (const testCase of testCases) {
    const result = expandRangeByIndex(
      testCase.input.range,
      testCase.input.expand,
    );
    assertEquals(
      result,
      testCase.expected,
      `Failed for ${testCase.description}: ${testCase.input.range} expanded by ${testCase.input.expand}`,
    );

    console.log(
      `✓ ${testCase.description}: ${testCase.input.range} + ${testCase.input.expand} → ${result}`,
    );
  }
});

// Test boundary conditions
Deno.test("expandRangeByIndex - boundary validation", () => {
  const testCases = [
    "A1:A1",
    "A1:B1",
    "A1:A2",
    "A1:B2",
    "A2:A3",
    "B1:C1",
    "A1:Z1",
    "A1:A26",
  ];

  for (const range of testCases) {
    for (const expansion of [0, 1, 2, 4, 8]) {
      const result = expandRangeByIndex(range, expansion);

      // Validate that result is a valid range
      assert(
        result.match(/^[A-Z]+\d+:[A-Z]+\d+$/),
        `Invalid range format: ${result} (from ${range} + ${expansion})`,
      );

      // Parse the result to ensure it's valid
      const parsed = parseRange(result);

      // Ensure start is not before A1
      assert(parsed.startColumn >= "A", "Start column should not be before A");
      assert(parsed.startRow >= 1, "Start row should not be before 1");

      // Ensure start <= end
      assert(
        columnToNumber(parsed.startColumn) <= columnToNumber(parsed.endColumn),
        "Start column should be <= end column",
      );
      assert(
        parsed.startRow <= parsed.endRow,
        "Start row should be <= end row",
      );

      console.log(`✓ Boundary test: ${range} + ${expansion} → ${result}`);
    }
  }
});

// Test that expandRangeByIndex delegates correctly to expandRange
Deno.test("expandRangeByIndex - delegates to expandRange correctly", () => {
  const testRange = "C3:E5";
  const expandIndex = 4;

  // Call both functions
  const byIndexResult = expandRangeByIndex(testRange, expandIndex);
  const directResult = expandRange(testRange, expandIndex, expandIndex);

  // Should produce identical results
  assertEquals(
    byIndexResult,
    directResult,
    "expandRangeByIndex should delegate to expandRange with same value for rows and columns",
  );
});

// Performance and stress testing
Deno.test("expandRangeByIndex - stress test", () => {
  const ranges = [
    "A1:A1",
    "A1:B2",
    "C3:F6",
    "Z1:Z1",
    "A26:B27",
    "M13:P16",
    "AA1:AB2",
    "A1:A100",
    "A1:Z1",
  ];

  const expansions = [0, 1, 2, 5, 10, 20, 50];

  for (const range of ranges) {
    for (const expansion of expansions) {
      const result = expandRangeByIndex(range, expansion);

      // Validate result format
      assert(
        result.match(/^[A-Z]+\d+:[A-Z]+\d+$/),
        `Invalid result format: ${result}`,
      );

      // Ensure we can parse the result
      const parsed = parseRange(result);

      // Basic validity checks
      assert(parsed.startRow >= 1, "Start row should be >= 1");
      assert(
        columnToNumber(parsed.startColumn) >= 1,
        "Start column should be >= A",
      );
      assert(
        parsed.startRow <= parsed.endRow,
        "Start row should be <= end row",
      );
      assert(
        columnToNumber(parsed.startColumn) <= columnToNumber(parsed.endColumn),
        "Start column should be <= end column",
      );
    }
  }

  console.log("✓ Stress test completed successfully");
});
