/**
 * Utility functions for Excel range manipulation
 */

export interface ParsedRange {
  startColumn: string;
  startRow: number;
  endColumn: string;
  endRow: number;
}

/**
 * Rectangle structure for geometric operations
 */
export interface RangeRectangle {
  left: number;
  top: number;
  right: number;
  bottom: number;
  originalRange: string;
}

/**
 * Result of range overlap analysis
 */
export interface RangeOverlapResult {
  ranges: string[];
}

/**
 * Convert column number to Excel column letters (1 = A, 26 = Z, 27 = AA, etc.)
 */
export function numberToColumn(num: number): string {
  let result = "";
  while (num > 0) {
    num--; // Make it 0-based
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

/**
 * Convert Excel column letters to number (A = 1, Z = 26, AA = 27, etc.)
 */
export function columnToNumber(column: string): number {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Parse an Excel range string like "A1:C10" or single cell "A1" into its components
 */
export function parseRange(range: string): ParsedRange {
  // Handle single cell reference (e.g., "A1")
  const singleCellMatch = range.match(/^([A-Z]+)(\d+)$/);
  if (singleCellMatch) {
    return {
      startColumn: singleCellMatch[1],
      startRow: parseInt(singleCellMatch[2], 10),
      endColumn: singleCellMatch[1],
      endRow: parseInt(singleCellMatch[2], 10),
    };
  }

  // Handle range reference (e.g., "A1:C10")
  const rangeMatch = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (rangeMatch) {
    return {
      startColumn: rangeMatch[1],
      startRow: parseInt(rangeMatch[2], 10),
      endColumn: rangeMatch[3],
      endRow: parseInt(rangeMatch[4], 10),
    };
  }

  throw new Error(
    `Invalid range format: ${range}. Expected format like A1 or A1:C10`,
  );
}

/**
 * Convert parsed range back to Excel range string
 */
export function rangeToParsedString(parsed: ParsedRange): string {
  return `${parsed.startColumn}${parsed.startRow}:${parsed.endColumn}${parsed.endRow}`;
}

/**
 * Expand an Excel range by the specified number of rows and columns in EACH direction
 * @param range - Original range string (e.g., "A1:C10")
 * @param increaseRows - Number of rows to add in each direction (up and down)
 * @param increaseColumns - Number of columns to add in each direction (left and right)
 * @returns Expanded range string
 */
export function expandRange(
  range: string,
  increaseRows: number,
  increaseColumns: number,
): string {
  const parsed = parseRange(range);

  // Use the full expansion values for each direction
  const rowsExpansion = increaseRows;
  const columnsExpansion = increaseColumns;

  // Convert columns to numbers for calculation
  const startColumnNumber = columnToNumber(parsed.startColumn);
  const endColumnNumber = columnToNumber(parsed.endColumn);

  // Expand start coordinates (move backward/up)
  const newStartColumnNumber = Math.max(
    1,
    startColumnNumber - columnsExpansion,
  );
  const newStartRow = Math.max(1, parsed.startRow - rowsExpansion);

  // Expand end coordinates (move forward/down)
  const newEndColumnNumber = endColumnNumber + columnsExpansion;
  const newEndRow = parsed.endRow + rowsExpansion;

  // Convert back to column letters
  const newStartColumn = numberToColumn(newStartColumnNumber);
  const newEndColumn = numberToColumn(newEndColumnNumber);

  return `${newStartColumn}${newStartRow}:${newEndColumn}${newEndRow}`;
}

/**
 * Expand an Excel range by a single increase index for both rows and columns in EACH direction
 * @param range - Original range string (e.g., "A1:C10")
 * @param increaseIndex - Number to add in each direction for both rows and columns
 * @returns Expanded range string
 */
export function expandRangeByIndex(
  range: string,
  increaseIndex: number,
): string {
  return expandRange(range, increaseIndex, increaseIndex);
}

export type Rectangle = {
  topLeft: [number, number];
  bottomRight: [number, number];
  value: number;
};

export function findAllRectangles(matrix: number[][]): Rectangle[] {
  const rows = matrix.length;
  if (rows === 0) return [];
  const cols = matrix[0].length;
  if (cols === 0) return [];

  const visited: boolean[][] = Array.from(
    { length: rows },
    () => Array(cols).fill(false),
  );
  const rectangles: Rectangle[] = [];

  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      if (visited[i][j]) continue;

      const value = matrix[i][j];

      // Find the maximum width for the current row
      let width = 0;
      while (
        j + width < cols && matrix[i][j + width] === value &&
        !visited[i][j + width]
      ) {
        width++;
      }

      // Find the maximum height that maintains the width
      let height = 1;
      for (let row = i + 1; row < rows; row++) {
        let canExtend = true;
        for (let col = j; col < j + width; col++) {
          if (visited[row][col] || matrix[row][col] !== value) {
            canExtend = false;
            break;
          }
        }
        if (canExtend) {
          height++;
        } else {
          break;
        }
      }

      // Mark all cells in the rectangle as visited
      for (let r = i; r < i + height; r++) {
        for (let c = j; c < j + width; c++) {
          visited[r][c] = true;
        }
      }

      // Only add rectangles with positive width and height
      if (width > 0 && height > 0) {
        rectangles.push({
          topLeft: [i, j],
          bottomRight: [i + height - 1, j + width - 1],
          value,
        });
      }
    }
  }

  // Sort by area descending (largest rectangles first)
  rectangles.sort((a, b) => {
    const areaA = (a.bottomRight[0] - a.topLeft[0] + 1) *
      (a.bottomRight[1] - a.topLeft[1] + 1);
    const areaB = (b.bottomRight[0] - b.topLeft[0] + 1) *
      (b.bottomRight[1] - b.topLeft[1] + 1);
    return areaB - areaA;
  });

  return rectangles;
}

/**
 * Get the top-left cell from a collection of cell data
 */
export function getTopLeftCell(cellData: Record<string, any>): string {
  const cellAddresses = Object.keys(cellData);
  const { minRow, minCol } = findRangeBounds(cellAddresses);
  return `${numberToColumn(minCol)}${minRow}`;
}

/**
 * Find the bounds of a range from cell addresses
 */
export function findRangeBounds(
  cellAddresses: string[],
): { minRow: number; maxRow: number; minCol: number; maxCol: number } {
  let minRow = Infinity, maxRow = -Infinity;
  let minCol = Infinity, maxCol = -Infinity;

  for (const address of cellAddresses) {
    const match = address.match(/([A-Z]+)(\d+)/);
    if (match) {
      const col = columnToNumber(match[1]);
      const row = parseInt(match[2]);

      minRow = Math.min(minRow, row);
      maxRow = Math.max(maxRow, row);
      minCol = Math.min(minCol, col);
      maxCol = Math.max(maxCol, col);
    }
  }

  return { minRow, maxRow, minCol, maxCol };
}

/**
 * Convert rectangle coordinates to Excel ranges using a top-left reference cell
 */
export function convertRectanglesToExcelRanges(
  rectangles: Rectangle[],
  topLeftCell: string,
): string[] {
  // Parse the top-left cell to get base coordinates
  const match = topLeftCell.match(/([A-Z]+)(\d+)/);
  if (!match) return [];

  const baseCol = columnToNumber(match[1]);
  const baseRow = parseInt(match[2]);

  return rectangles.map((rect) => {
    // Convert matrix indices to Excel coordinates
    const startCol = baseCol + rect.topLeft[1];
    const startRow = baseRow + rect.topLeft[0];
    const endCol = baseCol + rect.bottomRight[1];
    const endRow = baseRow + rect.bottomRight[0];

    const startCell = `${numberToColumn(startCol)}${startRow}`;
    const endCell = `${numberToColumn(endCol)}${endRow}`;

    // Return single cell or range
    return startCell === endCell ? startCell : `${startCell}:${endCell}`;
  });
}

/**
 * Convert cell address to coordinates
 */
function cellToCoords(cell: string): [number, number] {
  const match = cell.match(/([A-Z]+)(\d+)/);
  if (!match) throw new Error(`Invalid cell: ${cell}`);
  return [columnToNumber(match[1]), parseInt(match[2])];
}

/**
 * Convert coordinates to cell address
 */
function coordsToCell(col: number, row: number): string {
  return `${numberToColumn(col)}${row}`;
}

/**
 * Get all cells covered by a rectangle
 */
function getRectangleCells(rect: RangeRectangle): Set<string> {
  const cells = new Set<string>();
  for (let row = rect.top; row <= rect.bottom; row++) {
    for (let col = rect.left; col <= rect.right; col++) {
      cells.add(coordsToCell(col, row));
    }
  }
  return cells;
}

/**
 * Find the largest rectangle that can be formed from a set of cells using backtracking
 */
function findLargestRectangle(cells: Set<string>): RangeRectangle | null {
  if (cells.size === 0) return null;

  // Convert cells to coordinates
  const cellArray = Array.from(cells);
  const coords = cellArray.map((cell) => cellToCoords(cell));

  // Find bounds
  const maxCol = Math.max(...coords.map(([col, _]) => col));
  const maxRow = Math.max(...coords.map(([_, row]) => row));

  let largestRect: RangeRectangle | null = null;
  let maxArea = 0;

  // Try each cell as a potential top-left corner
  for (const [startCol, startRow] of coords) {
    // Find the largest rectangle starting from this top-left corner
    const rect = findLargestRectangleFromCorner(
      cells,
      startCol,
      startRow,
      maxCol,
      maxRow,
    );

    if (rect) {
      const area = (rect.right - rect.left + 1) * (rect.bottom - rect.top + 1);
      if (area > maxArea) {
        maxArea = area;
        largestRect = rect;
      }
    }
  }

  return largestRect;
}

/**
 * Find the largest rectangle starting from a specific top-left corner
 */
function findLargestRectangleFromCorner(
  cells: Set<string>,
  startCol: number,
  startRow: number,
  maxCol: number,
  maxRow: number,
): RangeRectangle | null {
  let bestRect: RangeRectangle | null = null;
  let maxArea = 0;

  // Try expanding rightward first, then downward
  for (let endCol = startCol; endCol <= maxCol; endCol++) {
    // Check if we can form a valid row from startCol to endCol at startRow
    let canFormRow = true;
    for (let col = startCol; col <= endCol; col++) {
      if (!cells.has(coordsToCell(col, startRow))) {
        canFormRow = false;
        break;
      }
    }

    if (!canFormRow) break;

    // Now try expanding downward
    for (let endRow = startRow; endRow <= maxRow; endRow++) {
      // Check if we can form a complete rectangle
      let canFormRect = true;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          if (!cells.has(coordsToCell(col, row))) {
            canFormRect = false;
            break;
          }
        }
        if (!canFormRect) break;
      }

      if (!canFormRect) break;

      // We found a valid rectangle
      const area = (endCol - startCol + 1) * (endRow - startRow + 1);
      if (area > maxArea) {
        maxArea = area;
        bestRect = {
          left: startCol,
          top: startRow,
          right: endCol,
          bottom: endRow,
          originalRange: `${coordsToCell(startCol, startRow)}:${
            coordsToCell(endCol, endRow)
          }`,
        };
      }
    }
  }

  return bestRect;
}

/**
 * Backtracking algorithm to find optimal set of non-overlapping rectangles
 */
function backtrackingRectangleSelection(
  allCells: Set<string>,
): RangeRectangle[] {
  const selectedRectangles: RangeRectangle[] = [];
  const remainingCells = new Set(allCells);

  // Recursively find the largest rectangle until no cells remain
  while (remainingCells.size > 0) {
    const largestRect = findLargestRectangle(remainingCells);

    if (!largestRect) {
      // If we can't find any rectangle, handle remaining cells individually
      for (const cell of remainingCells) {
        const [col, row] = cellToCoords(cell);
        selectedRectangles.push({
          left: col,
          top: row,
          right: col,
          bottom: row,
          originalRange: cell,
        });
      }
      break;
    }

    // Add the largest rectangle to our selection
    selectedRectangles.push(largestRect);

    // Remove all cells covered by this rectangle
    const rectCells = getRectangleCells(largestRect);
    rectCells.forEach((cell) => remainingCells.delete(cell));
  }

  return selectedRectangles;
}

/**
 * Analyze a list of ranges and return optimal non-overlapping ranges using backtracking algorithm
 * @param ranges - Array of Excel range strings (e.g., ["A1:B10", "A2:H8", "C5:D6"])
 * @returns Object containing array of optimally sized ranges (largest rectangles first)
 */
export function analyzeRangeOverlaps(ranges: string[]): RangeOverlapResult {
  if (ranges.length === 0) {
    return { ranges: [] };
  }

  if (ranges.length === 1) {
    return { ranges: ranges };
  }

  // Get all unique cells from input ranges
  const allCells = getCellsFromRanges(ranges);

  // Use backtracking algorithm to find optimal rectangles (largest first)
  const optimalRectangles = backtrackingRectangleSelection(allCells);

  // Convert back to range strings
  const resultRanges = optimalRectangles.map((rect) => rectangleToRange(rect));

  return {
    ranges: Array.from(new Set(resultRanges)),
  };
}

/**
 * Helper function to get all cells from ranges
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
 * Helper function to get cells from a single range
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
 * Convert rectangle coordinates back to Excel range
 */
function rectangleToRange(rect: RangeRectangle): string {
  const startCell = `${numberToColumn(rect.left)}${rect.top}`;
  const endCell = `${numberToColumn(rect.right)}${rect.bottom}`;

  return startCell === endCell ? startCell : `${startCell}:${endCell}`;
}
