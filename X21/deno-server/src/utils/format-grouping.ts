import {
  convertRectanglesToExcelRanges,
  findAllRectangles,
  getTopLeftCell,
} from "./excel-range.ts";
import { ReadFormatFinalResponseList } from "../types/index.ts";

/**
 * Groups cell format data by unique formats and their corresponding Excel ranges
 * @param cellFormatData - Record of cell addresses to format objects
 * @param dataType - Description for logging (e.g., "old values", "new values")
 * @returns Grouped format data as { format: any, ranges: string[] }[]
 */
export function groupCellFormatsByRanges(
  cellFormatData: Record<string, any>,
): ReadFormatFinalResponseList {
  if (Object.keys(cellFormatData).length === 0) {
    return [];
  }

  // Get unique formats (not just values)
  const uniqueFormats = getUniqueFormats(cellFormatData);

  // Create matrix structure with format indices
  const formatMatrix = createFormatMatrix(cellFormatData, uniqueFormats);

  // Find rectangular regions of identical formats
  const rectangles = findAllRectangles(formatMatrix);

  // Convert rectangles to Excel ranges
  const topLeftCell = getTopLeftCell(cellFormatData);
  const excelRanges = convertRectanglesToExcelRanges(rectangles, topLeftCell);

  // Group ranges by format
  return groupRangesByFormat(rectangles, excelRanges, uniqueFormats);
}

/**
 * Gets unique formats from cell format data
 */
function getUniqueFormats(cellData: Record<string, any>): any[] {
  const formatSet = new Set<string>();
  const uniqueFormats: any[] = [];

  for (const cellFormat of Object.values(cellData)) {
    const formatKey = JSON.stringify(cellFormat);
    if (!formatSet.has(formatKey)) {
      formatSet.add(formatKey);
      uniqueFormats.push(cellFormat);
    }
  }

  return uniqueFormats;
}

/**
 * Creates a matrix representation of format indices
 */
function createFormatMatrix(
  cellData: Record<string, any>,
  uniqueFormats: any[],
): number[][] {
  const formatToIndexMap = new Map<string, number>();
  uniqueFormats.forEach((format, index) => {
    formatToIndexMap.set(JSON.stringify(format), index);
  });

  const cellAddresses = Object.keys(cellData);
  const { minRow, maxRow, minCol, maxCol } = findRangeBounds(cellAddresses);

  const matrix: number[][] = [];
  for (let row = minRow; row <= maxRow; row++) {
    const matrixRow: number[] = [];
    for (let col = minCol; col <= maxCol; col++) {
      const cellAddress = `${numberToColumn(col)}${row}`;
      if (cellData[cellAddress]) {
        const formatKey = JSON.stringify(cellData[cellAddress]);
        const formatIndex = formatToIndexMap.get(formatKey);
        matrixRow.push(formatIndex !== undefined ? formatIndex : -1);
      } else {
        matrixRow.push(-1); // -1 for empty cells
      }
    }
    matrix.push(matrixRow);
  }

  return matrix;
}

/**
 * Groups ranges by their format
 */
function groupRangesByFormat(
  rectangles: any[],
  excelRanges: string[],
  uniqueFormats: any[],
): { format: any; ranges: string[] }[] {
  const formatToRangesMap = new Map<string, string[]>();
  const result: { format: any; ranges: string[] }[] = [];

  rectangles.forEach((rect, index) => {
    const formatValue = rect.value;
    const range = excelRanges[index];

    if (formatValue >= 0 && formatValue < uniqueFormats.length) {
      const format = uniqueFormats[formatValue];
      const formatKey = JSON.stringify(format);

      if (!formatToRangesMap.has(formatKey)) {
        formatToRangesMap.set(formatKey, []);
      }
      formatToRangesMap.get(formatKey)!.push(range);
    }
  });

  for (const [formatKey, ranges] of formatToRangesMap) {
    const format = JSON.parse(formatKey);
    result.push({ format, ranges });
  }

  return result;
}

// Import these from excel-range.ts - we need to make sure they're exported
import { findRangeBounds, numberToColumn } from "./excel-range.ts";
