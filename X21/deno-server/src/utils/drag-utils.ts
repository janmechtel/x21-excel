export function isValidAutoFillOverlap(
  sourceA1: string,
  destA1: string,
): boolean {
  const source = parseA1(sourceA1);
  const dest = parseA1(destA1);

  const srcRowEnd = source.row + source.rowCount - 1;
  const srcColEnd = source.col + source.colCount - 1;
  const destRowEnd = dest.row + dest.rowCount - 1;
  const destColEnd = dest.col + dest.colCount - 1;

  // Check if source is completely contained within destination
  const sourceInDest = dest.row <= source.row &&
    dest.col <= source.col &&
    destRowEnd >= srcRowEnd &&
    destColEnd >= srcColEnd;

  if (!sourceInDest) {
    return false; // Source must be within destination for valid AutoFill
  }

  // Check for valid AutoFill directions

  // Fill Down: Source at top of destination, same width
  const isValidFillDown = source.row === dest.row &&
    source.col === dest.col &&
    source.colCount === dest.colCount &&
    dest.rowCount > source.rowCount;

  // Fill Up: Source at bottom of destination, same width
  const isValidFillUp = srcRowEnd === destRowEnd &&
    source.col === dest.col &&
    source.colCount === dest.colCount &&
    dest.rowCount > source.rowCount;

  // Fill Right: Source at left of destination, same height
  const isValidFillRight = source.row === dest.row &&
    source.col === dest.col &&
    source.rowCount === dest.rowCount &&
    dest.colCount > source.colCount;

  // Fill Left: Source at right of destination, same height
  const isValidFillLeft = source.row === dest.row &&
    srcColEnd === destColEnd &&
    source.rowCount === dest.rowCount &&
    dest.colCount > source.colCount;

  // 2D AutoFill: Source at top-left corner, destination extends both ways
  const isValid2DFill = source.row === dest.row &&
    source.col === dest.col &&
    dest.rowCount > source.rowCount &&
    dest.colCount > source.colCount;

  return isValidFillDown || isValidFillUp || isValidFillRight ||
    isValidFillLeft || isValid2DFill;
}

export function getModifiedRangeA1(sourceA1: string, destA1: string): string {
  const source = parseA1(sourceA1);
  const dest = parseA1(destA1);

  const srcRowEnd = source.row + source.rowCount - 1;
  const srcColEnd = source.col + source.colCount - 1;
  const destRowEnd = dest.row + dest.rowCount - 1;
  const destColEnd = dest.col + dest.colCount - 1;

  // Fill Down
  if (
    destRowEnd > srcRowEnd && dest.col === source.col &&
    dest.colCount === source.colCount
  ) {
    return rangeToA1(
      srcRowEnd + 1,
      source.col,
      destRowEnd - srcRowEnd,
      source.colCount,
    );
  }

  // Fill Up
  if (
    dest.row < source.row && dest.col === source.col &&
    dest.colCount === source.colCount
  ) {
    return rangeToA1(
      dest.row,
      source.col,
      source.row - dest.row,
      source.colCount,
    );
  }

  // Fill Right
  if (
    destColEnd > srcColEnd && dest.row === source.row &&
    dest.rowCount === source.rowCount
  ) {
    return rangeToA1(
      source.row,
      srcColEnd + 1,
      source.rowCount,
      destColEnd - srcColEnd,
    );
  }

  // Fill Left
  if (
    dest.col < source.col && dest.row === source.row &&
    dest.rowCount === source.rowCount
  ) {
    return rangeToA1(
      source.row,
      dest.col,
      source.rowCount,
      source.col - dest.col,
    );
  }

  // 2D extension case
  return rangeToA1(
    srcRowEnd + 1, // start row below source
    srcColEnd + 1, // start col right of source
    destRowEnd - srcRowEnd, // rows extended
    destColEnd - srcColEnd, // cols extended
  );
}

function rangeToA1(
  row: number,
  col: number,
  rowCount: number,
  colCount: number,
): string {
  const start = `${colToLetter(col)}${row + 1}`;
  const end = `${colToLetter(col + colCount - 1)}${row + rowCount}`;
  return `${start}:${end}`;
}

interface RangeInfo {
  row: number; // top-left row (0-based)
  col: number; // top-left col (0-based)
  rowCount: number;
  colCount: number;
}

function letterToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col *= 26;
    col += letters.charCodeAt(i) - 64; // 'A' = 65
  }
  return col - 1; // 0-based
}

function colToLetter(col: number): string {
  let letters = "";
  while (col >= 0) {
    letters = String.fromCharCode((col % 26) + 65) + letters;
    col = Math.floor(col / 26) - 1;
  }
  return letters;
}

function parseA1(a1: string): RangeInfo {
  const parts = a1.split(":");
  const parseCell = (cell: string) => {
    const match = cell.match(/^([A-Z]+)([0-9]+)$/i);
    if (!match) throw new Error(`Invalid A1 address: ${cell}`);
    const col = letterToCol(match[1].toUpperCase());
    const row = parseInt(match[2], 10) - 1;
    return { row, col };
  };

  const start = parseCell(parts[0]);
  const end = parts[1] ? parseCell(parts[1]) : start;

  const row = Math.min(start.row, end.row);
  const col = Math.min(start.col, end.col);
  const rowCount = Math.abs(end.row - start.row) + 1;
  const colCount = Math.abs(end.col - start.col) + 1;

  return { row, col, rowCount, colCount };
}
