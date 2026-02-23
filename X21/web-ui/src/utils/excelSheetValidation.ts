const MAX_SHEET_NAME_LENGTH = 31;
const INVALID_SHEET_CHARS = /[:\\/?*\[\]]/;

export const normalizeSheetName = (raw: string): string => {
  let cleaned = raw.trim();
  if (cleaned.startsWith("@")) {
    cleaned = cleaned.slice(1);
  }
  if (cleaned.endsWith("!")) {
    cleaned = cleaned.slice(0, -1);
  }
  if (cleaned.startsWith("'") && cleaned.endsWith("'")) {
    cleaned = cleaned.slice(1, -1).replace(/''/g, "'");
  }
  return cleaned.trim();
};

export const isValidSheetNameFormat = (raw: string): boolean => {
  const normalized = normalizeSheetName(raw);
  if (!normalized) return false;
  if (normalized.length > MAX_SHEET_NAME_LENGTH) return false;
  if (INVALID_SHEET_CHARS.test(normalized)) return false;
  return true;
};

export type SheetLookup = {
  currentSheets: Set<string>;
  workbookSheets: Map<string, Set<string>>;
};

const normalizeWorkbookToken = (raw: string): string => {
  let cleaned = raw.trim();
  if (cleaned.startsWith("'") && cleaned.endsWith("'")) {
    cleaned = cleaned.slice(1, -1).replace(/''/g, "'");
  }
  return cleaned.trim();
};

const getWorkbookBasename = (value: string): string => {
  const normalized = value.replace(/[\\/]+$/, "");
  const parts = normalized.split(/[\\/]/);
  return parts[parts.length - 1] || normalized;
};

const addSheetNames = (target: Set<string>, names?: string[]) => {
  if (!names?.length) return;
  names.forEach((name) => {
    if (name) target.add(name.toLowerCase());
  });
};

const addWorkbookSheetNames = (
  workbookSheets: Map<string, Set<string>>,
  workbookToken: string | undefined,
  sheetNames?: string[],
) => {
  if (!workbookToken) return;
  const normalized = normalizeWorkbookToken(workbookToken).toLowerCase();
  if (!normalized) return;
  let sheetSet = workbookSheets.get(normalized);
  if (!sheetSet) {
    sheetSet = new Set<string>();
    workbookSheets.set(normalized, sheetSet);
  }
  addSheetNames(sheetSet, sheetNames);
};

const addWorkbookAliases = (
  workbookSheets: Map<string, Set<string>>,
  workbookToken: string | undefined,
  sheetNames?: string[],
) => {
  if (!workbookToken) return;
  const normalized = normalizeWorkbookToken(workbookToken).toLowerCase();
  if (!normalized) return;
  addWorkbookSheetNames(workbookSheets, normalized, sheetNames);
  const basename = getWorkbookBasename(normalized);
  if (basename && basename !== normalized) {
    addWorkbookSheetNames(workbookSheets, basename, sheetNames);
  }
};

export const buildSheetLookup = (
  currentSheets: string[],
  openWorkbooks?: Array<{
    workbookName?: string;
    workbookFullName?: string;
    sheets?: string[];
  }>,
  currentWorkbookName?: string,
  otherWorkbookSheets?: Array<{
    workbookName?: string;
    workbookFullName?: string;
    sheetName: string;
  }>,
): SheetLookup => {
  const currentSheetSet = new Set<string>();
  addSheetNames(currentSheetSet, currentSheets);

  const workbookSheets = new Map<string, Set<string>>();
  if (currentWorkbookName) {
    addWorkbookAliases(workbookSheets, currentWorkbookName, currentSheets);
  }

  openWorkbooks?.forEach((wb) => {
    if (!wb?.sheets?.length) return;
    addWorkbookAliases(workbookSheets, wb.workbookName, wb.sheets);
    addWorkbookAliases(workbookSheets, wb.workbookFullName, wb.sheets);
  });

  otherWorkbookSheets?.forEach((entry) => {
    if (!entry?.sheetName) return;
    const sheetNames = [entry.sheetName];
    addWorkbookAliases(workbookSheets, entry.workbookName, sheetNames);
    addWorkbookAliases(workbookSheets, entry.workbookFullName, sheetNames);
  });

  return { currentSheets: currentSheetSet, workbookSheets };
};

const getSheetSetForWorkbook = (
  sheetLookup: SheetLookup,
  workbookName?: string,
): Set<string> | null => {
  if (!workbookName) return sheetLookup.currentSheets;
  const normalized = normalizeWorkbookToken(workbookName).toLowerCase();
  if (!normalized) return null;
  return (
    sheetLookup.workbookSheets.get(normalized) ??
    sheetLookup.workbookSheets.get(getWorkbookBasename(normalized)) ??
    null
  );
};

export const isKnownSheetName = (
  raw: string,
  sheetLookup: SheetLookup,
  workbookName?: string,
): boolean => {
  const normalized = normalizeSheetName(raw).toLowerCase();
  if (!normalized) return false;

  const sheetSet = getSheetSetForWorkbook(sheetLookup, workbookName);
  if (!sheetSet || sheetSet.size === 0) {
    return workbookName ? false : isValidSheetNameFormat(raw);
  }
  return sheetSet.has(normalized);
};

export const buildWorkbookNameSet = (
  openWorkbooks: Array<{ workbookName: string; workbookFullName?: string }>,
  currentWorkbookName?: string,
  extraWorkbooks?: Array<{ workbookName?: string; workbookFullName?: string }>,
): Set<string> => {
  const names = new Set<string>();
  openWorkbooks.forEach((wb) => {
    if (wb?.workbookName) {
      const lower = wb.workbookName.toLowerCase();
      names.add(lower);
      names.add(getWorkbookBasename(lower));
    }
    if (wb?.workbookFullName) {
      const lower = wb.workbookFullName.toLowerCase();
      names.add(lower);
      names.add(getWorkbookBasename(lower));
    }
  });
  if (currentWorkbookName) {
    const lower = currentWorkbookName.toLowerCase();
    names.add(lower);
    names.add(getWorkbookBasename(lower));
  }
  extraWorkbooks?.forEach((wb) => {
    if (wb?.workbookName) {
      const lower = wb.workbookName.toLowerCase();
      names.add(lower);
      names.add(getWorkbookBasename(lower));
    }
    if (wb?.workbookFullName) {
      const lower = wb.workbookFullName.toLowerCase();
      names.add(lower);
      names.add(getWorkbookBasename(lower));
    }
  });
  return names;
};

export const isKnownWorkbook = (
  raw: string,
  knownWorkbooks: Set<string>,
): boolean => {
  if (knownWorkbooks.size === 0) return false;
  const normalized = normalizeWorkbookToken(raw).toLowerCase();
  if (!normalized) return false;
  if (knownWorkbooks.has(normalized)) return true;
  return knownWorkbooks.has(getWorkbookBasename(normalized));
};
