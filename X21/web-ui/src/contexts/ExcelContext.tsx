import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
} from "react";
import type { ReactNode } from "react";

import { getApiBase } from "@/services/apiBase";
import { webViewBridge } from "@/services/webViewBridge";
import {
  buildSheetLookup,
  buildWorkbookNameSet,
  type SheetLookup,
} from "@/utils/excelSheetValidation";

type OpenWorkbook = {
  workbookName: string;
  workbookFullName?: string;
  sheets?: string[];
};

type ExcelContextValue = {
  sheetNames: string[];
  openWorkbooks: OpenWorkbook[];
  sheetLookup: SheetLookup;
  knownWorkbooks: Set<string>;
  isLoading: boolean;
  error: string | null;
  refresh: () => Promise<void>;
};

const ExcelContext = createContext<ExcelContextValue | null>(null);

export function ExcelContextProvider({ children }: { children: ReactNode }) {
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [openWorkbooks, setOpenWorkbooks] = useState<OpenWorkbook[]>([]);
  const [sheetLookup, setSheetLookup] = useState<SheetLookup>({
    currentSheets: new Set(),
    workbookSheets: new Map(),
  });
  const [knownWorkbooks, setKnownWorkbooks] = useState<Set<string>>(new Set());
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refresh = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const [sheetNames, workbookName] = await Promise.all([
        webViewBridge.getWorksheetNames().catch(() => [] as string[]),
        webViewBridge.getWorkbookName().catch(() => ""),
      ]);

      let openWorkbooks: OpenWorkbook[] = [];
      try {
        const apiBase = await getApiBase();
        const listResp = await fetch(`${apiBase}/api/excel/open-workbooks`);
        if (listResp.ok) {
          const listJson = await listResp.json();
          openWorkbooks = listJson?.workbooks ?? [];
        }
      } catch {
        openWorkbooks = [];
      }

      const confirmedWorkbooks = openWorkbooks.filter(
        (wb) => Array.isArray(wb.sheets) && wb.sheets.length > 0,
      );
      setSheetNames(sheetNames);
      setOpenWorkbooks(confirmedWorkbooks);
      setSheetLookup(
        buildSheetLookup(sheetNames, confirmedWorkbooks, workbookName),
      );
      setKnownWorkbooks(buildWorkbookNameSet(confirmedWorkbooks, workbookName));
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "Failed to refresh Excel context";
      console.warn("[ExcelContext] refresh failed", err);
      setError(message);
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const value = useMemo<ExcelContextValue>(
    () => ({
      sheetNames,
      openWorkbooks,
      sheetLookup,
      knownWorkbooks,
      isLoading,
      error,
      refresh,
    }),
    [
      sheetNames,
      openWorkbooks,
      sheetLookup,
      knownWorkbooks,
      isLoading,
      error,
      refresh,
    ],
  );

  return (
    <ExcelContext.Provider value={value}>{children}</ExcelContext.Provider>
  );
}

export function useExcelContext() {
  return useContext(ExcelContext);
}
