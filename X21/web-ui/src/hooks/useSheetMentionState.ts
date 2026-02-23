import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { webViewBridge } from "@/services/webViewBridge";
import { getApiBase } from "@/services/apiBase";

interface SheetMentionState {
  query: string;
  startIndex: number;
  endIndex: number;
}

const getWorkbookIdentifier = (workbook: {
  workbookName?: string;
  workbookFullName?: string;
}) => {
  const fullName = workbook.workbookFullName?.trim();
  if (fullName) return fullName;
  return workbook.workbookName?.trim() ?? "";
};

const getWorkbookBasename = (value?: string) => {
  const trimmed = value?.trim() ?? "";
  if (!trimmed) return "";
  const normalized = trimmed.replace(/[\\/]+$/, "");
  const parts = normalized.split(/[\\/]/);
  return parts[parts.length - 1] || normalized;
};

const normalizeWorkbookIdentity = (value?: string) => {
  const trimmed = value?.trim() ?? "";
  if (!trimmed) return "";
  return trimmed.toLowerCase();
};

export function useSheetMentionState() {
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [currentWorkbookName, setCurrentWorkbookName] = useState("");
  const [openWorkbooks, setOpenWorkbooks] = useState<
    Array<{ workbookName: string; workbookFullName?: string }>
  >([]);
  const [mentionState, setMentionState] = useState<SheetMentionState | null>(
    null,
  );
  const [highlightIndex, setHighlightIndex] = useState(0);
  const [otherWorkbookSheets, setOtherWorkbookSheets] = useState<
    Array<{
      workbookName: string;
      sheetName: string;
      workbookFullName?: string;
    }>
  >([]);
  const [isRequestingOtherWorkbooks, setIsRequestingOtherWorkbooks] =
    useState(false);
  const isRequestingOtherWorkbooksRef = useRef(false);
  const isLoadingOtherWorkbooksRef = useRef(false);

  const mentionQuery = mentionState?.query ?? "";

  const refreshSheetNames = useCallback(async () => {
    const [names, workbookName] = await Promise.all([
      webViewBridge.getWorksheetNames(),
      webViewBridge.getWorkbookName().catch(() => ""),
    ]);
    setSheetNames(names);
    if (workbookName) {
      setCurrentWorkbookName(workbookName);
    }
  }, []);

  const refreshOpenWorkbooks = useCallback(async () => {
    try {
      const apiBase = await getApiBase();
      const listResp = await fetch(`${apiBase}/api/excel/open-workbooks`);
      if (!listResp.ok) {
        throw new Error(`Failed to list workbooks: ${listResp.status}`);
      }
      const listJson = await listResp.json();
      const workbooks: Array<{
        workbookName?: string;
        workbookFullName?: string;
      }> = listJson?.workbooks ?? [];
      setOpenWorkbooks(
        workbooks
          .filter((wb) => wb?.workbookName || wb?.workbookFullName)
          .map((wb) => ({
            workbookName:
              (typeof wb?.workbookName === "string" &&
                wb.workbookName.trim()) ||
              getWorkbookBasename(wb.workbookFullName),
            workbookFullName: wb.workbookFullName,
          })),
      );
    } catch (error) {
      console.warn("Failed to load open workbooks", error);
      setOpenWorkbooks([]);
    }
  }, []);

  useEffect(() => {
    void refreshSheetNames();
    void refreshOpenWorkbooks();
  }, [refreshSheetNames, refreshOpenWorkbooks]);

  useEffect(() => {
    if (mentionState && sheetNames.length === 0) {
      void refreshSheetNames();
    }
  }, [mentionState, sheetNames.length, refreshSheetNames]);

  // Auto-load other workbook sheets when mention palette opens
  const loadOtherWorkbookSheets = useCallback(async () => {
    if (
      isRequestingOtherWorkbooksRef.current ||
      isLoadingOtherWorkbooksRef.current
    )
      return;
    isLoadingOtherWorkbooksRef.current = true;

    try {
      isRequestingOtherWorkbooksRef.current = true;
      setIsRequestingOtherWorkbooks(true);
      const apiBase = await getApiBase();
      const currentWorkbook = await webViewBridge.getWorkbookName();
      const currentWorkbookName =
        typeof currentWorkbook === "string" ? currentWorkbook.trim() : "";
      const listResp = await fetch(`${apiBase}/api/excel/open-workbooks`);
      if (!listResp.ok) {
        throw new Error(`Failed to list workbooks: ${listResp.status}`);
      }
      const listJson = await listResp.json();
      const workbooks: Array<{
        workbookName?: string;
        workbookFullName?: string;
        sheets?: string[];
      }> = listJson?.workbooks ?? [];

      const sheetEntries: Array<{
        workbookName: string;
        sheetName: string;
        workbookFullName?: string;
      }> = [];
      const confirmedWorkbooks: Array<{
        workbookName: string;
        workbookFullName?: string;
      }> = [];
      const confirmedKeys = new Set<string>();

      const addConfirmedWorkbook = (
        workbookName: string,
        workbookFullName?: string,
      ) => {
        const key = normalizeWorkbookIdentity(workbookFullName || workbookName);
        if (!key || confirmedKeys.has(key)) return;
        confirmedKeys.add(key);
        confirmedWorkbooks.push({
          workbookName,
          workbookFullName: workbookFullName || undefined,
        });
      };

      for (const wb of workbooks) {
        const workbookName =
          (typeof wb?.workbookName === "string" && wb.workbookName.trim()) ||
          getWorkbookBasename(wb.workbookFullName);
        if (!workbookName) continue;
        const workbookFullName =
          typeof wb?.workbookFullName === "string"
            ? wb.workbookFullName.trim()
            : "";
        const normalizedCurrent =
          normalizeWorkbookIdentity(currentWorkbookName);
        const normalizedCurrentBase = normalizeWorkbookIdentity(
          getWorkbookBasename(currentWorkbookName),
        );
        const normalizedWorkbook = normalizeWorkbookIdentity(workbookName);
        const normalizedFullName = normalizeWorkbookIdentity(workbookFullName);
        if (
          (normalizedCurrent &&
            (normalizedWorkbook === normalizedCurrent ||
              (normalizedFullName &&
                normalizedFullName === normalizedCurrent))) ||
          (normalizedCurrentBase &&
            (normalizedWorkbook === normalizedCurrentBase ||
              (normalizedFullName &&
                normalizedFullName === normalizedCurrentBase)))
        )
          continue;
        const sheetNames = Array.isArray(wb.sheets)
          ? wb.sheets.filter(
              (name): name is string =>
                typeof name === "string" && name.trim().length > 0,
            )
          : [];
        if (sheetNames.length > 0) {
          addConfirmedWorkbook(workbookName, workbookFullName);
          sheetNames.forEach((sheetName) => {
            sheetEntries.push({
              workbookName,
              workbookFullName: workbookFullName || undefined,
              sheetName,
            });
          });
          continue;
        }
        try {
          const workbookIdentifier = getWorkbookIdentifier({
            workbookName,
            workbookFullName,
          });
          if (!workbookIdentifier) {
            continue;
          }
          const metaResp = await fetch(
            `${apiBase}/api/excel/workbook-metadata?workbookName=${encodeURIComponent(
              workbookIdentifier,
            )}`,
          );
          if (!metaResp.ok) {
            throw new Error(`Failed metadata for ${workbookName}`);
          }
          const metadata = await metaResp.json();
          const sheets = Array.isArray(metadata?.sheets)
            ? metadata.sheets
            : (metadata?.allSheets ?? []).map((name: string) => ({ name }));
          sheets.forEach((sheet: any) => {
            if (!sheet?.name) return;
            sheetEntries.push({
              workbookName,
              workbookFullName: workbookFullName || undefined,
              sheetName: sheet.name,
            });
          });
          if (sheets.length > 0) {
            addConfirmedWorkbook(workbookName, workbookFullName);
          }
        } catch (err) {
          console.warn("Failed to load sheets for workbook", workbookName, err);
        }
      }

      setOtherWorkbookSheets(sheetEntries);
      setOpenWorkbooks(confirmedWorkbooks);
    } catch (error) {
      console.warn("Auto-load other workbooks failed", error);
      setOtherWorkbookSheets([]);
    } finally {
      setIsRequestingOtherWorkbooks(false);
      isRequestingOtherWorkbooksRef.current = false;
      isLoadingOtherWorkbooksRef.current = false;
    }
  }, []);

  const wasMentionActiveRef = useRef(false);

  // Auto-load when mention palette opens
  useEffect(() => {
    const isMentionActive = Boolean(mentionState);
    const wasMentionActive = wasMentionActiveRef.current;
    wasMentionActiveRef.current = isMentionActive;

    if (!wasMentionActive && isMentionActive) {
      void refreshSheetNames();
      if (
        !isRequestingOtherWorkbooksRef.current &&
        !isLoadingOtherWorkbooksRef.current
      ) {
        void loadOtherWorkbookSheets();
      }
    }
  }, [mentionState, loadOtherWorkbookSheets, refreshSheetNames]);

  const updateMentionTrigger = useCallback(
    (value: string, cursorPosition: number | null) => {
      if (cursorPosition === null || cursorPosition === undefined) {
        setMentionState(null);
        return;
      }

      const uptoCursor = value.slice(0, cursorPosition);
      const mentionMatch = uptoCursor.match(/(^|[^A-Za-z0-9_])@([^\s@]*)$/);

      if (!mentionMatch) {
        setMentionState(null);
        return;
      }

      const query = mentionMatch[2] ?? "";
      const startIndex = cursorPosition - query.length - 1;
      const endIndex = cursorPosition;

      setMentionState({ query, startIndex, endIndex });
    },
    [],
  );

  const resetMentionState = useCallback(() => {
    setMentionState(null);
    setHighlightIndex(0);
  }, []);

  const mentionMatches = useMemo(() => {
    const term = mentionQuery.trim().toLowerCase();
    if (!term) return sheetNames;

    return sheetNames.filter((name) => name.toLowerCase().includes(term));
  }, [sheetNames, mentionQuery]);

  useEffect(() => {
    if (!mentionState) {
      setHighlightIndex(0);
      return;
    }

    if (mentionMatches.length === 0) {
      setHighlightIndex(0);
      return;
    }

    setHighlightIndex((prev) => Math.min(prev, mentionMatches.length - 1));
  }, [mentionState, mentionMatches]);

  const highlightedSheetName =
    mentionMatches.length > 0
      ? mentionMatches[Math.min(highlightIndex, mentionMatches.length - 1)]
      : null;

  const isMentionActive = Boolean(mentionState);

  return {
    sheetMentionState: mentionState,
    mentionMatches,
    mentionHighlightIndex: highlightIndex,
    setMentionHighlightIndex: setHighlightIndex,
    mentionQuery,
    isMentionActive,
    highlightedSheetName,
    sheetNames,
    currentWorkbookName,
    openWorkbooks,
    updateMentionTrigger,
    resetMentionState,
    openSheetMentionPalette: (
      currentPrompt: string,
      cursorPosition?: number,
    ) => {
      const cursor = cursorPosition ?? currentPrompt.length;
      setMentionState({
        query: "",
        startIndex: cursor,
        endIndex: cursor,
      });
      setHighlightIndex(0);
      void refreshSheetNames();
    },
    otherWorkbookSheets,
    isRequestingOtherWorkbooks,
    requestOtherWorkbooks: loadOtherWorkbookSheets, // For manual triggering (e.g., checkbox toggle)
    clearOtherWorkbookSheets: () => setOtherWorkbookSheets([]),
    refreshSheetNames,
  };
}

export type { SheetMentionState };
