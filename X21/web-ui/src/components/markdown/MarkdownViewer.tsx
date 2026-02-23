import React from "react";
import ReactMarkdown, { type Components } from "react-markdown";
import remarkGfm from "remark-gfm";
import { cn } from "@/lib/utils";
import {
  remarkExcelRanges,
  remarkExcelSheetMentions,
} from "@/lib/remarkExcelRanges";
import { ExcelRangePill } from "@/components/excel/ExcelRangePill";
import { ExcelSheetMentionPill } from "@/components/excel/ExcelSheetMentionPill";
import { webViewBridge } from "@/services/webViewBridge";
import { useExcelContext } from "@/contexts/ExcelContext";
import {
  isKnownSheetName,
  isKnownWorkbook,
  isValidSheetNameFormat,
  normalizeSheetName,
} from "@/utils/excelSheetValidation";

interface MarkdownViewerProps {
  text: string;
  className?: string;
  onRangeClick?: (range: string, workbookName?: string) => void;
}

const EXCEL_PROTOCOL = "excel-range://";
const EXCEL_SHEET_PROTOCOL = "excel-sheet://";
const EXCEL_FILE_PROTOCOL = "excel-file://";
const EMPTY_SET = new Set<string>();
const EMPTY_SHEET_LOOKUP = {
  currentSheets: EMPTY_SET,
  workbookSheets: new Map<string, Set<string>>(),
};

// Preprocessor to wrap Excel file names in backticks for markdown processing
const preprocessText = (text: string): string => {
  let processedText = text;

  // Enhanced pattern to detect structured output with location and filename
  // Matches patterns like:
  // - "Output file: filename.xlsx\nLocation: D:\path\"
  // - "New output file: filename.xlsx" followed by "Location: D:\path\"
  const structuredPattern =
    /(?:output file|new file|file):\s*[`]?([\w\-_]+\.(?:xlsx|xlsm|xls|xlsb))[`]?\s*\n\s*(?:Location|Path):\s*((?:[A-Z]:\\)?[\w\-_./\\]+)\\?\s*/gis;

  processedText = processedText.replace(
    structuredPattern,
    (_match, filename, location) => {
      const fullPath =
        location.endsWith("\\") || location.endsWith("/")
          ? `${location}${filename}`
          : `${location}\\${filename}`;
      // Create a markdown link with the full path
      return `**Output file:** [\`${fullPath}\`](excel-file://${fullPath})\n**Location:** ${location}`;
    },
  );

  // Fallback: try to detect "Location: PATH" pattern and combine it with filenames
  // This handles the case where the LLM provides location and filename separately (reverse order)
  const locationPattern =
    /(?:Location|Path):\s*((?:[A-Z]:\\)?[\w\-_./\\]+)\\?\s*\n.*?(?:output file|New file|File).*?[`\s]?([\w\-_]+\.(?:xlsx|xlsm|xls|xlsb))/gis;
  processedText = processedText.replace(
    locationPattern,
    (match, location, filename) => {
      const fullPath =
        location.endsWith("\\") || location.endsWith("/")
          ? `${location}${filename}`
          : `${location}\\${filename}`;
      return match.replace(filename, `\`${fullPath}\``);
    },
  );

  // Match standalone Excel files (not already in code blocks or links)
  // Pattern: word characters, optional path separators, then .xlsx/.xlsm/.xls/.xlsb
  const excelFilePattern =
    /(?<!`|\[)(\b[\w\-_./\\]+\.(?:xlsx|xlsm|xls|xlsb)\b)(?!`|\])/gi;
  processedText = processedText.replace(excelFilePattern, "`$1`");

  return processedText;
};

export const MarkdownViewer: React.FC<MarkdownViewerProps> = ({
  text,
  className,
  onRangeClick,
}) => {
  const excelContext = useExcelContext();
  const sheetLookup = excelContext?.sheetLookup ?? EMPTY_SHEET_LOOKUP;
  const knownWorkbooks = excelContext?.knownWorkbooks ?? EMPTY_SET;

  const handleExcelRangeClick = (range: string, workbookName?: string) => {
    if (!range) return;
    onRangeClick?.(range, workbookName);
  };

  const quoteNameIfNeeded = (name?: string) => {
    if (!name) return "";
    const needsQuotes = /[\s,\/!]/.test(name) || name.includes("'");
    if (!needsQuotes) return name;
    return `'${name.replace(/'/g, "''")}'`;
  };

  const parseExcelTarget = (raw?: string) => {
    if (!raw) return null;

    const cleaned = raw.trim();
    const parts = cleaned.split("!");
    let workbookPart: string | undefined;
    let sheetPart: string | undefined;
    let cellRange: string | undefined;

    if (parts.length >= 3) {
      workbookPart = parts[0];
      sheetPart = parts[1];
      cellRange = parts.slice(2).join("!");
    } else if (parts.length === 2) {
      const combined = parts[0];
      cellRange = parts[1];
      const bracketMatch = combined.match(/^\[([^\]]+)\](.+)$/);
      if (bracketMatch) {
        workbookPart = bracketMatch[1];
        sheetPart = bracketMatch[2];
      } else if (combined.includes("/")) {
        const lastSlash = combined.lastIndexOf("/");
        workbookPart = combined.slice(0, lastSlash);
        sheetPart = combined.slice(lastSlash + 1);
      } else {
        sheetPart = combined;
      }
    } else {
      return null;
    }

    if (!cellRange || !sheetPart) return null;

    // If workbookPart is still embedded in sheetPart via '/', split it out
    if (!workbookPart && sheetPart.includes("/")) {
      const lastSlash = sheetPart.lastIndexOf("/");
      workbookPart = sheetPart.slice(0, lastSlash);
      sheetPart = sheetPart.slice(lastSlash + 1);
    }

    if (workbookPart && !isKnownWorkbook(workbookPart, knownWorkbooks)) {
      return null;
    }

    const rawSheetToken = sheetPart ?? "";
    const normalizedSheet = normalizeSheetName(rawSheetToken);
    const requiresKnownSheet =
      (rawSheetToken.startsWith("'") && rawSheetToken.endsWith("'")) ||
      /\s/.test(normalizedSheet);

    const hasSheetContext =
      sheetLookup.currentSheets.size > 0 || sheetLookup.workbookSheets.size > 0;
    if (hasSheetContext) {
      if (!isKnownSheetName(sheetPart, sheetLookup, workbookPart)) {
        return null;
      }
    } else {
      if (requiresKnownSheet) {
        return null;
      }
      if (!isValidSheetNameFormat(sheetPart)) {
        return null;
      }
    }

    const stripQuotes = (val?: string) =>
      val ? val.replace(/^'|'+$/g, "").trim() : undefined;
    const workbookName = stripQuotes(workbookPart);
    const sheetName = stripQuotes(sheetPart);
    const safeSheet = quoteNameIfNeeded(sheetName);

    const range = safeSheet ? `${safeSheet}!${cellRange}` : cellRange;
    return { range, workbookName };
  };

  const parseExcelSheetTarget = (raw?: string) => {
    if (!raw) return null;

    const cleaned = raw.trim().replace(/!$/, "");
    if (!cleaned) return null;

    let workbookPart: string | undefined;
    let sheetPart = cleaned;

    const bracketMatch = cleaned.match(/^\[([^\]]+)\](.+)$/);
    if (bracketMatch) {
      workbookPart = bracketMatch[1];
      sheetPart = bracketMatch[2];
    } else if (cleaned.includes("/")) {
      const lastSlash = cleaned.lastIndexOf("/");
      workbookPart = cleaned.slice(0, lastSlash);
      sheetPart = cleaned.slice(lastSlash + 1);
    }

    if (workbookPart && !isKnownWorkbook(workbookPart, knownWorkbooks)) {
      return null;
    }

    const rawSheetToken = sheetPart ?? "";
    const normalizedSheet = normalizeSheetName(rawSheetToken);
    const requiresKnownSheet =
      (rawSheetToken.startsWith("'") && rawSheetToken.endsWith("'")) ||
      /\s/.test(normalizedSheet);
    const hasSheetContext =
      sheetLookup.currentSheets.size > 0 || sheetLookup.workbookSheets.size > 0;

    if (hasSheetContext) {
      if (!isKnownSheetName(sheetPart, sheetLookup, workbookPart)) {
        return null;
      }
    } else {
      if (requiresKnownSheet) {
        return null;
      }
      if (!isValidSheetNameFormat(sheetPart)) {
        return null;
      }
    }

    const stripQuotes = (val?: string) =>
      val ? val.replace(/^'|'+$/g, "").trim() : undefined;
    const workbookName = stripQuotes(workbookPart);
    const sheetName = stripQuotes(sheetPart);
    const safeSheet = quoteNameIfNeeded(sheetName);
    const displayLabel = workbookName
      ? `${workbookName}/${sheetName}`
      : sheetName;

    return { sheetName, safeSheet, workbookName, displayLabel };
  };

  const handleExcelFileClick = async (filePath: string) => {
    console.log("[MarkdownViewer] handleExcelFileClick called with:", filePath);
    try {
      const success = await webViewBridge.openWorkbook(filePath);
      console.log(
        "[MarkdownViewer] webViewBridge.openWorkbook returned:",
        success,
      );
      if (!success) {
        console.error("[MarkdownViewer] Failed to open workbook:", filePath);
      } else {
        console.log("[MarkdownViewer] Successfully opened workbook:", filePath);
      }
    } catch (error) {
      console.error("[MarkdownViewer] Error opening workbook:", error);
    }
  };

  const isExcelFile = (text: string): boolean => {
    return /\.(xlsx|xlsm|xls|xlsb)$/i.test(text);
  };

  const components: Components = {
    a: ({ href, children, ...props }) => {
      // Check for Excel file protocol
      if (href?.startsWith(EXCEL_FILE_PROTOCOL)) {
        const filePath = href.replace(EXCEL_FILE_PROTOCOL, "");
        console.log("[MarkdownViewer] Rendering Excel file link:", {
          href,
          filePath,
          children,
        });
        return (
          <button
            onClick={() => handleExcelFileClick(filePath)}
            className="text-green-600 dark:text-green-400 font-medium hover:underline cursor-pointer inline-flex items-center gap-1 bg-green-50 dark:bg-green-900/20 px-2 py-1 rounded"
          >
            <span className="text-sm">📊</span>
            {children}
          </button>
        );
      }

      let target: { range: string; workbookName?: string } | null = null;
      let sheetTarget: {
        sheetName?: string;
        safeSheet?: string;
        workbookName?: string;
        displayLabel?: string;
      } | null = null;

      if (href?.startsWith(EXCEL_PROTOCOL)) {
        target = parseExcelTarget(href.replace(EXCEL_PROTOCOL, ""));
        if (!target?.range) {
          return <span>{children}</span>;
        }
      } else if (href?.startsWith(EXCEL_SHEET_PROTOCOL)) {
        sheetTarget = parseExcelSheetTarget(
          href.replace(EXCEL_SHEET_PROTOCOL, ""),
        );
        if (!sheetTarget?.sheetName) {
          return <span>{children}</span>;
        }
      } else if (!href) {
        const text =
          typeof children === "string"
            ? children
            : Array.isArray(children) &&
              children.length === 1 &&
              typeof children[0] === "string"
            ? children[0]
            : null;

        if (text) {
          target = parseExcelTarget(text);
          if (!target?.range) {
            sheetTarget = parseExcelSheetTarget(text);
          }
        }
      }

      if (target?.range) {
        return (
          <ExcelRangePill
            range={target.range}
            workbookName={target.workbookName}
            onClick={handleExcelRangeClick}
          />
        );
      }
      if (sheetTarget?.sheetName) {
        return (
          <ExcelSheetMentionPill
            sheetName={sheetTarget.displayLabel ?? sheetTarget.sheetName}
            onClick={() => {
              const safeSheet = sheetTarget.safeSheet ?? sheetTarget.sheetName;
              if (!safeSheet) return;
              handleExcelRangeClick(
                `${safeSheet}!A1`,
                sheetTarget.workbookName,
              );
            }}
          />
        );
      }

      // Check if the link text is an Excel file
      const linkText =
        typeof children === "string"
          ? children
          : Array.isArray(children) &&
            children.length === 1 &&
            typeof children[0] === "string"
          ? children[0]
          : null;

      if (linkText && isExcelFile(linkText)) {
        console.log(
          "[MarkdownViewer] Rendering Excel file from link text:",
          linkText,
        );
        return (
          <button
            onClick={() => handleExcelFileClick(linkText)}
            className="text-green-600 dark:text-green-400 font-medium hover:underline cursor-pointer inline-flex items-center gap-1"
          >
            <span className="text-xs">📊</span>
            {children}
          </button>
        );
      }

      return (
        <a
          href={href}
          target={href ? "_blank" : undefined}
          rel={href ? "noreferrer" : undefined}
          className="text-blue-600 dark:text-blue-400 font-medium hover:underline"
          {...props}
        >
          {children}
        </a>
      );
    },
    code: ({ inline, className: codeClassName, children, ...props }: any) => {
      if (inline) {
        const text = typeof children === "string" ? children : null;

        // Check if this is an Excel file
        if (text && isExcelFile(text)) {
          console.log(
            "[MarkdownViewer] Rendering Excel file from inline code:",
            text,
          );
          return (
            <button
              onClick={() => handleExcelFileClick(text)}
              className="text-green-600 dark:text-green-400 font-semibold hover:underline cursor-pointer inline-flex items-center gap-1 text-[11px]"
            >
              <span className="text-xs">📊</span>
              {children}
            </button>
          );
        }

        return (
          <code
            className={cn("text-[11px] font-semibold", codeClassName)}
            {...props}
          >
            {children}
          </code>
        );
      }

      return (
        <pre className="my-2 rounded-lg bg-slate-100 dark:bg-slate-900/90 text-slate-800 dark:text-slate-100 text-[11px] overflow-x-auto px-3 py-2.5">
          <code className={codeClassName} {...props}>
            {children}
          </code>
        </pre>
      );
    },
    input: ({ type, checked, disabled, ...props }: any) => {
      // Custom styling for checkboxes in task lists
      if (type === "checkbox") {
        return (
          <input
            type="checkbox"
            defaultChecked={checked}
            className="mr-2 h-4 w-4 rounded border-slate-300 text-green-600 focus:ring-green-500 dark:border-slate-600 dark:bg-slate-700 cursor-pointer"
            {...props}
          />
        );
      }
      return <input type={type} {...props} />;
    },
  };

  const processedText = preprocessText(text);

  if (text !== processedText) {
    console.log("[MarkdownViewer] Text was preprocessed:", {
      original: text.substring(0, 200),
      processed: processedText.substring(0, 200),
    });
  }

  return (
    <div
      className={cn(
        "prose prose-sm dark:prose-invert max-w-none prose-pre:bg-slate-100 dark:prose-pre:bg-slate-900/90 prose-pre:text-slate-800 dark:prose-pre:text-slate-100 prose-pre:px-3 prose-pre:py-2.5 prose-pre:rounded-lg prose-pre:border prose-pre:border-slate-300 dark:prose-pre:border-slate-800/70",
        "prose-code:before:hidden prose-code:after:hidden prose-code:px-1.5 prose-code:py-0.5 prose-code:rounded prose-code:bg-slate-200/70 dark:prose-code:bg-slate-700/70",
        "prose-ul:mt-2 prose-ul:mb-2 prose-li:my-1 prose-p:my-the 3 prose-p:leading-relaxed",
        className,
      )}
    >
      <ReactMarkdown
        remarkPlugins={[
          remarkGfm,
          [remarkExcelRanges, { sheetLookup }],
          [remarkExcelSheetMentions, { sheetLookup }],
        ]}
        components={components}
      >
        {processedText}
      </ReactMarkdown>
    </div>
  );
};
