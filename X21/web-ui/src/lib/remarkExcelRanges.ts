import type { Plugin } from "unified";
import type { Link, Parent, Root, Text } from "mdast";
import { visit } from "unist-util-visit";
import {
  type SheetLookup,
  isValidSheetNameFormat,
  isKnownSheetName,
} from "@/utils/excelSheetValidation";

const EXCEL_RANGE_REGEX =
  /(?:\[[^\]]+\][^!]+|(?:'[^']*'|[A-Za-z0-9_]+)!(?:'[^']*'|[A-Za-z0-9_]+)|(?:'[^']*'|[A-Za-z0-9_]+))!\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?/g;
const EXCEL_SHEET_REGEX =
  /(?:\[[^\]]+\])?(?:'[^']*'|[A-Za-z0-9_.-]+(?:\s+[A-Za-z0-9_.-]+)*)!(?=$|\s|[.,;:)\]])/g;

type TextNode = Text & { value: string };

const shouldSkipParent = (parent?: Parent | null) => {
  if (!parent) return true;
  return (
    parent.type === "link" ||
    parent.type === "inlineCode" ||
    parent.type === "code"
  );
};

type ExcelRangePluginOptions = {
  knownSheets?: Set<string> | string[];
  sheetLookup?: SheetLookup;
};

const normalizeKnownSheets = (raw?: Set<string> | string[]): Set<string> => {
  if (!raw) return new Set();
  if (raw instanceof Set) return raw;
  return new Set(raw.map((name) => name.toLowerCase()));
};

export const remarkExcelRanges: Plugin<[ExcelRangePluginOptions?], Root> = (
  options = {},
) => {
  const knownSheets = normalizeKnownSheets(options.knownSheets);
  const sheetLookup: SheetLookup = options.sheetLookup ?? {
    currentSheets: knownSheets,
    workbookSheets: new Map<string, Set<string>>(),
  };
  return (tree) => {
    visit(tree, "text", (node: TextNode, index, parent) => {
      if (!parent || typeof index !== "number" || shouldSkipParent(parent)) {
        return;
      }

      const value = node.value || "";
      EXCEL_RANGE_REGEX.lastIndex = 0;
      let match: RegExpExecArray | null;
      const children: Parent["children"] = [];
      let cursor = 0;
      let hasMatch = false;

      while ((match = EXCEL_RANGE_REGEX.exec(value)) !== null) {
        const [range] = match;
        const bangIndex = range.indexOf("!");
        if (bangIndex < 0) {
          continue;
        }

        let sheetPart = range.slice(0, bangIndex);
        let workbookPart: string | undefined;
        const parts = range.split("!");
        if (parts.length >= 3) {
          workbookPart = parts[0];
          sheetPart = parts[1];
        } else {
          const bracketMatch = sheetPart.match(/^\[([^\]]+)\](.+)$/);
          if (bracketMatch) {
            workbookPart = bracketMatch[1];
            sheetPart = bracketMatch[2];
          } else if (sheetPart.includes("/")) {
            const lastSlash = sheetPart.lastIndexOf("/");
            workbookPart = sheetPart.slice(0, lastSlash);
            sheetPart = sheetPart.slice(lastSlash + 1);
          }
        }

        if (!sheetPart || !isValidSheetNameFormat(sheetPart)) {
          continue;
        }

        const hasSheetContext =
          sheetLookup.currentSheets.size > 0 ||
          sheetLookup.workbookSheets.size > 0;
        if (hasSheetContext) {
          if (!isKnownSheetName(sheetPart, sheetLookup, workbookPart)) {
            continue;
          }
        } else if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) {
          continue;
        }

        hasMatch = true;
        const start = match.index;

        if (start > cursor) {
          children.push({
            type: "text",
            value: value.slice(cursor, start),
          } as TextNode);
        }

        const linkNode: Link = {
          type: "link",
          url: `excel-range://${range}`,
          title: null,
          children: [
            {
              type: "text",
              value: range,
            },
          ],
        };

        children.push(linkNode);

        cursor = start + range.length;
      }

      if (!hasMatch) {
        return;
      }

      if (cursor < value.length) {
        children.push({
          type: "text",
          value: value.slice(cursor),
        } as TextNode);
      }

      parent.children.splice(index, 1, ...children);
      return index + children.length;
    });
  };
};

export const remarkExcelSheetMentions: Plugin<
  [ExcelRangePluginOptions?],
  Root
> = (options = {}) => {
  const knownSheets = normalizeKnownSheets(options.knownSheets);
  const sheetLookup: SheetLookup = options.sheetLookup ?? {
    currentSheets: knownSheets,
    workbookSheets: new Map<string, Set<string>>(),
  };
  return (tree) => {
    visit(tree, "text", (node: TextNode, index, parent) => {
      if (!parent || typeof index !== "number" || shouldSkipParent(parent)) {
        return;
      }

      const value = node.value || "";
      EXCEL_SHEET_REGEX.lastIndex = 0;
      let match: RegExpExecArray | null;
      const children: Parent["children"] = [];
      let cursor = 0;
      let hasMatch = false;

      while ((match = EXCEL_SHEET_REGEX.exec(value)) !== null) {
        const [token] = match;
        const bangIndex = token.indexOf("!");
        if (bangIndex < 0) {
          continue;
        }

        let sheetPart = token.slice(0, bangIndex);
        let workbookPart: string | undefined;
        const bracketMatch = sheetPart.match(/^\[([^\]]+)\](.+)$/);
        if (bracketMatch) {
          workbookPart = bracketMatch[1];
          sheetPart = bracketMatch[2];
        } else if (sheetPart.includes("/")) {
          const lastSlash = sheetPart.lastIndexOf("/");
          workbookPart = sheetPart.slice(0, lastSlash);
          sheetPart = sheetPart.slice(lastSlash + 1);
        }

        if (!sheetPart || !isValidSheetNameFormat(sheetPart)) {
          continue;
        }

        const hasSheetContext =
          sheetLookup.currentSheets.size > 0 ||
          sheetLookup.workbookSheets.size > 0;
        if (hasSheetContext) {
          if (!isKnownSheetName(sheetPart, sheetLookup, workbookPart)) {
            continue;
          }
        } else if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) {
          continue;
        }

        hasMatch = true;
        const start = match.index;

        if (start > cursor) {
          children.push({
            type: "text",
            value: value.slice(cursor, start),
          } as TextNode);
        }

        const linkNode: Link = {
          type: "link",
          url: `excel-sheet://${token}`,
          title: null,
          children: [
            {
              type: "text",
              value: token,
            },
          ],
        };

        children.push(linkNode);
        cursor = start + token.length;
      }

      if (!hasMatch) {
        return;
      }

      if (cursor < value.length) {
        children.push({
          type: "text",
          value: value.slice(cursor),
        } as TextNode);
      }

      parent.children.splice(index, 1, ...children);
      return index + children.length;
    });
  };
};
