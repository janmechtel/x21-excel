import type {
  ClipboardEventHandler,
  DragEventHandler,
  KeyboardEventHandler,
  RefObject,
} from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { LexicalComposer } from "@lexical/react/LexicalComposer";
import { PlainTextPlugin } from "@lexical/react/LexicalPlainTextPlugin";
import { ContentEditable } from "@lexical/react/LexicalContentEditable";
import { HistoryPlugin } from "@lexical/react/LexicalHistoryPlugin";
import { OnChangePlugin } from "@lexical/react/LexicalOnChangePlugin";
import { useLexicalComposerContext } from "@lexical/react/LexicalComposerContext";
import { LexicalErrorBoundary } from "@lexical/react/LexicalErrorBoundary";
import { $createLineBreakNode, $getRoot, $getSelection } from "lexical";
import {
  $createParagraphNode,
  $createTextNode,
  $isElementNode,
  $isLineBreakNode,
  $isRangeSelection,
  $isTextNode,
  COMMAND_PRIORITY_LOW,
  DecoratorNode,
  EditorState,
  KEY_BACKSPACE_COMMAND,
  KEY_DELETE_COMMAND,
  SELECTION_CHANGE_COMMAND,
} from "lexical";

import { ExcelRangePill } from "@/components/excel/ExcelRangePill";
import { ExcelSheetMentionPill } from "@/components/excel/ExcelSheetMentionPill";
import type { PromptSegment } from "@/components/chat/ChatInputForm";

type RichChatEditorProps = {
  promptSegments: PromptSegment[];
  placeholder: string;
  disabled: boolean;
  hasActiveSlashCommand: boolean;
  textareaRef: RefObject<HTMLTextAreaElement>;
  onRangePillClick: (range: string) => void;
  onSheetMentionClick: (sheetName: string) => void;
  onPromptChange: (event: React.ChangeEvent<HTMLTextAreaElement>) => void;
  onPromptKeyDown: KeyboardEventHandler<HTMLTextAreaElement>;
  onPaste: ClipboardEventHandler<HTMLTextAreaElement>;
  onDragOver: DragEventHandler;
  onDragEnter: DragEventHandler;
  onDragLeave: DragEventHandler;
  onDrop: DragEventHandler;
};

// Simple placeholder chip for sheet mentions to differentiate from ranges.
function SheetMentionPill({ value }: { value: string }) {
  return <ExcelSheetMentionPill sheetName={value} />;
}

class RangeNode extends DecoratorNode<JSX.Element> {
  __value: string;

  static getType(): string {
    return "range-pill";
  }

  static clone(node: RangeNode) {
    return new RangeNode(node.__value, node.__key);
  }

  constructor(value: string, key?: string) {
    super(key);
    this.__value = value;
  }

  createDOM(): HTMLElement {
    const span = document.createElement("span");
    span.className = "inline-flex";
    return span;
  }

  decorate(): JSX.Element {
    return (
      <span
        data-lexical-range-value={this.__value}
        className="inline-flex cursor-pointer"
      >
        <ExcelRangePill
          range={this.__value}
          className="my-0.5 mx-0.5"
          variant="minimal"
        />
      </span>
    );
  }

  isInline(): boolean {
    return true;
  }

  // Let Lexical treat the node as text for export/selection math.
  getTextContent(): string {
    return this.__value;
  }

  exportJSON() {
    return {
      type: RangeNode.getType(),
      version: 1,
      value: this.__value,
    };
  }

  static importJSON(serializedNode: any): RangeNode {
    return new RangeNode(serializedNode.value);
  }

  getValue(): string {
    return this.__value;
  }
}

class SheetMentionNode extends DecoratorNode<JSX.Element> {
  __value: string;

  static getType(): string {
    return "sheet-mention";
  }

  static clone(node: SheetMentionNode) {
    return new SheetMentionNode(node.__value, node.__key);
  }

  constructor(value: string, key?: string) {
    super(key);
    this.__value = value;
  }

  createDOM(): HTMLElement {
    const span = document.createElement("span");
    span.className = "inline-flex";
    return span;
  }

  decorate(): JSX.Element {
    const rawValue = this.__value.replace(/^@/, "").replace(/!$/, "").trim();
    const bracketMatch = rawValue.match(/^\[([^\]]+)\](.+)$/);
    const workbookName = bracketMatch?.[1];
    const sheetToken = bracketMatch?.[2] ?? rawValue;
    const cleanedSheet = sheetToken.replace(/^'+|'+$/g, "").replace(/''/g, "'");
    const displayValue = workbookName
      ? `${workbookName}/${cleanedSheet}`
      : cleanedSheet;
    return (
      <span
        data-lexical-sheet-name={rawValue}
        className="inline-flex cursor-pointer"
      >
        <SheetMentionPill value={displayValue} />
      </span>
    );
  }

  isInline(): boolean {
    return true;
  }

  getTextContent(): string {
    return this.__value;
  }

  exportJSON() {
    return {
      type: SheetMentionNode.getType(),
      version: 1,
      value: this.__value,
    };
  }

  static importJSON(serializedNode: any): SheetMentionNode {
    return new SheetMentionNode(serializedNode.value);
  }

  getValue(): string {
    return this.__value;
  }
}

const NODES = [RangeNode, SheetMentionNode];

const theme = {
  paragraph: "",
};

const placeholderClass =
  "absolute left-2 top-2 pointer-events-none text-xs text-slate-400 dark:text-slate-500";

function createEditorStateFromSegments(segments: PromptSegment[]) {
  return () => {
    const root = $getRoot();
    root.clear();
    const paragraph = $createParagraphNode();

    segments.forEach((segment) => {
      if (segment.type === "text") {
        const parts = segment.value.split(/\n/);
        parts.forEach((part, partIdx) => {
          if (part) {
            paragraph.append($createTextNode(part));
          }
          if (partIdx < parts.length - 1) {
            paragraph.append($createLineBreakNode());
          }
        });
        return;
      }

      if (segment.type === "range") {
        paragraph.append(new RangeNode(segment.value));
        return;
      }

      paragraph.append(new SheetMentionNode(segment.value));
    });

    if (paragraph.getChildrenSize() === 0) {
      paragraph.append($createTextNode(""));
    }

    root.append(paragraph);
    paragraph.selectEnd();
  };
}

function segmentsToPlainText(segments: PromptSegment[]): string {
  return segments.map((segment) => segment.value).join("");
}

function collectSegments(): PromptSegment[] {
  const root = $getRoot();
  const paragraph = root.getFirstChild();
  if (!$isElementNode(paragraph)) {
    return [{ type: "text", value: "" }] as PromptSegment[];
  }

  const children = paragraph.getChildren();
  const results: PromptSegment[] = [];
  let pendingText = "";

  const flushText = () => {
    if (pendingText.length === 0) return;
    if (results.length > 0 && results[results.length - 1].type === "text") {
      results[results.length - 1].value += pendingText;
    } else {
      results.push({ type: "text", value: pendingText });
    }
    pendingText = "";
  };

  children.forEach((node) => {
    if ($isTextNode(node)) {
      pendingText += node.getTextContent();
      return;
    }

    if ($isLineBreakNode(node)) {
      pendingText += "\n";
      return;
    }

    flushText();
    if (node instanceof RangeNode) {
      results.push({ type: "range", value: node.getValue() });
    } else if (node instanceof SheetMentionNode) {
      results.push({ type: "sheetMention", value: node.getValue() });
    }
  });

  flushText();
  return results.length ? results : [{ type: "text", value: "" }];
}

function getNodeLength(node: any): number {
  if ($isTextNode(node)) return node.getTextContentSize();
  if ($isLineBreakNode(node)) return 1;
  if (node instanceof RangeNode || node instanceof SheetMentionNode) {
    return node.getTextContent().length;
  }
  return 0;
}

function getSelectionOffset(): number | null {
  const selection = $getSelection();
  if (!$isRangeSelection(selection)) return null;

  const anchor = selection.anchor;
  const root = $getRoot();
  const paragraph = root.getFirstChild();
  if (!$isElementNode(paragraph)) return null;

  let offset = 0;
  const children = paragraph.getChildren();
  for (const node of children) {
    if (node.getKey() === anchor.key) {
      if ($isTextNode(node)) {
        offset += anchor.offset;
      } else if ($isLineBreakNode(node)) {
        offset += 1;
      } else if (
        node instanceof RangeNode ||
        node instanceof SheetMentionNode
      ) {
        offset += anchor.offset > 0 ? node.getTextContent().length : 0;
      }
      return offset;
    }

    offset += getNodeLength(node);
  }

  return offset;
}

function setSelectionAtOffset(editor: any, targetOffset: number) {
  editor.update(() => {
    const root = $getRoot();
    const paragraph = root.getFirstChild();
    if (!$isElementNode(paragraph)) return;

    let remaining = targetOffset;
    const children = paragraph.getChildren();

    for (const node of children) {
      const length = getNodeLength(node);
      if (remaining <= length) {
        if ($isTextNode(node)) {
          node.select(remaining, remaining);
          return;
        }

        if ($isLineBreakNode(node)) {
          node.selectPrevious();
          return;
        }

        if (node instanceof RangeNode || node instanceof SheetMentionNode) {
          if (remaining === 0) {
            node.selectPrevious();
          } else {
            node.selectNext();
          }
          return;
        }
      }

      remaining -= length;
    }

    paragraph.selectEnd();
  });
}

function EditableStatePlugin({ disabled }: { disabled: boolean }) {
  const [editor] = useLexicalComposerContext();
  useEffect(() => {
    editor.setEditable(!disabled);
  }, [disabled, editor]);
  return null;
}

function ExternalValueSyncPlugin({
  promptSegments,
  lastPlainTextRef,
}: {
  promptSegments: PromptSegment[];
  lastPlainTextRef: React.MutableRefObject<string>;
}) {
  const [editor] = useLexicalComposerContext();
  const serialized = useMemo(
    () => segmentsToPlainText(promptSegments),
    [promptSegments],
  );

  useEffect(() => {
    if (serialized === lastPlainTextRef.current) return;
    editor.update(() => {
      createEditorStateFromSegments(promptSegments)();
    });
    lastPlainTextRef.current = serialized;
  }, [editor, promptSegments, serialized, lastPlainTextRef]);

  return null;
}

function SelectionTrackerPlugin({
  selectionRef,
}: {
  selectionRef: React.MutableRefObject<number | null>;
}) {
  const [editor] = useLexicalComposerContext();

  useEffect(() => {
    return editor.registerCommand(
      SELECTION_CHANGE_COMMAND,
      () => {
        const offset = editor.getEditorState().read(() => getSelectionOffset());
        selectionRef.current = offset;
        return false;
      },
      COMMAND_PRIORITY_LOW,
    );
  }, [editor]);

  return null;
}

export function RichChatEditor({
  promptSegments,
  placeholder,
  disabled,
  hasActiveSlashCommand,
  textareaRef,
  onRangePillClick,
  onSheetMentionClick,
  onPromptChange,
  onPromptKeyDown,
  onPaste,
  onDragOver,
  onDragEnter,
  onDragLeave,
  onDrop,
}: RichChatEditorProps) {
  const selectionRef = useRef<number | null>(null);
  const lastPlainTextRef = useRef<string>(segmentsToPlainText(promptSegments));
  const styleRef = useRef<{ height?: string }>({});
  const contentEditableRef = useRef<HTMLDivElement | null>(null);
  const [editorKey, setEditorKey] = useState(0);

  const handleLexicalError = useCallback((error: Error) => {
    console.error("[RichChatEditor] Lexical error, resetting editor:", error);
    setEditorKey((prev) => prev + 1);
  }, []);

  const initialConfig = useMemo(
    () => ({
      namespace: "chat-rich-editor",
      nodes: NODES,
      theme,
      editable: !disabled,
      editorState: createEditorStateFromSegments(promptSegments),
      onError: handleLexicalError,
    }),
    [disabled, handleLexicalError, promptSegments],
  );

  const emitChange = useCallback(
    (text: string, selection: number | null) => {
      lastPlainTextRef.current = text;
      if (!onPromptChange) return;

      const syntheticEvent = {
        target: {
          value: text,
          selectionStart: selection ?? text.length,
          selectionEnd: selection ?? text.length,
        },
      } as unknown as React.ChangeEvent<HTMLTextAreaElement>;

      onPromptChange(syntheticEvent);
    },
    [onPromptChange],
  );

  const handleLexicalChange = useCallback(
    (editorState: EditorState) => {
      editorState.read(() => {
        const segments = collectSegments();
        const text = segmentsToPlainText(segments);
        const selection = getSelectionOffset();
        selectionRef.current = selection;
        emitChange(text, selection);
      });
    },
    [emitChange],
  );

  const updateTextareaProxy = useCallback(
    (editor: any) => {
      if (!textareaRef) return;
      const styleTarget =
        contentEditableRef.current?.style ??
        (styleRef.current as unknown as CSSStyleDeclaration);
      const listenerTarget = contentEditableRef.current;
      const proxy: any = {
        focus: () => editor.focus(),
        setSelectionRange: (start: number, _end?: number) => {
          setSelectionAtOffset(editor, start);
          selectionRef.current = start;
        },
        get selectionStart() {
          return selectionRef.current ?? 0;
        },
        get selectionEnd() {
          return selectionRef.current ?? 0;
        },
        get value() {
          return lastPlainTextRef.current;
        },
        style: styleTarget,
        disabled: disabled,
        hidden: false,
        get scrollTop() {
          return contentEditableRef.current?.scrollTop ?? 0;
        },
        set scrollTop(val: number) {
          if (contentEditableRef.current) {
            contentEditableRef.current.scrollTop = val;
          }
        },
        get scrollHeight() {
          return contentEditableRef.current?.scrollHeight ?? 0;
        },
        addEventListener: (
          ...args: Parameters<HTMLElement["addEventListener"]>
        ) => listenerTarget?.addEventListener(...args),
        removeEventListener: (
          ...args: Parameters<HTMLElement["removeEventListener"]>
        ) => listenerTarget?.removeEventListener(...args),
      };

      (textareaRef as unknown as React.MutableRefObject<any>).current = proxy;
    },
    [contentEditableRef, disabled, textareaRef],
  );

  return (
    <LexicalComposer initialConfig={initialConfig} key={editorKey}>
      <EditableStatePlugin disabled={disabled} />
      <SelectionTrackerPlugin selectionRef={selectionRef} />
      <ExternalValueSyncPlugin
        promptSegments={promptSegments}
        lastPlainTextRef={lastPlainTextRef}
      />
      <OnChangePlugin onChange={handleLexicalChange} />
      <HistoryPlugin />
      <PlainTextPlugin
        contentEditable={
          <ContentEditable
            ref={contentEditableRef}
            className={`resize-none border-none focus:ring-0 focus:outline-none flex-1 overflow-y-auto px-2 font-sans ${
              hasActiveSlashCommand ? "pt-1.5" : "pt-2"
            } pb-2 bg-transparent text-xs leading-5 ${
              disabled ? "cursor-not-allowed text-slate-500/80" : ""
            }`}
            style={{
              minHeight: "40px",
              scrollbarWidth: "none",
              msOverflowStyle: "none",
            }}
            onKeyDown={
              onPromptKeyDown as unknown as KeyboardEventHandler<HTMLDivElement>
            }
            onPaste={
              onPaste as unknown as ClipboardEventHandler<HTMLDivElement>
            }
            onDragOver={onDragOver}
            onDragEnter={onDragEnter}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            aria-label="Chat input"
          />
        }
        placeholder={<span className={placeholderClass}>{placeholder}</span>}
        ErrorBoundary={LexicalErrorBoundary}
      />
      <ProxyBridgePlugin
        updateProxy={updateTextareaProxy}
        contentEditableRef={contentEditableRef}
        selectionRef={selectionRef}
      />
      <RangeClickBridgePlugin
        onRangePillClick={onRangePillClick}
        onSheetMentionClick={onSheetMentionClick}
      />
      <DeletionPlugin />
    </LexicalComposer>
  );
}

function ProxyBridgePlugin({
  updateProxy,
  contentEditableRef,
  selectionRef,
}: {
  updateProxy: (editor: any) => void;
  contentEditableRef: React.MutableRefObject<HTMLDivElement | null>;
  selectionRef: React.MutableRefObject<number | null>;
}) {
  const [editor] = useLexicalComposerContext();
  useEffect(() => {
    updateProxy(editor);
  }, [editor, updateProxy, contentEditableRef, selectionRef]);
  return null;
}

function DeletionPlugin() {
  const [editor] = useLexicalComposerContext();

  useEffect(() => {
    const deleteNearbyNode = (isBackspace: boolean) => {
      const selection = $getSelection();
      if (!$isRangeSelection(selection)) return false;
      const anchor = selection.anchor;
      const root = $getRoot();
      const paragraph = root.getFirstChild();
      if (!$isElementNode(paragraph)) return false;

      const children = paragraph.getChildren();
      let caretOffset = 0;
      for (const node of children) {
        const length = getNodeLength(node);
        const isAnchorNode = node.getKey() === anchor.key;

        if (isAnchorNode) {
          const cursorPos = anchor.offset;
          if (isBackspace && cursorPos === 0) {
            const prevNode = node.getPreviousSibling();
            if (
              prevNode instanceof RangeNode ||
              prevNode instanceof SheetMentionNode
            ) {
              prevNode.remove();
              return true;
            }
          }
          if (!isBackspace && cursorPos === length) {
            const nextNode = node.getNextSibling();
            if (
              nextNode instanceof RangeNode ||
              nextNode instanceof SheetMentionNode
            ) {
              nextNode.remove();
              return true;
            }
          }
          return false;
        }

        caretOffset += length;
        if (anchor.offset === 0 && length > 0) {
          // When anchor is on decorator nodes themselves
          if (isBackspace) {
            if (node instanceof RangeNode || node instanceof SheetMentionNode) {
              node.remove();
              return true;
            }
          }
        }
      }
      return false;
    };

    const backspaceUnregister = editor.registerCommand(
      KEY_BACKSPACE_COMMAND,
      () => deleteNearbyNode(true),
      COMMAND_PRIORITY_LOW,
    );
    const deleteUnregister = editor.registerCommand(
      KEY_DELETE_COMMAND,
      () => deleteNearbyNode(false),
      COMMAND_PRIORITY_LOW,
    );

    return () => {
      backspaceUnregister();
      deleteUnregister();
    };
  }, [editor]);

  return null;
}

function RangeClickBridgePlugin({
  onRangePillClick,
  onSheetMentionClick,
}: {
  onRangePillClick: (range: string) => void;
  onSheetMentionClick: (sheetName: string) => void;
}) {
  const [editor] = useLexicalComposerContext();

  useEffect(() => {
    const clickHandler = (event: MouseEvent) => {
      const target = event.target as HTMLElement | null;
      if (!target) return;
      const sheet = target.closest("[data-lexical-sheet-name]");
      if (sheet && sheet instanceof HTMLElement) {
        const value = sheet.dataset.lexicalSheetName;
        if (value) {
          event.preventDefault();
          event.stopPropagation();
          onSheetMentionClick(value);
          return;
        }
      }
      const pill = target.closest("[data-lexical-range-value]");
      if (pill && pill instanceof HTMLElement) {
        const value = pill.dataset.lexicalRangeValue;
        if (value) {
          event.preventDefault();
          event.stopPropagation();
          onRangePillClick(value);
        }
      }
    };

    return editor.registerRootListener(
      (
        rootElement: HTMLElement | null,
        prevRootElement: HTMLElement | null,
      ) => {
        if (prevRootElement) {
          prevRootElement.removeEventListener("click", clickHandler, true);
        }
        if (rootElement) {
          rootElement.addEventListener("click", clickHandler, true);
        }
      },
    );
  }, [editor, onRangePillClick, onSheetMentionClick]);

  return null;
}
