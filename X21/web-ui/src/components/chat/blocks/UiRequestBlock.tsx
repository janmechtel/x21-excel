import { useEffect, useMemo, useState } from "react";

import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { webViewBridge } from "@/services/webViewBridge";
import type { ContentBlock } from "@/types/chat";
import type {
  RangePickerControl,
  SegmentedOption,
  UiRequestControl,
  UiRequestPayload,
  UiRequestResponse,
  FolderPickerControl,
} from "@/types/uiRequest";

interface UiRequestBlockProps {
  block: ContentBlock;
  selectedRange: string;
  onSubmit: (
    toolUseId: string,
    response: UiRequestResponse,
    summary?: string,
  ) => Promise<void>;
}

type ResponseState = Record<string, any>;

const inputClass =
  "w-full border border-slate-200 dark:border-slate-700 px-3 py-2 text-sm text-slate-900 dark:text-slate-100 placeholder:text-slate-400 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500/30 shadow-sm";

export function UiRequestBlock({
  block,
  selectedRange,
  onSubmit,
}: UiRequestBlockProps) {
  const request = block.uiRequest as UiRequestPayload | undefined;
  const initialResponse = block.uiRequestResponse;
  const [responses, setResponses] = useState<ResponseState>(
    initialResponse || {},
  );
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [submittedSummary, setSubmittedSummary] = useState(
    block.uiRequestSummary,
  );

  useEffect(() => {
    if (initialResponse) {
      setResponses(initialResponse);
      setSubmittedSummary(block.uiRequestSummary);
    }
  }, [initialResponse, block.uiRequestSummary]);

  if (!request || !block.toolUseId) return null;

  const setControlResponse = (controlId: string, value: any) => {
    setResponses((prev) => {
      return {
        ...prev,
        [controlId]: value,
      };
    });
  };

  const visibleControls = useMemo(() => {
    const hasFolderPicker = request.controls.some(
      (c) => c.kind === "folder_picker",
    );
    return hasFolderPicker
      ? request.controls.filter((c) => c.kind === "folder_picker")
      : request.controls;
  }, [request.controls]);

  const hasSingleFolderPicker = useMemo(
    () =>
      visibleControls.length === 1 &&
      visibleControls[0]?.kind === "folder_picker",
    [visibleControls],
  );

  const isFormValid = useMemo(() => {
    // Check if all required fields have valid responses
    return request.controls.every((control) => {
      if (!control.required) return true;

      const answer = responses[control.id];
      if (!answer) return false;

      switch (control.kind) {
        case "boolean":
          return typeof answer.value === "boolean";
        case "segmented":
          return Boolean((answer as any).choiceId);
        case "multi_choice":
          return (
            Array.isArray((answer as any).choiceIds) &&
            (answer as any).choiceIds.length > 0
          );
        case "range_picker": {
          const rangeAnswer = answer as any;
          return Boolean(rangeAnswer.choiceId && rangeAnswer.rangeAddress);
        }
        case "folder_picker": {
          const folderAnswer = answer as any;
          return Boolean(folderAnswer.path && folderAnswer.path.trim());
        }
        case "text": {
          const textAnswer = answer as any;
          return Boolean(textAnswer.text && textAnswer.text.trim());
        }
        default:
          return false;
      }
    });
  }, [request.controls, responses]);

  const buildSummary = (resp: UiRequestResponse) => {
    const parts: string[] = [];
    request.controls.forEach((control) => {
      const answer = resp[control.id];
      if (!answer) return;

      switch (control.kind) {
        case "boolean":
          parts.push(
            `${control.label}: ${(answer as any).value ? "Yes" : "No"}`,
          );
          break;
        case "segmented": {
          const choiceId = (answer as any).choiceId;
          const freeText = (answer as any).freeText;
          const optionLabel =
            (control as any).options?.find(
              (o: SegmentedOption) => o.id === choiceId,
            )?.label || choiceId;
          parts.push(
            `${control.label}: ${optionLabel}${
              freeText ? ` (${freeText})` : ""
            }`,
          );
          break;
        }
        case "multi_choice": {
          const choiceIds = (answer as any).choiceIds || [];
          const labels = choiceIds
            .map(
              (id: string) =>
                (control as any).options?.find(
                  (o: SegmentedOption) => o.id === id,
                )?.label || id,
            )
            .filter(Boolean);
          const freeText = (answer as any).freeText;
          parts.push(
            `${control.label}: ${labels.join(", ") || "None"}${
              freeText ? ` (${freeText})` : ""
            }`,
          );
          break;
        }
        case "range_picker": {
          const choiceId = (answer as any).choiceId;
          const range = (answer as any).rangeAddress;
          const label =
            (control as RangePickerControl).presetOptions?.find(
              (o) => o.id === choiceId,
            )?.label || choiceId;
          parts.push(`${control.label}: ${label}${range ? ` (${range})` : ""}`);
          break;
        }
        case "folder_picker": {
          const path = (answer as any).path;
          const files = (answer as any).files;
          const fileCount = Array.isArray(files) ? files.length : 0;
          const fileInfo =
            fileCount > 0
              ? ` (${fileCount} file${fileCount === 1 ? "" : "s"})`
              : "";
          parts.push(
            `${control.label}: ${path || "No folder selected"}${fileInfo}`,
          );
          break;
        }
        case "text":
          parts.push(`${control.label}: ${(answer as any).text || ""}`);
          break;
        default:
          break;
      }
    });
    return parts.filter(Boolean).join("\n");
  };

  const handleSubmit = async () => {
    if (!block.toolUseId) return;
    const payload = responses as UiRequestResponse;
    console.log("[UiRequestBlock] Submitting UI request", {
      toolUseId: block.toolUseId,
      payload,
    });
    setIsSubmitting(true);
    setError(null);
    try {
      const summary = buildSummary(payload);
      await onSubmit(block.toolUseId, payload, summary);
      setSubmittedSummary(summary);
    } catch (err) {
      console.error("[UiRequestBlock] Failed to submit UI request", err);
      const message =
        err instanceof Error ? err.message : "Failed to send form response";
      setError(message);
    } finally {
      setIsSubmitting(false);
    }
  };

  // Auto-submit when a single folder picker is present and filled
  useEffect(() => {
    if (!hasSingleFolderPicker) return;
    if (isSubmitting || initialResponse || submittedSummary) return;
    const control = visibleControls[0];
    const path = responses[control.id]?.path;
    if (typeof path === "string" && path.trim().length > 0) {
      void handleSubmit();
    }
  }, [
    hasSingleFolderPicker,
    isSubmitting,
    initialResponse,
    submittedSummary,
    responses,
    visibleControls,
    handleSubmit,
  ]);

  const handleRangeSelect = async (controlId: string) => {
    try {
      const picked = await webViewBridge.requestRangeSelection();
      if (picked) {
        setControlResponse(controlId, {
          choiceId: "custom",
          rangeAddress: picked,
        });
      }
    } catch {
      // ignore failures; UI will stay unchanged
    }
  };

  const renderBooleanControl = (control: UiRequestControl) => {
    const answer = responses[control.id];
    const yesActive = answer?.value === true;
    const noActive = answer?.value === false;
    const yesLabel = (control as any).yesLabel || "Yes";
    const noLabel = (control as any).noLabel || "No";

    return (
      <div className="flex gap-2">
        <Button
          type="button"
          variant={yesActive ? "default" : "outline"}
          size="sm"
          className={`whitespace-normal break-words text-left leading-tight h-auto min-h-[2rem] ${
            yesActive
              ? "bg-blue-600 text-white hover:bg-blue-700 border-blue-600"
              : ""
          }`}
          onClick={() => setControlResponse(control.id, { value: true })}
          disabled={isSubmitting || !!initialResponse}
        >
          {yesLabel}
        </Button>
        <Button
          type="button"
          variant={noActive ? "default" : "outline"}
          size="sm"
          className={`whitespace-normal break-words text-left leading-tight h-auto min-h-[2rem] ${
            noActive
              ? "bg-blue-600 text-white hover:bg-blue-700 border-blue-600"
              : ""
          }`}
          onClick={() => setControlResponse(control.id, { value: false })}
          disabled={isSubmitting || !!initialResponse}
        >
          {noLabel}
        </Button>
      </div>
    );
  };

  const renderOptions = (control: UiRequestControl, multiple = false) => {
    const answer = responses[control.id] || (multiple ? { choiceIds: [] } : {});
    const selectedIds: string[] = multiple
      ? answer.choiceIds || []
      : [answer.choiceId].filter(Boolean);

    return (
      <div className="flex flex-wrap gap-2">
        {(control as any).options?.map((option: SegmentedOption) => {
          const isSelected = multiple
            ? selectedIds.includes(option.id)
            : selectedIds[0] === option.id;
          return (
            <Button
              key={option.id}
              type="button"
              variant={isSelected ? "default" : "outline"}
              size="sm"
              className={`w-full sm:w-auto justify-start text-left leading-snug whitespace-normal break-words overflow-hidden ${
                isSelected
                  ? "bg-blue-600 text-white hover:bg-blue-700 border-blue-600"
                  : ""
              }`}
              onClick={() => {
                if (multiple) {
                  const next = isSelected
                    ? selectedIds.filter((id) => id !== option.id)
                    : [...selectedIds, option.id];
                  setControlResponse(control.id, {
                    choiceIds: next,
                    freeText: answer.freeText,
                  });
                } else {
                  setControlResponse(control.id, {
                    choiceId: option.id,
                    freeText: answer.freeText,
                  });
                }
              }}
              disabled={isSubmitting || !!initialResponse}
            >
              {option.label}
            </Button>
          );
        })}
      </div>
    );
  };

  const renderRangePicker = (control: RangePickerControl) => {
    const answer = responses[control.id] || {};
    const isSubmitted = !!initialResponse;
    const customRange = answer.rangeAddress || "";

    return (
      <div className="space-y-2">
        {control.presetOptions && control.presetOptions.length > 0 && (
          <div className="flex flex-wrap gap-2">
            {control.presetOptions.map((option) => {
              const isSelected = answer.choiceId === option.id;
              return (
                <Button
                  key={option.id}
                  type="button"
                  variant={isSelected ? "default" : "outline"}
                  size="sm"
                  className={`w-full sm:w-auto ${
                    isSelected
                      ? "bg-blue-600 text-white hover:bg-blue-700 border-blue-600"
                      : ""
                  }`}
                  onClick={() =>
                    setControlResponse(control.id, {
                      choiceId: option.id,
                      rangeAddress:
                        option.id === "selection" && selectedRange
                          ? selectedRange
                          : answer.rangeAddress,
                    })
                  }
                  disabled={isSubmitting || isSubmitted}
                >
                  {option.label}
                </Button>
              );
            })}
            {selectedRange && (
              <Button
                type="button"
                variant={
                  answer.choiceId === "selection" ? "default" : "outline"
                }
                size="sm"
                className={`w-full sm:w-auto ${
                  answer.choiceId === "selection"
                    ? "bg-blue-600 text-white hover:bg-blue-700 border-blue-600"
                    : ""
                }`}
                onClick={() =>
                  setControlResponse(control.id, {
                    choiceId: "selection",
                    rangeAddress: selectedRange,
                  })
                }
                disabled={isSubmitting || isSubmitted}
              >
                Use current selection
              </Button>
            )}
          </div>
        )}
        <div className="flex flex-col gap-2">
          <input
            type="text"
            className={inputClass}
            placeholder="e.g. Sheet1!B2:G20"
            value={customRange}
            onChange={(e) =>
              setControlResponse(control.id, {
                choiceId: "custom",
                rangeAddress: e.target.value,
              })
            }
            disabled={isSubmitting || isSubmitted}
          />
          <div className="flex gap-2 flex-col sm:flex-row">
            <Button
              type="button"
              variant="outline"
              size="sm"
              className="text-blue-700 border-blue-200 w-full sm:w-auto"
              onClick={() => void handleRangeSelect(control.id)}
              disabled={isSubmitting || isSubmitted}
            >
              Select range in Excel
            </Button>
            <Button
              type="button"
              variant={answer.choiceId === "custom" ? "default" : "outline"}
              size="sm"
              className={`w-full sm:w-auto ${
                answer.choiceId === "custom"
                  ? "bg-blue-600 text-white hover:bg-blue-700 border-blue-600"
                  : ""
              }`}
              onClick={() =>
                setControlResponse(control.id, {
                  choiceId: "custom",
                  rangeAddress: customRange || selectedRange,
                })
              }
              disabled={isSubmitting || isSubmitted}
            >
              Use custom range
            </Button>
          </div>
        </div>
      </div>
    );
  };

  const renderFolderPickerControl = (control: FolderPickerControl) => {
    const answer = responses[control.id] || {};
    const isSubmitted = !!initialResponse;
    const path = answer.path || "";
    const files: string[] = Array.isArray(answer.files)
      ? answer.files.filter((name: string) => !name.startsWith("~$"))
      : [];
    const showFileList = control.showFileList ?? true;

    const handleBrowse = async () => {
      try {
        setError(null);
        const picked = await webViewBridge.pickFolder({
          allowFileListing: showFileList,
          extensions: control.extensions,
        });
        if (picked?.path) {
          setControlResponse(control.id, {
            path: picked.path,
            files: showFileList ? picked.files : undefined,
          });
        }
      } catch (err) {
        const message =
          err instanceof Error ? err.message : "Unable to select a folder";
        setError(message);
      }
    };

    return (
      <div className="space-y-2">
        <div className="flex flex-col sm:flex-row gap-2 items-start">
          <Button
            type="button"
            variant="outline"
            size="sm"
            className="w-full sm:w-auto bg-blue-50 text-blue-700 border-blue-200"
            onClick={() => void handleBrowse()}
            disabled={isSubmitting || isSubmitted}
          >
            Choose folder
          </Button>
          {path ? (
            <div className="text-xs text-slate-700 dark:text-slate-200 break-all">
              {path}
            </div>
          ) : (
            <div className="text-xs text-slate-500 dark:text-slate-400">
              {control.description ||
                "Select the folder containing Excel files to merge"}
            </div>
          )}
        </div>
        {showFileList && files.length > 0 && (
          <div className="rounded-md border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/50 p-2 text-[11px] text-slate-700 dark:text-slate-200">
            <div className="font-semibold mb-1">
              Found files ({files.length})
            </div>
            <div className="max-h-32 overflow-y-auto space-y-0.5">
              {files.map((file) => (
                <div key={file} className="truncate">
                  {file}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  };

  const renderTextControl = (control: UiRequestControl) => {
    const answer = responses[control.id] || {};
    if ((control as any).inputType === "number") {
      return (
        <input
          type="number"
          inputMode="numeric"
          pattern="[0-9]*"
          className={inputClass}
          value={answer.text || ""}
          placeholder={(control as any).placeholder || "Enter a number"}
          onChange={(e) =>
            setControlResponse(control.id, { text: e.target.value })
          }
          disabled={isSubmitting || !!initialResponse}
        />
      );
    }
    return (
      <Textarea
        value={answer.text || ""}
        placeholder={(control as any).placeholder || "Type your response"}
        onChange={(e) =>
          setControlResponse(control.id, { text: e.target.value })
        }
        className="min-h-[64px]"
        disabled={isSubmitting || !!initialResponse}
      />
    );
  };

  const renderControl = (control: UiRequestControl) => {
    switch (control.kind) {
      case "boolean":
        return renderBooleanControl(control);
      case "segmented":
        return (
          <div className="space-y-2">
            {renderOptions(control)}
            {(() => {
              const selectedOption = (control as any).options?.find(
                (opt: SegmentedOption) =>
                  opt.id === responses[control.id]?.choiceId,
              );
              const showFreeText =
                selectedOption?.allowFreeText ||
                Boolean(responses[control.id]?.freeText);
              if (!showFreeText) return null;
              return (
                <Textarea
                  placeholder="Add details (optional)"
                  value={responses[control.id]?.freeText || ""}
                  onChange={(e) =>
                    setControlResponse(control.id, {
                      ...responses[control.id],
                      freeText: e.target.value,
                    })
                  }
                  className="min-h-[48px]"
                  disabled={isSubmitting || !!initialResponse}
                />
              );
            })()}
          </div>
        );
      case "multi_choice": {
        const selectedIds = responses[control.id]?.choiceIds || [];
        const hasFreeText = (control as any).options?.some(
          (opt: SegmentedOption) =>
            opt.allowFreeText && selectedIds.includes(opt.id),
        );
        return (
          <div className="space-y-2">
            {renderOptions(control, true)}
            {hasFreeText || responses[control.id]?.freeText ? (
              <Textarea
                placeholder="Tell us more (optional)"
                value={responses[control.id]?.freeText || ""}
                onChange={(e) =>
                  setControlResponse(control.id, {
                    ...responses[control.id],
                    freeText: e.target.value,
                  })
                }
                className="min-h-[48px]"
                disabled={isSubmitting || !!initialResponse}
              />
            ) : null}
          </div>
        );
      }
      case "range_picker":
        return renderRangePicker(control as RangePickerControl);
      case "folder_picker":
        return renderFolderPickerControl(control as FolderPickerControl);
      case "text":
        return renderTextControl(control);
      default:
        return null;
    }
  };

  const showSummary = submittedSummary || block.uiRequestResponse;

  const renderSummary = () => {
    const summary = submittedSummary || block.uiRequestSummary;
    if (summary) return summary;

    // Fallback: build a readable multi-line summary
    const resp = block.uiRequestResponse as UiRequestResponse | undefined;
    if (!resp || !request) return "Response submitted";

    const lines: string[] = [];
    request.controls.forEach((control) => {
      const answer = resp[control.id];
      if (!answer) return;

      switch (control.kind) {
        case "boolean":
          lines.push(
            `${control.label}: ${(answer as any).value ? "Yes" : "No"}`,
          );
          break;
        case "segmented": {
          const choiceId = (answer as any).choiceId;
          const freeText = (answer as any).freeText;
          const optionLabel =
            (control as any).options?.find(
              (o: SegmentedOption) => o.id === choiceId,
            )?.label || choiceId;
          lines.push(
            `${control.label}: ${optionLabel}${
              freeText ? ` (${freeText})` : ""
            }`,
          );
          break;
        }
        case "multi_choice": {
          const choiceIds = (answer as any).choiceIds || [];
          const labels = choiceIds
            .map(
              (id: string) =>
                (control as any).options?.find(
                  (o: SegmentedOption) => o.id === id,
                )?.label || id,
            )
            .filter(Boolean);
          const freeText = (answer as any).freeText;
          lines.push(
            `${control.label}: ${labels.join(", ") || "None"}${
              freeText ? ` (${freeText})` : ""
            }`,
          );
          break;
        }
        case "range_picker": {
          const choiceId = (answer as any).choiceId;
          const range = (answer as any).rangeAddress;
          const label =
            (control as RangePickerControl).presetOptions?.find(
              (o) => o.id === choiceId,
            )?.label || choiceId;
          lines.push(`${control.label}: ${label}${range ? ` (${range})` : ""}`);
          break;
        }
        case "folder_picker": {
          const path = (answer as any).path;
          const files = (answer as any).files;
          const count = Array.isArray(files) ? files.length : 0;
          lines.push(
            `${control.label}: ${path || "No folder selected"}${
              count ? ` (${count} file${count === 1 ? "" : "s"})` : ""
            }`,
          );
          break;
        }
        case "text":
          lines.push(`${control.label}: ${(answer as any).text || ""}`);
          break;
        default:
          break;
      }
    });

    return lines.join("\n");
  };

  return (
    <div className="scroll-mt-1" data-testid="ui-request-card">
      <div className="pt-1 pb-1">
        <div className="text-base font-semibold text-slate-900 dark:text-slate-100">
          {request.title}
        </div>
        {request.description && (
          <p className="text-xs text-slate-600 dark:text-slate-400 mt-1">
            {request.description}
          </p>
        )}
      </div>
      <div className="pb-3 space-y-4">
        {visibleControls.map((control, index) => (
          <div
            key={control.id}
            className={`space-y-2 ${
              index !== visibleControls.length - 1
                ? "pb-2 border-b border-slate-100 dark:border-slate-800"
                : ""
            }`}
          >
            <div className="flex items-center gap-2">
              <p className="text-sm font-semibold text-slate-900 dark:text-slate-100 leading-tight whitespace-normal break-words">
                {control.label}
              </p>
              {control.required && (
                <span
                  className="text-[10px] font-semibold text-red-600"
                  aria-label="Required"
                >
                  *
                </span>
              )}
            </div>
            {"helperText" in control && (control as any).helperText ? (
              <p className="text-[11px] text-slate-500 dark:text-slate-400 leading-snug whitespace-normal break-words">
                {(control as any).helperText}
              </p>
            ) : null}
            {renderControl(control)}
          </div>
        ))}

        {showSummary ? (
          <div className="border border-emerald-200 dark:border-emerald-800/60 px-1 py-2 text-xs text-emerald-900 dark:text-emerald-200 whitespace-pre-line">
            {renderSummary()}
          </div>
        ) : (
          <div className="flex items-center gap-2 pt-2">
            <Button
              onClick={() => void handleSubmit()}
              disabled={isSubmitting || !isFormValid}
              size="sm"
            >
              {isSubmitting ? "Submitting..." : "Continue"}
            </Button>
            {error && (
              <span className="text-xs text-red-600 dark:text-red-400">
                {error}
              </span>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
