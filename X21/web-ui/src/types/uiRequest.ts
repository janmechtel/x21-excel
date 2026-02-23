export type UiRequestMode = "blocking";

export type UiRequestControlKind =
  | "boolean"
  | "segmented"
  | "multi_choice"
  | "range_picker"
  | "text"
  | "folder_picker";

interface BaseUiControl {
  id: string;
  kind: UiRequestControlKind;
  label: string;
  required?: boolean;
  helperText?: string;
}

export interface BooleanControl extends BaseUiControl {
  kind: "boolean";
  yesLabel?: string;
  noLabel?: string;
}

export interface SegmentedOption {
  id: string;
  label: string;
  allowFreeText?: boolean;
}

export interface SegmentedControl extends BaseUiControl {
  kind: "segmented";
  options: SegmentedOption[];
}

export interface MultiChoiceControl extends BaseUiControl {
  kind: "multi_choice";
  options: SegmentedOption[];
}

export interface RangePickerPresetOption {
  id: string;
  label: string;
}

export interface RangePickerControl extends BaseUiControl {
  kind: "range_picker";
  presetOptions?: RangePickerPresetOption[];
}

export interface TextControl extends BaseUiControl {
  kind: "text";
  inputType?: "text" | "number";
  placeholder?: string;
}

export interface FolderPickerControl extends BaseUiControl {
  kind: "folder_picker";
  description?: string;
  showFileList?: boolean;
  extensions?: string[];
}

export type UiRequestControl =
  | BooleanControl
  | SegmentedControl
  | MultiChoiceControl
  | RangePickerControl
  | TextControl
  | FolderPickerControl;

export interface UiRequestPayload {
  title: string;
  description?: string;
  mode: UiRequestMode;
  controls: UiRequestControl[];
}

export type UiRequestResponseValue =
  | { value: boolean }
  | { choiceId: string; freeText?: string }
  | { choiceIds: string[]; freeText?: string }
  | { choiceId: string; rangeAddress?: string }
  | { text: string }
  | { path: string; files?: string[] };

export type UiRequestResponse = Record<string, UiRequestResponseValue>;
