export const ColumnWidthModes = {
  ALWAYS: "always",
  SMART: "smart",
  NEVER: "never",
} as const;

export type ColumnWidthMode =
  (typeof ColumnWidthModes)[keyof typeof ColumnWidthModes];

export const getColumnWidthMessage = (mode: ColumnWidthMode) => {
  if (mode === ColumnWidthModes.NEVER) {
    return "Column-width fitting is off";
  }
  if (mode === ColumnWidthModes.ALWAYS) {
    return "Column-width fitted automatically";
  }
  return "Auto-fit default-width columns only";
};
