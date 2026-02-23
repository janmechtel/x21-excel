namespace X21.Constants
{
    /// <summary>
    /// Tool name constants shared with TypeScript codebase.
    /// These values must match X21/shared/types/constants.ts ToolNames.
    ///
    /// Note: Only includes tools that are actually used in the C# codebase.
    /// Tools like collect_input, list_sheets, get_metadata, and list_open_workbooks
    /// are handled differently (GET endpoints or TypeScript-only) and don't need constants here.
    /// </summary>
    public static class ExcelToolNames
    {
        public const string ReadValuesBatch = "read_values_batch";
        public const string WriteValuesBatch = "write_values_batch";
        public const string WriteFormatBatch = "write_format_batch";
        public const string DragFormula = "drag_formula";
        public const string AddSheets = "add_sheets";
        public const string RemoveSheets = "remove_sheets";
        public const string AddColumns = "add_columns";
        public const string RemoveColumns = "remove_columns";
        public const string AddRows = "add_rows";
        public const string RemoveRows = "remove_rows";
        public const string VbaCreate = "vba_create";
        public const string VbaRead = "vba_read";
        public const string VbaUpdate = "vba_update";
        public const string ReadFormatBatch = "read_format_batch";
        public const string CreateChart = "create_chart";

        public const string CopyPaste = "copy_paste";
        public const string DeleteCells = "delete_cells";
    }
}
