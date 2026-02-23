using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace X21.Models
{
    public class ToolExecutionRequest
    {
        [JsonPropertyName("toolName")]
        public string ToolName { get; set; }

        [JsonPropertyName("parameters")]
        public Dictionary<string, object> Parameters { get; set; }

        [JsonPropertyName("context")]
        public ExcelContext Context { get; set; }
    }

    public class ExcelContext
    {
        [JsonPropertyName("activeSheet")]
        public string ActiveSheet { get; set; }

        [JsonPropertyName("selectedRange")]
        public string SelectedRange { get; set; }

        [JsonPropertyName("allSheets")]
        public List<string> AllSheets { get; set; }

        [JsonPropertyName("usedRange")]
        public string UsedRange { get; set; }
    }

    public class ExcelRangeResponse
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("values")]
        public object[][] Values { get; set; }
    }

    public class ReadRangeResponse
    {
        [JsonPropertyName("cellValues")]
        public Dictionary<string, CellValue> CellValues { get; set; }
    }

    public class ReadRangeRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }
    }

    public class ReadRangeBatchOperation
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }
    }

    public class ReadRangeBatchRequest
    {
        [JsonPropertyName("operations")]
        public ReadRangeBatchOperation[] Operations { get; set; }
    }

    public class ReadRangeBatchResult
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("cellValues")]
        public Dictionary<string, CellValue> CellValues { get; set; }
    }

    public class ReadRangeBatchResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("results")]
        public ReadRangeBatchResult[] Results { get; set; }
    }

    public class WriteRangeRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("values")]
        public object[][] Values { get; set; }

        [JsonPropertyName("columnWidthMode")]
        public string ColumnWidthMode { get; set; }
    }

    public class WriteRangeBatchOperation
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("values")]
        public object[][] Values { get; set; }
    }

    public class WriteRangeBatchRequest
    {
        [JsonPropertyName("operations")]
        public WriteRangeBatchOperation[] Operations { get; set; }

        [JsonPropertyName("columnWidthMode")]
        public string ColumnWidthMode { get; set; }
    }

    public class WriteRangeResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }
    }

    public class WriteRangeBatchResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("results")]
        public WriteRangeResponse[] Results { get; set; }
    }

    public class FormatSettings
    {
        [JsonPropertyName("bold")]
        public bool? Bold { get; set; }

        [JsonPropertyName("italic")]
        public bool? Italic { get; set; }

        [JsonPropertyName("underline")]
        public bool? Underline { get; set; }

        [JsonPropertyName("fontColor")]
        public string FontColor { get; set; }

        [JsonPropertyName("backgroundColor")]
        public string BackgroundColor { get; set; } // Hex color or "none"

        [JsonPropertyName("alignment")]
        public string Alignment { get; set; }

        [JsonPropertyName("numberFormat")]
        public string NumberFormat { get; set; }

        [JsonPropertyName("fontSize")]
        public int? FontSize { get; set; }

        [JsonPropertyName("fontName")]
        public string FontName { get; set; }

        [JsonPropertyName("clearBorders")]
        public bool? ClearBorders { get; set; }

        // [JsonPropertyName("border")] // COMMENTED OUT TO STRIP BORDERS FROM FORMATTING
        // public BorderSettings Border { get; set; }
    }

    // public class BorderSettings // COMMENTED OUT TO STRIP BORDERS FROM FORMATTING
    // {
    //     [JsonPropertyName("bottom")]
    //     public string Bottom { get; set; } // "thin", "medium", "thick", "none"
    //
    //     [JsonPropertyName("top")]
    //     public string Top { get; set; } // "thin", "medium", "thick", "none"
    //
    //     [JsonPropertyName("left")]
    //     public string Left { get; set; } // "thin", "medium", "thick", "none"
    //
    //     [JsonPropertyName("right")]
    //     public string Right { get; set; } // "thin", "medium", "thick", "none"
    // }

    public class FormatRangeResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }
    }

    public class ReadFormatRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("propertiesToRead")]
        public List<string> PropertiesToRead { get; set; }
    }

    public class ReadFormatResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("cellFormats")]
        public Dictionary<string, FormatSettings> CellFormats { get; set; }
    }

    public class CellValue
    {
        [JsonPropertyName("value")]
        public string Value { get; set; }

        [JsonPropertyName("formula")]
        public string Formula { get; set; }
    }
}
