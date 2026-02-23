using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using System.Text.Json;
using X21.Utils;

namespace X21.Models
{
    /// <summary>
    /// Represents a spreadsheet cell with its data and metadata
    /// </summary>
    public class Cell
    {
        private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            WriteIndented = false
        };

        [JsonPropertyName("address")]
        public string Address { get; set; }

        [JsonPropertyName("value")]
        public object Value { get; set; }

        [JsonPropertyName("formula")]
        public string Formula { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("formatting")]
        public string Formatting { get; set; }

        public Cell() { }

        public Cell(string address, object value, string formula = null, string type = null, string formatting = null)
        {
            Address = address;
            Value = value;
            Formula = formula;
            Type = type;
            Formatting = formatting;
        }

        public override string ToString()
        {
            return $"[address: {Address}, value: {Value}, formula: {Formula}]";
        }

        /// <summary>
        /// Converts this cell to array format [address, value, formula]
        /// </summary>
        public string[] ToArray()
        {
            return new string[] { Address ?? "", InvariantValueFormatter.ToInvariantString(Value), Formula ?? "" };
        }

        public static string CellListToString(List<Cell> cells)
        {
            if (cells == null)
            {
                return "[]";
            }
            return $"[{string.Join(", ", cells.Select(c => c.ToString()))}]";
        }

        /// <summary>
        /// Converts a list of cells to the new array format [[address, value, formula], ...]
        /// </summary>
        public static string CellListToArrayFormat(List<Cell> cells)
        {
            if (cells == null || cells.Count == 0)
            {
                return "[]";
            }

            var arrays = cells.Select(c =>
                $"[\"{c.Address ?? ""}\", \"{InvariantValueFormatter.ToInvariantString(c.Value) ?? ""}\", \"{c.Formula ?? ""}\"]");
            return $"[{string.Join(", ", arrays)}]";
        }
    }
}
