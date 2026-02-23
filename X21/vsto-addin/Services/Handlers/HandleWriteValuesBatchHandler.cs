using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using X21.Constants;
using X21.Common.Data;
using X21.Models;

namespace X21.Services.Handlers
{
    public class HandleWriteValuesBatchHandler : BaseExcelApiService
    {
        private const int MaxOperationsPerBatch = 10;
        private static readonly Regex CellRegex =
            new Regex(@"^([A-Z]+)(\d+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private readonly HttpClient _progressHttpClient = new HttpClient();

        public HandleWriteValuesBatchHandler(Container container) : base(container)
        {
        }

        public async Task HandleWriteValuesBatch(string requestBody, HttpListenerResponse response)
        {
            var writeBatchRequest = JsonSerializer.Deserialize<WriteRangeBatchRequest>(requestBody, _jsonOptions);

            if (writeBatchRequest == null)
            {
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, "Invalid request body");
                return;
            }

            if (writeBatchRequest.Operations == null || writeBatchRequest.Operations.Length == 0)
            {
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, "operations array is required and cannot be empty");
                return;
            }

            if (writeBatchRequest.Operations.Length > MaxOperationsPerBatch)
            {
                response.StatusCode = 400;
                await WriteJsonErrorAsync(
                    response,
                    $"operations array exceeds maximum of {MaxOperationsPerBatch}",
                    400);
                return;
            }

            var results = new List<WriteRangeResponse>();
            var totalOps = writeBatchRequest.Operations.Length;
            var progressWorkbook = writeBatchRequest.Operations
                .Select(op => op?.WorkbookName)
                .FirstOrDefault(name => !string.IsNullOrWhiteSpace(name));

            var columnWidthMode = writeBatchRequest.ColumnWidthMode;

            for (int i = 0; i < totalOps; i++)
            {
                var op = writeBatchRequest.Operations[i];

                if (string.IsNullOrWhiteSpace(progressWorkbook) && !string.IsNullOrWhiteSpace(op?.WorkbookName))
                {
                    progressWorkbook = op.WorkbookName;
                }

                if (op == null || string.IsNullOrEmpty(op.Worksheet) || string.IsNullOrEmpty(op.Range))
                {
                    results.Add(new WriteRangeResponse
                    {
                        Success = false,
                        Message = "worksheet and range are required for each operation",
                    });
                    await SendProgressUpdate(progressWorkbook, op, i + 1, totalOps);
                    continue;
                }

                if (!TryParseRange(op.Range, out var startRow, out var endRow, out var startCol, out var endCol))
                {
                    results.Add(new WriteRangeResponse
                    {
                        Success = false,
                        Message = "Invalid range format. Expected A1 or A1:B2",
                    });
                    await SendProgressUpdate(progressWorkbook, op, i + 1, totalOps);
                    continue;
                }

                var expectedRows = endRow - startRow + 1;
                var expectedCols = endCol - startCol + 1;
                if (!TryValidateValues(op.Values, expectedRows, expectedCols, out var valuesError))
                {
                    results.Add(new WriteRangeResponse
                    {
                        Success = false,
                        Message = valuesError,
                    });
                    await SendProgressUpdate(progressWorkbook, op, i + 1, totalOps);
                    continue;
                }

                try
                {
                    var result = await WriteExcelRangeAsync(
                        op.Worksheet,
                        op.WorkbookName,
                        op.Range,
                        op.Values,
                        columnWidthMode);
                    results.Add(result);
                }
                catch (WorkbookResolutionException wre)
                {
                    results.Add(new WriteRangeResponse
                    {
                        Success = false,
                        Message = wre.Message,
                    });
                }
                catch (Exception ex)
                {
                    Logger.LogException(ex);
                    results.Add(new WriteRangeResponse
                    {
                        Success = false,
                        Message = $"Failed to write range {op.Range} in sheet {op.Worksheet}: {ex.Message}",
                    });
                }
                await SendProgressUpdate(progressWorkbook, op, i + 1, totalOps);
            }

            var successCount = results.Count(r => r.Success);
            var batchResponse = new WriteRangeBatchResponse
            {
                Success = successCount == results.Count,
                Message = $"Batch write completed ({successCount}/{results.Count} successful)",
                Results = results.ToArray(),
            };

            await SendJsonResponse(response, batchResponse);
        }

        private async Task SendProgressUpdate(
            string workbookName,
            WriteRangeBatchOperation op,
            int current,
            int total)
        {
            try
            {
                var payload = new
                {
                    status = OperationStatusValues.WritingExcel,
                    message = $"Writing values ({current}/{total})",
                    progress = new
                    {
                        current = Math.Max(0, current),
                        total = Math.Max(1, total),
                        unit = "ops"
                    },
                    metadata = new
                    {
                        range = op?.Range,
                        worksheet = op?.Worksheet
                    }
                };

                var progressUrl = $"{BackendConfigService.Instance.BaseUrl.TrimEnd('/')}/api/progress";
                var json = JsonSerializer.Serialize(payload, _jsonOptions);

                using var requestMessage = new HttpRequestMessage(HttpMethod.Post, progressUrl);
                requestMessage.Content = new StringContent(json, Encoding.UTF8, "application/json");

                if (!string.IsNullOrEmpty(workbookName))
                {
                    requestMessage.Headers.Add("X-Workbook-Name", workbookName);
                }

                await _progressHttpClient.SendAsync(requestMessage);
            }
            catch (Exception ex)
            {
                Logger.Info($"Progress update failed: {ex.Message}");
            }
        }

        private static bool TryParseRange(
            string range,
            out int startRow,
            out int endRow,
            out int startCol,
            out int endCol)
        {
            startRow = endRow = startCol = endCol = 0;

            if (string.IsNullOrWhiteSpace(range)) return false;
            var trimmed = range.Trim();
            if (trimmed.Contains("!") || trimmed.Contains(",")) return false;

            var parts = trimmed.Split(':');
            if (parts.Length == 1)
            {
                if (!TryParseCellAddress(parts[0], out var row, out var col)) return false;
                startRow = endRow = row;
                startCol = endCol = col;
                return true;
            }

            if (parts.Length == 2)
            {
                if (!TryParseCellAddress(parts[0], out var startRowLocal, out var startColLocal)) return false;
                if (!TryParseCellAddress(parts[1], out var endRowLocal, out var endColLocal)) return false;
                if (endRowLocal < startRowLocal || endColLocal < startColLocal) return false;

                startRow = startRowLocal;
                endRow = endRowLocal;
                startCol = startColLocal;
                endCol = endColLocal;
                return true;
            }

            return false;
        }

        private static bool TryParseCellAddress(string address, out int row, out int col)
        {
            row = 0;
            col = 0;

            if (string.IsNullOrWhiteSpace(address)) return false;
            var match = CellRegex.Match(address.Trim());
            if (!match.Success) return false;

            if (!int.TryParse(match.Groups[2].Value, out row) || row <= 0) return false;
            col = ColumnToNumber(match.Groups[1].Value);
            return col > 0;
        }

        private static int ColumnToNumber(string column)
        {
            var result = 0;
            foreach (var ch in column.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                result = (result * 26) + (ch - 'A' + 1);
            }
            return result;
        }

        private static bool TryValidateValues(
            object[][] values,
            int expectedRows,
            int expectedCols,
            out string errorMessage)
        {
            errorMessage = null;

            if (values == null || values.Length == 0)
            {
                errorMessage = "values array is required and cannot be empty for each operation";
                return false;
            }

            if (values.Length != expectedRows)
            {
                errorMessage = $"values row count ({values.Length}) does not match range rows ({expectedRows})";
                return false;
            }

            for (int row = 0; row < values.Length; row++)
            {
                var rowValues = values[row];
                if (rowValues == null)
                {
                    errorMessage = $"values row {row + 1} is null";
                    return false;
                }

                if (rowValues.Length != expectedCols)
                {
                    errorMessage =
                        $"values column count ({rowValues.Length}) does not match range columns ({expectedCols}) at row {row + 1}";
                    return false;
                }
            }

            return true;
        }
    }
}
