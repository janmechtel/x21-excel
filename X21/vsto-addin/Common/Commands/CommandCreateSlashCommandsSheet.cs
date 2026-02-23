using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using X21.Common.Data;
using X21.Common.Model;
using X21.Logging;
using X21.Services;

namespace X21.Common.Commands
{
    /// <summary>
    /// Creates an "X21_Commands" sheet and seeds it with slash command metadata.
    /// </summary>
    public class CommandCreateSlashCommandsSheet : CommandBase
    {
        private const string SheetName = "X21_Commands";
        private static readonly HttpClient HttpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(5)
        };

        public CommandCreateSlashCommandsSheet(Container container, Annotation annotation)
            : base(container, annotation, nameof(CommandCreateSlashCommandsSheet))
        {
        }

        protected override async void ExecuteCore(object value)
        {
            var excelApp = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
            var workbook = excelApp?.ActiveWorkbook;

            if (workbook == null)
            {
                Logger.Info("No active workbook found when creating slash command sheet.");
                MessageBox.Show("No active workbook found. Please open a workbook first.", "X21", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var sheet = FindSheet(workbook, SheetName);
            var isNewSheet = false;

            if (sheet == null)
            {
                sheet = CreateSheetAtEnd(workbook);
                isNewSheet = sheet != null;
            }
            else
            {
                sheet.Activate();
                MessageBox.Show($"Sheet \"{SheetName}\" exists. Please update data there.", "X21", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            sheet?.Activate();

            if (isNewSheet && sheet != null)
            {
                try
                {
                    await PopulateSheetAsync(sheet);
                }
                catch (HttpRequestException ex)
                {
                    Logger.LogExceptionUI(ex);
                    MessageBox.Show("Failed to create commands sheet. Please try again.", "X21", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (System.Threading.Tasks.TaskCanceledException ex)
                {
                    Logger.LogExceptionUI(ex);
                    MessageBox.Show("Failed to create commands sheet. Please try again.", "X21", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (JsonException ex)
                {
                    Logger.LogExceptionUI(ex);
                    MessageBox.Show("Failed to create commands sheet. Please try again.", "X21", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    Logger.LogExceptionUI(ex);
                    MessageBox.Show("Failed to create commands sheet. Please try again.", "X21", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private Worksheet FindSheet(Workbook workbook, string name)
        {
            foreach (Worksheet candidate in workbook.Worksheets)
            {
                if (string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return candidate;
                }
            }

            return null;
        }

        private Worksheet CreateSheetAtEnd(Workbook workbook)
        {
            try
            {
                var worksheets = workbook.Worksheets;
                var newSheet = worksheets.Add(After: worksheets[worksheets.Count]) as Worksheet;
                if (newSheet != null)
                {
                    newSheet.Name = SheetName;
                }

                return newSheet;
            }
            catch (Exception ex)
            {
                Logger.LogExceptionUI(ex);
                return null;
            }
        }

        private async Task PopulateSheetAsync(Worksheet sheet)
        {
            var commands = await FetchBaseCommandsFromBackendAsync();
            if (commands == null || commands.Count == 0)
            {
                Logger.Info("Falling back to sample slash commands for sheet creation.");
                commands = GetFallbackCommands();
                MessageBox.Show(
                    "Could not load slash commands from the backend. Added a couple of sample commands instead. Replace them with your own once the backend is available.",
                    "X21",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            var headers = new[]
            {
                "Title",
                "Description",
                "Prompt",
                "Name",
                "Requires Input"
            };

            var rowCount = commands.Count + 1; // header + data rows
            var columnCount = headers.Length;

            var values = Array.CreateInstance(typeof(object), new[] { rowCount, columnCount }, new[] { 1, 1 });

            for (int col = 0; col < headers.Length; col++)
            {
                values.SetValue(headers[col], 1, col + 1);
            }

            for (int i = 0; i < commands.Count; i++)
            {
                var rowIndex = i + 2; // skip header
                var command = commands[i];

                values.SetValue(command.Title, rowIndex, 1);
                values.SetValue(command.Description, rowIndex, 2);
                values.SetValue(command.Prompt, rowIndex, 3);
                values.SetValue(command.Name, rowIndex, 4);
                values.SetValue(command.RequiresInput, rowIndex, 5);
            }

            var dataRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rowCount, columnCount]];
            dataRange.Value2 = values;

            ApplyFormatting(sheet, dataRange);
        }

        private async Task<List<SlashCommandRow>> FetchBaseCommandsFromBackendAsync()
        {
            var backendUrl = BackendConfigService.Instance.BaseUrl.TrimEnd('/');
            try
            {
                var requestUrl = $"{backendUrl}/api/slash-commands?separated=true";

                var response = await HttpClient.GetAsync(requestUrl).ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    Logger.Info($"Backend returned {response.StatusCode} when loading slash commands.");
                    ShowBackendErrorMessage(backendUrl, $"Status code: {response.StatusCode}");
                    return null;
                }

                var content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var parsed = JsonSerializer.Deserialize<SlashCommandsApiResponse>(content, options);

                if (parsed?.BaseCommands == null || parsed.BaseCommands.Count == 0)
                {
                    Logger.Info("Backend returned no base slash commands.");
                    ShowBackendErrorMessage(backendUrl, "No commands returned.");
                    return null;
                }

                return parsed.BaseCommands
                    .Select(SlashCommandRow.FromApiCommand)
                    .Where(row => row != null && !string.IsNullOrWhiteSpace(row.Name))
                    .ToList();
            }
            catch (HttpRequestException ex)
            {
                Logger.LogException(ex);
                ShowBackendErrorMessage(backendUrl, ex.Message);
                return null;
            }
            catch (System.Threading.Tasks.TaskCanceledException ex)
            {
                Logger.LogException(ex);
                ShowBackendErrorMessage(backendUrl, ex.Message);
                return null;
            }
            catch (JsonException ex)
            {
                Logger.LogException(ex);
                ShowBackendErrorMessage(backendUrl, ex.Message);
                return null;
            }
        }

        private void ShowBackendErrorMessage(string backendUrl, string details)
        {
            var port = GetPortFromBaseUrl(backendUrl);
            var message =
                $"Could not load slash commands from backend at {backendUrl} (port {port}).\n" +
                $"Details: {details}\n" +
                "Ensure the Deno backend is running and reachable. Added sample commands as fallback.";

            MessageBox.Show(message, "X21", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private string GetPortFromBaseUrl(string backendUrl)
        {
            try
            {
                var uri = new Uri(backendUrl);
                return uri.IsDefaultPort ? "default" : uri.Port.ToString();
            }
            catch (UriFormatException)
            {
                return "unknown";
            }
        }

        private List<SlashCommandRow> GetFallbackCommands()
        {
            return new List<SlashCommandRow>
            {
                new SlashCommandRow
                {
                    Name = "tell_a_joke",
                    Title = "Tell a Joke",
                    Description = "Return a short, clean joke.",
                    Prompt = "Tell a short, clean joke about work or technology.",
                    RequiresInput = false
                },
                new SlashCommandRow
                {
                    Name = "summarize_selection",
                    Title = "Summarize Selection",
                    Description = "Summarize the current Excel selection and highlight notable findings.",
                    Prompt = "Summarize the current Excel selection, highlight the most important findings, and call out outliers or trends that require attention.",
                    RequiresInput = false
                }
            };
        }

        private void ApplyFormatting(Worksheet sheet, Range dataRange)
        {
            var headerRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 5]];
            headerRange.Font.Bold = true;
            headerRange.Font.Size = 12;
            headerRange.Font.Color = ColorTranslator.ToOle(Color.Black);
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));

            try
            {
                var table = sheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, dataRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing);
                table.Name = "X21SlashCommands";
                table.TableStyle = "TableStyleMedium2";
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
            }

            ((Range)sheet.Columns[1]).ColumnWidth = 24; // Title
            ((Range)sheet.Columns[2]).ColumnWidth = 50; // Description
            ((Range)sheet.Columns[3]).ColumnWidth = 70; // Prompt
            ((Range)sheet.Columns[4]).ColumnWidth = 20; // Name
            ((Range)sheet.Columns[5]).ColumnWidth = 14; // Requires Input

            ((Range)sheet.Columns[2]).WrapText = true;
            ((Range)sheet.Columns[3]).WrapText = true;
        }

        private class SlashCommandRow
        {
            public string Name { get; set; } = string.Empty;
            public string Title { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
            public bool RequiresInput { get; set; }
            public string Prompt { get; set; } = string.Empty;

            public static SlashCommandRow FromApiCommand(SlashCommandApiModel command)
            {
                if (command == null) return null;

                return new SlashCommandRow
                {
                    Name = command.Name ?? command.Id ?? string.Empty,
                    Title = command.Title ?? command.Name ?? command.Id ?? string.Empty,
                    Description = command.Description ?? string.Empty,
                    RequiresInput = command.RequiresInput ?? false,
                    Prompt = command.Prompt ?? string.Empty
                };
            }
        }

        private class SlashCommandApiModel
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Title { get; set; }
            public string Description { get; set; }
            public bool? RequiresInput { get; set; }
            public string Prompt { get; set; }
            public string Icon { get; set; }
            public string InputPlaceholder { get; set; }
            public string DefaultInput { get; set; }
            public string Category { get; set; }
            public List<string> Keywords { get; set; }
        }

        private class SlashCommandsApiResponse
        {
            public List<SlashCommandApiModel> BaseCommands { get; set; }
        }
    }
}
