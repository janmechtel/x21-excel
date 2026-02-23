using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Linq;
using X21.Common.Data;
using X21.Common.Model;
using Microsoft.Office.Interop.Excel;

namespace X21.Services
{
    /// <summary>
    /// Tracks workbook changes by capturing snapshots and sending them to the backend.
    /// Snapshot storage and comparison is handled server-side.
    ///
    /// Flow:
    /// - SendInitialSnapshot: Called when workbook opens, sends isInitialSnapshot=true
    /// - GenerateChangelog: Called for consecutive snapshots, sends isInitialSnapshot=false
    ///
    /// Server-side behavior:
    /// - Initial snapshots establish the baseline (only saved if none exists)
    /// - Consecutive snapshots generate diffs and summaries (only saved if changes detected)
    /// </summary>
    public class WorkbookChangeTrackerService : Component
    {
        private readonly HttpClient _httpClient = new HttpClient();
        private readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };

        public WorkbookChangeTrackerService(Container container) : base(container)
        {
        }

        public override void Init()
        {
            base.Init();
            Logger.Info("WorkbookChangeTrackerService initialized");
        }

        /// <summary>
        /// Captures the initial snapshot of the workbook and sends it to the backend.
        /// Called when a workbook is opened or created.
        /// </summary>
        public async Task SendInitialSnapshot(Workbook workbook)
        {
            if (workbook == null)
            {
                Logger.Info("Workbook is null, skipping snapshot");
                return;
            }

            var workbookName = workbook.Name;
            var snapshotExists = await SnapshotExistsAsync(workbookName);
            if (snapshotExists == true)
            {
                Logger.Info($"Snapshot already exists for '{workbookName}', skipping initial snapshot");
                return;
            }

            if (snapshotExists == null)
            {
                Logger.Info($"Could not verify snapshot existence for '{workbookName}', skipping initial snapshot");
                return;
            }

            await SendSnapshotInternal(workbook, isInitialSnapshot: true, "capturing snapshot", "Error capturing snapshot");
        }

        /// <summary>
        /// Sends the snapshot to the backend with summary generation.
        /// </summary>
        public async Task GenerateChangelog(Workbook workbook, string comparisonFilePath = null)
        {
            await SendSnapshotInternal(workbook, isInitialSnapshot: false, "Generating changelog", "Error generating changelog manually", comparisonFilePath);
        }

        /// <summary>
        /// Copies snapshot from source workbook to target workbook (e.g., on Save As).
        /// </summary>
        public async Task CopySnapshot(string sourceWorkbookName, string targetWorkbookName)
        {
            try
            {
                Logger.Info($"Copying snapshot from '{sourceWorkbookName}' to '{targetWorkbookName}'");

                var requestData = new
                {
                    sourceWorkbookName,
                    targetWorkbookName
                };

                var json = JsonSerializer.Serialize(requestData, _jsonOptions);
                using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
                {
                    var backendUrl = BackendConfigService.Instance.BaseUrl;
                    var response = await _httpClient.PostAsync($"{backendUrl}/api/workbook-snapshot/copy", content);

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = await response.Content.ReadAsStringAsync();
                        Logger.Info($"Successfully copied snapshot. Response: {responseContent}");
                    }
                    else
                    {
                        Logger.Info($"Failed to copy snapshot. Status: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error copying snapshot: {ex.Message}");
            }
        }

        /// <summary>
        /// Checks if the user has permission to save copies (snapshots).
        /// </summary>
        private async Task<bool> HasSaveCopiesPermissionAsync()
        {
            try
            {
                var backendUrl = BackendConfigService.Instance.BaseUrl;
                var response = await _httpClient.GetAsync($"{backendUrl}/api/user-preference?key=save_snapshots");

                if (response.IsSuccessStatusCode)
                {
                    var responseContent = await response.Content.ReadAsStringAsync();
                    var result = JsonSerializer.Deserialize<JsonElement>(responseContent);

                    if (result.TryGetProperty("preferenceValue", out var valueElement))
                    {
                        var hasPermission = valueElement.GetBoolean();
                        Logger.Info($"Save copies permission: {hasPermission}");
                        return hasPermission;
                    }
                }

                Logger.Info("Could not determine save copies permission, defaulting to false");
                return false;
            }
            catch (HttpRequestException ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error checking save copies permission: {ex.Message}, defaulting to false");
                return false;
            }
            catch (TaskCanceledException ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error checking save copies permission (operation canceled): {ex.Message}, defaulting to false");
                return false;
            }
            catch (JsonException ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error checking save copies permission (invalid JSON): {ex.Message}, defaulting to false");
                return false;
            }
        }

        /// <summary>
        /// Checks if a snapshot already exists for the workbook.
        /// </summary>
        private async Task<bool?> SnapshotExistsAsync(string workbookName)
        {
            try
            {
                var backendUrl = BackendConfigService.Instance.BaseUrl;
                var encodedName = Uri.EscapeDataString(workbookName);
                var response = await _httpClient.GetAsync($"{backendUrl}/api/workbook-snapshot/exists?workbookName={encodedName}");

                if (response.IsSuccessStatusCode)
                {
                    var responseContent = await response.Content.ReadAsStringAsync();
                    var result = JsonSerializer.Deserialize<JsonElement>(responseContent);

                    if (result.TryGetProperty("exists", out var existsElement))
                    {
                        var exists = existsElement.GetBoolean();
                        Logger.Info($"Snapshot exists for '{workbookName}': {exists}");
                        return exists;
                    }
                }

                Logger.Info($"Could not determine snapshot existence for '{workbookName}'");
                return null;
            }
            catch (HttpRequestException ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error checking snapshot existence: {ex.Message}");
                return null;
            }
            catch (TaskCanceledException ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error checking snapshot existence (operation canceled): {ex.Message}");
                return null;
            }
            catch (JsonException ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error checking snapshot existence (invalid JSON): {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Internal method that captures a snapshot and sends it to the backend.
        /// </summary>
        private async Task SendSnapshotInternal(Workbook workbook, bool isInitialSnapshot, string actionDescription, string errorMessage, string comparisonFilePath = null)
        {
            try
            {
                if (workbook == null)
                {
                    Logger.Info("Workbook is null, skipping snapshot");
                    return;
                }

                // When generating changes from other files (comparisonFilePath provided),
                // we don't need to check permissions because we're not saving snapshots,
                // just comparing files and generating a summary
                bool isGeneratingFromOtherFile = !string.IsNullOrWhiteSpace(comparisonFilePath);

                if (!isGeneratingFromOtherFile)
                {
                    // Check permission before creating snapshot (only when saving snapshots)
                    bool hasPermission = await HasSaveCopiesPermissionAsync();
                    if (!hasPermission)
                    {
                        Logger.Info("User has not consented to save copies, skipping snapshot");
                        return;
                    }
                }
                else
                {
                    Logger.Info("Generating changes from other file - permission check skipped");
                }

                var workbookName = workbook.Name;
                Logger.Info($"{actionDescription}: {workbookName}");

                // Capture current snapshot
                WorkbookSnapshot snapshot = ExtractWorkbookXmlContentFromOpenWorkbook(workbook);
                Logger.Info($"Snapshot captured with {snapshot.SheetXmls.Count} sheets");

                // If comparison file path is provided, extract snapshot from that file too
                WorkbookSnapshot comparisonSnapshot = null;
                if (!string.IsNullOrWhiteSpace(comparisonFilePath) && File.Exists(comparisonFilePath))
                {
                    Logger.Info($"Extracting comparison snapshot from file: {comparisonFilePath}");
                    comparisonSnapshot = ExtractWorkbookXmlContentFromFile(comparisonFilePath);
                    Logger.Info($"Comparison snapshot captured with {comparisonSnapshot.SheetXmls.Count} sheets");
                }

                // Send to backend
                await SendSnapshotToEndpoint(workbookName, snapshot, isInitialSnapshot, comparisonSnapshot);
            }
            catch (TaskCanceledException ex)
            {
                Logger.LogException(ex);
                Logger.Info($"{errorMessage} (operation canceled): {ex.Message}");
            }
            catch (OperationCanceledException ex)
            {
                Logger.LogException(ex);
                Logger.Info($"{errorMessage} (operation canceled): {ex.Message}");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"{errorMessage}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Sends the snapshot to the backend endpoint.
        ///
        /// Request payload:
        /// - workbookName: Name of the workbook
        /// - currentSnapshot: The snapshot data (sharedStringsXml, sheetXmls)
        /// - isInitialSnapshot: true for initial snapshots, false for consecutive snapshots
        /// - comparisonSnapshot: Optional snapshot from comparison file (instead of previous snapshot from DB)
        /// </summary>
        private async Task SendSnapshotToEndpoint(string workbookName, WorkbookSnapshot snapshot, bool isInitialSnapshot, WorkbookSnapshot comparisonSnapshot = null)
        {
            try
            {
                Logger.Info($"Sending snapshot to backend: {workbookName}, isInitialSnapshot: {isInitialSnapshot}");

                var requestData = new
                {
                    workbookName,
                    currentSnapshot = new
                    {
                        workbookXml = snapshot.WorkbookXml,
                        workbookRelsXml = snapshot.WorkbookRelsXml,
                        sharedStringsXml = snapshot.SharedStringsXml,
                        sheetXmls = snapshot.SheetXmls
                    },
                    isInitialSnapshot,
                    comparisonSnapshot = comparisonSnapshot != null ? new
                    {
                        workbookXml = comparisonSnapshot.WorkbookXml,
                        workbookRelsXml = comparisonSnapshot.WorkbookRelsXml,
                        sharedStringsXml = comparisonSnapshot.SharedStringsXml,
                        sheetXmls = comparisonSnapshot.SheetXmls,
                        filePath = comparisonSnapshot.FilePath,
                        lastModified = comparisonSnapshot.LastModified.HasValue
                            ? (long?)(comparisonSnapshot.LastModified.Value.Subtract(new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds)
                            : null
                    } : null
                };

                var json = JsonSerializer.Serialize(requestData, _jsonOptions);
                using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
                {
                    var backendUrl = BackendConfigService.Instance.BaseUrl;
                    var response = await _httpClient.PostAsync($"{backendUrl}/api/workbook-snapshot", content);

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = await response.Content.ReadAsStringAsync();
                        Logger.Info($"Successfully sent snapshot to backend. Response: {responseContent}");
                    }
                    else
                    {
                        Logger.Info($"Failed to send snapshot to backend. Status: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error sending snapshot to backend: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts XML content from a closed Excel file by path
        /// </summary>
        private WorkbookSnapshot ExtractWorkbookXmlContentFromFile(string filePath)
        {
            try
            {
                Logger.Info($"Extracting XML from file: {filePath}");

                // Get file info for metadata
                var fileInfo = new FileInfo(filePath);
                var fullPath = fileInfo.FullName;
                var lastModified = fileInfo.LastWriteTimeUtc;

                var snapshot = new WorkbookSnapshot
                {
                    WorkbookXml = null,
                    WorkbookRelsXml = null,
                    SheetXmls = new Dictionary<string, string>(),
                    SharedStringsXml = null,
                    FilePath = fullPath,
                    LastModified = lastModified
                };

                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Read))
                {
                    // Extract workbook.xml (contains sheet name mappings)
                    var workbookEntry = archive.GetEntry("xl/workbook.xml");
                    if (workbookEntry != null)
                    {
                        using (var stream = workbookEntry.Open())
                        using (var reader = new StreamReader(stream))
                        {
                            snapshot.WorkbookXml = reader.ReadToEnd();
                        }
                    }

                    // Extract workbook.xml.rels (maps rId to actual sheet paths)
                    var workbookRelsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
                    if (workbookRelsEntry != null)
                    {
                        using (var stream = workbookRelsEntry.Open())
                        using (var reader = new StreamReader(stream))
                        {
                            snapshot.WorkbookRelsXml = reader.ReadToEnd();
                        }
                    }

                    // Extract shared strings
                    var sharedStringsEntry = archive.GetEntry("xl/sharedStrings.xml");
                    if (sharedStringsEntry != null)
                    {
                        using (var stream = sharedStringsEntry.Open())
                        using (var reader = new StreamReader(stream))
                        {
                            snapshot.SharedStringsXml = reader.ReadToEnd();
                        }
                    }

                    // Extract all worksheet XMLs
                    foreach (var entry in archive.Entries.Where(e =>
                        e.FullName.StartsWith("xl/worksheets/sheet") && e.FullName.EndsWith(".xml")))
                    {
                        using (var stream = entry.Open())
                        using (var reader = new StreamReader(stream))
                        {
                            snapshot.SheetXmls[entry.FullName] = reader.ReadToEnd();
                        }
                    }
                }

                Logger.Info($"Extracted {snapshot.SheetXmls.Count} sheet XMLs from file");

                return snapshot;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error extracting XML from file {filePath}: {ex.Message}");

                return new WorkbookSnapshot
                {
                    WorkbookXml = null,
                    WorkbookRelsXml = null,
                    SheetXmls = new Dictionary<string, string>(),
                    SharedStringsXml = null,
                    FilePath = null,
                    LastModified = null
                };
            }
        }

        /// <summary>
        /// Extracts XML content from an open Excel workbook by saving to a temporary file
        /// </summary>
        private WorkbookSnapshot ExtractWorkbookXmlContentFromOpenWorkbook(Workbook workbook)
        {
            string tempFilePath = null;

            try
            {
                // Create a temporary file path
                tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

                Logger.Info($"Saving workbook to temporary file: {tempFilePath}");

                // Save a copy of the workbook to the temp file
                workbook.SaveCopyAs(tempFilePath);

                // Extract XML from the temp file using the shared method
                return ExtractWorkbookXmlContentFromFile(tempFilePath);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error extracting XML from open workbook: {ex.Message}");

                return new WorkbookSnapshot
                {
                    WorkbookXml = null,
                    WorkbookRelsXml = null,
                    SheetXmls = new Dictionary<string, string>(),
                    SharedStringsXml = null,
                    FilePath = null,
                    LastModified = null
                };
            }
            finally
            {
                // Clean up the temporary file
                if (tempFilePath != null && File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                        Logger.Info($"Deleted temporary file: {tempFilePath}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Info($"Failed to delete temporary file {tempFilePath}: {ex.Message}");
                    }
                }
            }
        }
    }

    /// <summary>
    /// Represents a snapshot of a workbook's XML content
    /// </summary>
    internal class WorkbookSnapshot
    {
        public string WorkbookXml { get; set; }
        public string WorkbookRelsXml { get; set; }
        public string SharedStringsXml { get; set; }
        public Dictionary<string, string> SheetXmls { get; set; }
        public string FilePath { get; set; }
        public DateTime? LastModified { get; set; }
    }
}
