using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using X21.Logging;
using X21.Models;
using X21.Services.Formatting;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Main orchestrator for reading Excel formatting using specialized format readers.
    /// COM calls stay on the STA/main thread. Optional snapshot-capable readers can compute in parallel without COM.
    /// </summary>
    public class FormatManager
    {
        private readonly IFormatReader[] _formatReaders;

        public FormatManager()
        {
            // Initialize all format readers in the order they should be executed
            _formatReaders = new IFormatReader[]
            {
                new BoldFormatReader(),
                new ItalicFormatReader(),
                new UnderlineFormatReader(),
                new FontSizeFormatReader(),
                new NumberFormatReader(),
                new ColorFormatReader(),
                new AlignmentFormatReader()
                // new BorderFormatReader() // Commented out to strip borders from formatting
            };
        }

        /// <summary>
        /// Synchronous entrypoint for callers that cannot await.
        /// </summary>
        public Dictionary<string, FormatSettings> GetFormattedCells(Range targetRange, Action<FormatProgressUpdate> progress = null)
        {
            return GetFormattedCells(targetRange, null, progress);
        }

        /// <summary>
        /// Synchronous entrypoint with property filter.
        /// </summary>
        public Dictionary<string, FormatSettings> GetFormattedCells(Range targetRange, List<string> propertiesToRead, Action<FormatProgressUpdate> progress = null)
        {
            var scheduler = SynchronizationContext.Current != null
                ? TaskScheduler.FromCurrentSynchronizationContext()
                : TaskScheduler.Current;

            return GetFormattedCells(targetRange, propertiesToRead, CancellationToken.None, scheduler, progress);
        }

        public Task<Dictionary<string, FormatSettings>> GetFormattedCellAsync(Range targetRange, List<string> propertiesToRead, Action<FormatProgressUpdate> progress = null)
        {
            var scheduler = SynchronizationContext.Current != null
                ? TaskScheduler.FromCurrentSynchronizationContext()
                : TaskScheduler.Current;

            return GetFormattedCellsAsync(targetRange, propertiesToRead, CancellationToken.None, scheduler, progress);
        }

        /// <summary>
        /// Synchronous entrypoint that allows explicit STA scheduler and cancellation.
        /// </summary>
        public Dictionary<string, FormatSettings> GetFormattedCells(
            Range targetRange,
            List<string> propertiesToRead,
            CancellationToken cancellationToken,
            TaskScheduler uiScheduler,
            Action<FormatProgressUpdate> progress = null)
        {
            return GetFormattedCellsAsync(targetRange, propertiesToRead, cancellationToken, uiScheduler)
                .ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
        }

        /// <summary>
        /// Async-friendly orchestration. All COM access runs on the provided STA scheduler (Excel main thread).
        /// Snapshot-capable readers: COM snapshot on STA, parallel CPU compute on snapshots, then merge.
        /// Legacy readers: run sequentially on STA.
        /// </summary>
        public Task<Dictionary<string, FormatSettings>> GetFormattedCellsAsync(
            Range targetRange,
            List<string> propertiesToRead,
            CancellationToken cancellationToken,
            TaskScheduler uiScheduler = null,
            Action<FormatProgressUpdate> progress = null)
        {
            if (targetRange == null) throw new ArgumentNullException(nameof(targetRange));

            // If we are already on the Excel STA thread but there's no SynchronizationContext
            // (e.g., running inside ExcelStaDispatcher), run inline to keep COM on this thread.
            if (uiScheduler == null && SynchronizationContext.Current == null)
            {
                Logger.Info("No SynchronizationContext detected; running format read inline on current STA thread");
                var inlineResult = ComputeFormats(targetRange, propertiesToRead, cancellationToken, progress);
                return Task.FromResult(inlineResult);
            }

            if (uiScheduler == null)
            {
                uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            }

            return Task.Factory.StartNew(() =>
            {
                return ComputeFormats(targetRange, propertiesToRead, cancellationToken, progress);
            }, cancellationToken, TaskCreationOptions.DenyChildAttach, uiScheduler);
        }

        /// <summary>
        /// Gets the number of format readers being used
        /// </summary>
        public int ReaderCount => _formatReaders.Length;

        /// <summary>
        /// Gets the names of all format readers for debugging purposes
        /// </summary>
        public string[] GetReaderNames()
        {
            var names = new string[_formatReaders.Length];
            for (int i = 0; i < _formatReaders.Length; i++)
            {
                names[i] = _formatReaders[i].GetType().Name;
            }
            return names;
        }

        /// <summary>
        /// Filters format readers based on which properties need to be read
        /// </summary>
        private IFormatReader[] FilterReadersByProperties(List<string> propertiesToRead)
        {
            var readers = new List<IFormatReader>();
            var propsLower = new HashSet<string>(propertiesToRead.ConvertAll(p => p.ToLowerInvariant()));

            foreach (var reader in _formatReaders)
            {
                if (reader.SupportedProperties.Any(p => propsLower.Contains(p.ToLowerInvariant())))
                    readers.Add(reader);
            }

            Logger.Info($"Filtered {readers.Count}/{_formatReaders.Length} readers for properties: {string.Join(", ", propertiesToRead)}");
            return readers.ToArray();
        }

        private Dictionary<string, FormatSettings> ComputeFormats(
            Range targetRange,
            List<string> propertiesToRead,
            CancellationToken cancellationToken,
            Action<FormatProgressUpdate> progress = null)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var formattedCells = new Dictionary<string, FormatSettings>(StringComparer.OrdinalIgnoreCase);
            var app = targetRange.Application;
            var builder = new FormatSnapshotBuilder();
            var rows = Convert.ToInt32(targetRange.Rows.Count);
            var cols = Convert.ToInt32(targetRange.Columns.Count);
            var totalCells = Math.Max(1, rows * cols);

            try
            {
                var readersToUse = propertiesToRead == null
                    ? _formatReaders
                    : FilterReadersByProperties(propertiesToRead);

                Logger.Info($"Reading formats (unified snapshot + CPU compute) with {readersToUse.Length} readers" +
                            (propertiesToRead != null
                                ? $" (selective: {string.Join(", ", propertiesToRead)})"
                                : ""));

                // Stage plan: 1 = snapshot, 2..(readers+1) = readers, final = completion
                var totalStages = readersToUse.Length + 2;
                var stageIndex = 1;

                ReportProgress(progress, stageIndex, totalStages, totalCells, "Capturing formatting snapshot...");
                var snapshot = builder.Build(targetRange, propertiesToRead);
                stageIndex++;
                ReportProgress(progress, stageIndex, totalStages, totalCells, "Snapshot captured, processing formatting...");
                var useParallel = snapshot.CellCount >= 2000 && readersToUse.Length > 1;

                var swCompute = Stopwatch.StartNew();
                if (useParallel)
                {
                    var parallelResults = readersToUse
                        .AsParallel()
                        .WithCancellation(cancellationToken)
                        .Select(reader => ComputeReader(reader, snapshot, cancellationToken))
                        .ToList();

                    foreach (var result in parallelResults)
                    {
                        MergeReaderResults(formattedCells, result);
                        stageIndex++;
                        ReportProgress(
                            progress,
                            stageIndex,
                            totalStages,
                            totalCells,
                            $"Processed formatting reader {stageIndex - 1}/{readersToUse.Length}");
                    }
                }
                else
                {
                    foreach (var reader in readersToUse)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var readerResult = ComputeReader(reader, snapshot, cancellationToken);
                        MergeReaderResults(formattedCells, readerResult);
                        stageIndex++;
                        var friendlyName = reader.GetType().Name.Replace("FormatReader", string.Empty);
                        ReportProgress(
                            progress,
                            stageIndex,
                            totalStages,
                            totalCells,
                            $"Processed {friendlyName}");
                    }
                }
                ReportProgress(progress, totalStages, totalStages, totalCells, "Formatting read complete");
                Logger.Info($"CPU compute for formats finished in {swCompute.ElapsedMilliseconds} ms " +
                            $"({(useParallel ? "parallel" : "sequential")}, {formattedCells.Count} cells)");

                // Return clone to avoid exposing internal mutable instances.
                return formattedCells.ToDictionary(
                    kvp => kvp.Key,
                    kvp => CloneFormatSettings(kvp.Value),
                    StringComparer.OrdinalIgnoreCase);
            }
            finally
            {
                // Always clear search format when done; Excel can throw if already cleared.
                try { app.FindFormat.Clear(); } catch { }
            }
        }

        private static void MergeReaderResults(
            Dictionary<string, FormatSettings> aggregate,
            Dictionary<string, FormatSettings> readerResults)
        {
            foreach (var kvp in readerResults)
            {
                if (!aggregate.TryGetValue(kvp.Key, out var existing))
                {
                    aggregate[kvp.Key] = CloneFormatSettings(kvp.Value);
                    continue;
                }

                aggregate[kvp.Key] = MergeFormatSettings(existing, kvp.Value);
            }
        }

        // Treat FormatSettings as immutable during merge to avoid sharing mutable instances across readers.
        private static FormatSettings CloneFormatSettings(FormatSettings source)
        {
            if (source == null) return new FormatSettings();

            return new FormatSettings
            {
                Bold = source.Bold,
                Italic = source.Italic,
                Underline = source.Underline,
                FontColor = source.FontColor,
                BackgroundColor = source.BackgroundColor,
                Alignment = source.Alignment,
                NumberFormat = source.NumberFormat,
                FontSize = source.FontSize,
                FontName = source.FontName
            };
        }

        private static FormatSettings MergeFormatSettings(FormatSettings existing, FormatSettings incoming)
        {
            // Clone-on-write to prevent mutating shared state.
            var merged = CloneFormatSettings(existing);

            merged.Bold = incoming.Bold ?? merged.Bold;
            merged.Italic = incoming.Italic ?? merged.Italic;
            merged.Underline = incoming.Underline ?? merged.Underline;
            merged.FontSize = incoming.FontSize ?? merged.FontSize;
            merged.FontName = string.IsNullOrEmpty(incoming.FontName) ? merged.FontName : incoming.FontName;
            merged.FontColor = string.IsNullOrEmpty(incoming.FontColor) ? merged.FontColor : incoming.FontColor;
            merged.BackgroundColor = string.IsNullOrEmpty(incoming.BackgroundColor) ? merged.BackgroundColor : incoming.BackgroundColor;
            merged.NumberFormat = string.IsNullOrEmpty(incoming.NumberFormat) ? merged.NumberFormat : incoming.NumberFormat;
            merged.Alignment = string.IsNullOrEmpty(incoming.Alignment) ? merged.Alignment : incoming.Alignment;

            return merged;
        }

        private static void ReportProgress(Action<FormatProgressUpdate> progress, int stage, int totalStages, int totalCells, string message)
        {
            if (progress == null) return;

            var normalizedStages = Math.Max(totalStages, 1);
            var normalizedTotalCells = Math.Max(totalCells, normalizedStages);
            var ratio = Math.Max(0d, Math.Min(1d, (double)stage / normalizedStages));
            var current = (int)Math.Min(normalizedTotalCells, Math.Round(normalizedTotalCells * ratio));

            progress(new FormatProgressUpdate(current, normalizedTotalCells, message));
        }

        private static Dictionary<string, FormatSettings> ComputeReader(IFormatReader reader, FormatSnapshot snapshot, CancellationToken cancellationToken)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                var results = new Dictionary<string, FormatSettings>(StringComparer.OrdinalIgnoreCase);
                reader.ComputeFormats(snapshot, results);
                return results;
            }
            catch (Exception ex)
            {
                Logger.Info($"Error in {reader.Name}.ComputeFormats: {ex.Message}");
                return new Dictionary<string, FormatSettings>(StringComparer.OrdinalIgnoreCase);
            }
        }
    }
}
