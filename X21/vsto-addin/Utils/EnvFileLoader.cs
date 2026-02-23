using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using X21.Logging;

namespace X21.Utils
{
    public static class EnvFileLoader
    {
        /// <summary>
        /// Whitelist of allowed environment variable names that can be loaded from .env file.
        /// This prevents arbitrary environment variable injection attacks.
        /// </summary>
        private static readonly HashSet<string> AllowedVariables = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "POSTHOG_API_KEY",
            "POSTHOG_API_HOST",
            "POSTHOG_LOGS_ENABLED",
            "POSTHOG_LOGS_ENDPOINT"
        };

        public static void Load()
        {
            try
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var assemblyLocation = Assembly.GetExecutingAssembly().Location;
                var assemblyDir = string.IsNullOrWhiteSpace(assemblyLocation)
                    ? null
                    : Path.GetDirectoryName(assemblyLocation);

                var candidatePaths = new List<string>();
                if (!string.IsNullOrWhiteSpace(baseDir))
                {
                    candidatePaths.Add(BuildEnvPath(baseDir));
                }
                if (!string.IsNullOrWhiteSpace(assemblyDir))
                {
                    candidatePaths.Add(BuildEnvPath(assemblyDir));
                }

                var checkedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var loadedAny = false;
                foreach (var path in candidatePaths)
                {
                    if (!checkedPaths.Add(path))
                    {
                        continue;
                    }

                    if (!File.Exists(path))
                    {
                        continue;
                    }

                    LoadFile(path);
                    Logger.Info($"Loaded environment file: {path}");
                    loadedAny = true;
                }

                if (!loadedAny && checkedPaths.Count > 0)
                {
                    Logger.Info($"No .env file found. Checked: {string.Join("; ", checkedPaths)}");
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Failed to load .env files: {ex.Message}");
            }
        }

        private static void LoadFile(string path)
        {
            foreach (var rawLine in File.ReadAllLines(path))
            {
                var line = rawLine.Trim();
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#") || line.StartsWith(";"))
                {
                    continue;
                }

                if (line.StartsWith("export ", StringComparison.Ordinal))
                {
                    line = line.Substring("export ".Length).TrimStart();
                }

                var separatorIndex = line.IndexOf('=');
                if (separatorIndex <= 0)
                {
                    continue;
                }

                var key = line.Substring(0, separatorIndex).Trim();
                var value = line.Substring(separatorIndex + 1).Trim();
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                // Security: Only load whitelisted environment variables
                if (!AllowedVariables.Contains(key))
                {
                    Logger.Info($"Skipping environment variable '{key}' - not in whitelist");
                    continue;
                }

                if ((value.StartsWith("\"") && value.EndsWith("\"")) ||
                    (value.StartsWith("'") && value.EndsWith("'")))
                {
                    value = value.Substring(1, value.Length - 2);
                }

                // Only set if not already set (existing environment variables take precedence)
                if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable(key)))
                {
                    Environment.SetEnvironmentVariable(key, value);
                    Logger.Info($"Loaded environment variable: {key}");
                }
            }
        }

        private static string BuildEnvPath(string directory)
        {
            var trimmed = directory.TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar);
            return $"{trimmed}{Path.DirectorySeparatorChar}.env";
        }
    }
}
