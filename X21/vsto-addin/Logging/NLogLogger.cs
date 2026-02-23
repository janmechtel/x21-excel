using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.IO;
using System.Reflection;
using X21.Utils;
using X21.Services;

namespace X21.Logging
{
    public class NLogLogger : ILogger
    {
        public static void Init()
        {
            string configPath = null;
            bool configLoaded = false;

            try
            {
                // Try multiple locations to find NLog.config
                // 1. AppDomain.CurrentDomain.BaseDirectory (works for F5 debug)
                // 2. Assembly location (works for ClickOnce deployments)
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string assemblyLocation = Assembly.GetExecutingAssembly().Location;
                string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);

                // Try BaseDirectory first (better for F5 debugging)
                configPath = Path.Combine(baseDirectory, "NLog.config");
                if (!File.Exists(configPath))
                {
                    // Fallback to assembly directory
                    configPath = Path.Combine(assemblyDirectory, "NLog.config");
                }

                // Check if NLog.config exists
                if (File.Exists(configPath))
                {
                    // Explicitly load the configuration from the file
                    LogManager.Configuration = new XmlLoggingConfiguration(configPath);
                    configLoaded = true;

                    // Override the log path to be environment-specific
                    var config = LogManager.Configuration;
                    if (config != null)
                    {
                        var assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
                        var envName = EnvironmentHelper.GetEnvironmentName();
                        var envLogPath = EnvironmentHelper.GetExcelLogPath();
                        var machineName = Environment.MachineName;
                        var logFileName = $"xls-{machineName}";

                        // Debug: Write to desktop to verify what's happening
                        try
                        {
                            var debugPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "NLog-Debug.txt");
                            File.WriteAllText(debugPath,
                                $"Base Directory: {baseDirectory}\n" +
                                $"Assembly Directory: {assemblyDirectory}\n" +
                                $"Config Path: {configPath}\n" +
                                $"Config Exists: {File.Exists(configPath)}\n" +
                                $"Assembly Name: {assemblyName}\n" +
                                $"Environment Name: {envName}\n" +
                                $"Excel Log Path: {envLogPath}\n" +
                                $"Log File Name: {logFileName}\n" +
                                $"Target Count: {config.AllTargets.Count}\n");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Trace.WriteLine($"NLog debug write failed: {ex}");
                        }

                        // Ensure the log directory exists
                        Directory.CreateDirectory(envLogPath);

                        // Update all file targets to use the environment-specific path
                        foreach (var target in config.AllTargets)
                        {
                            if (target is FileTarget fileTarget)
                            {
                                fileTarget.FileName = Path.Combine(envLogPath, $"{logFileName}.log");
                                fileTarget.ArchiveFileName = Path.Combine(envLogPath, $"{logFileName}_{{##}}.log");
                            }
                        }

                        // Reconfigure to apply the new paths
                        LogManager.ReconfigExistingLoggers();
                    }
                }
                else
                {
                    // Fallback: Try to use the default configuration search
                    // NLog will look for NLog.config in the application base directory
                    LogManager.ReconfigExistingLoggers();
                }
            }
            catch (Exception ex)
            {
                // If configuration fails, write to desktop as a fallback
                Logger.LogExceptionToDesktop(ex, "NLog-Init-Error.txt");
            }

            var logger = new NLogLogger();
            Logger.Instance = logger;

            // Log initialization info to help troubleshoot
            try
            {
                var nlogLogger = LogManager.GetCurrentClassLogger();
                nlogLogger.Info("=== NLog Initialized ===");
                nlogLogger.Info("Config Path: {0}", configPath ?? "null");
                nlogLogger.Info("Config Loaded: {0}", configLoaded);
                nlogLogger.Info("Config File Exists: {0}", configPath != null && File.Exists(configPath));
                nlogLogger.Info("Assembly Location: {0}", Assembly.GetExecutingAssembly().Location);
                nlogLogger.Info("Log targets configured: {0}", LogManager.Configuration?.AllTargets?.Count ?? 0);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"NLog diagnostic logging failed: {ex}");
            }
        }
        public void LogException(Exception ex)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Error, ex);
            PostHogService.Instance.CaptureError(ex);
        }

        public void LogExceptionUI(Exception ex)
        {
            throw new NotImplementedException();
        }

        public void Debug(object obj)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Debug, obj);
        }

        public void Debug(string message, params object[] args)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Debug, message, args);
        }

        public void Info(string message, params object[] args)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Info, message, args);
        }

        public void Info(object obj)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Info, obj);
        }

        public void Warn(string message, params object[] args)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Warn, message, args);
            PostHogService.Instance.CaptureWarning(FormatMessage(message, args));
        }

        public void Warn(object obj)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Warn, obj);
            PostHogService.Instance.CaptureWarning(obj?.ToString() ?? "Warning");
        }

        public void Error(string message, params object[] args)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Error, message, args);
            PostHogService.Instance.CaptureError(FormatMessage(message, args));
        }

        public void Error(object obj)
        {
            var logger = LogManager.GetLogger(GetCallingClass());
            logger.Log(LogLevel.Error, obj);
            PostHogService.Instance.CaptureError(obj?.ToString() ?? "Error");
        }

        private string FormatMessage(string message, object[] args)
        {
            try
            {
                return args == null || args.Length == 0
                    ? message
                    : string.Format(message, args);
            }
            catch (FormatException ex)
            {
                System.Diagnostics.Trace.WriteLine($"NLog message formatting failed: {ex}");
                return message;
            }
            catch (ArgumentNullException ex)
            {
                System.Diagnostics.Trace.WriteLine($"NLog message formatting failed: {ex}");
                return message;
            }
        }

        private string GetCallingClass()
        {
            var stackTrace = new System.Diagnostics.StackTrace();
            // Skip this method, the Logger static method, and get the actual calling class
            var frame = stackTrace.GetFrame(3);
            return frame?.GetMethod()?.DeclaringType?.Name ?? "Unknown";
        }
    }
}
