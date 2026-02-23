using System;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;
using MsLogLevel = Microsoft.Extensions.Logging.LogLevel;
using OpenTelemetry.Exporter;
using OpenTelemetry.Logs;
using OpenTelemetry.Resources;
using X21.Logging;
using X21.Utils;

namespace X21.Services
{
    public class PostHogService : IDisposable
    {
        private static readonly object _lock = new object();
        private static PostHogService _instance;
        private Microsoft.Extensions.Logging.ILoggerFactory _loggerFactory;
        private Microsoft.Extensions.Logging.ILogger _otelLogger;
        private bool _disposed;

        // PostHog configuration (from environment variables)
        private static string GetPostHogApiKey()
        {
            return Environment.GetEnvironmentVariable("POSTHOG_API_KEY");
        }

        private static string GetPostHogLogsEndpoint()
        {
            return Environment.GetEnvironmentVariable("POSTHOG_LOGS_ENDPOINT") ??
                "https://us.i.posthog.com/i/v1/logs";
        }

        private static bool IsPostHogLogsEnabled()
        {
            var raw = Environment.GetEnvironmentVariable("POSTHOG_LOGS_ENABLED");
            return string.IsNullOrWhiteSpace(raw) ||
                !raw.Equals("false", StringComparison.OrdinalIgnoreCase);
        }

        private PostHogService()
        {
        }

        public static PostHogService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new PostHogService();
                        }
                    }
                }
                return _instance;
            }
        }

        public void Identify(string distinctId)
        {
            return;
        }

        public void CaptureWarning(string message)
        {
            Log(MsLogLevel.Warning, message, null);
        }

        public void CaptureError(string message)
        {
            Log(MsLogLevel.Error, message, null);
        }

        public void CaptureError(Exception ex)
        {
            if (ex == null) return;
            Log(MsLogLevel.Error, ex.Message, ex);
        }

        private bool EnsureLogger()
        {
            if (_disposed)
            {
                return false;
            }

            var apiKey = GetPostHogApiKey();
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                return false;
            }

            if (!IsPostHogLogsEnabled())
            {
                return false;
            }

            if (_otelLogger != null)
            {
                return true;
            }

            lock (_lock)
            {
                if (_otelLogger != null)
                {
                    return true;
                }

                try
                {
                    var environment = EnvironmentHelper.GetEnvironmentName();
                    var resourceBuilder = ResourceBuilder.CreateDefault()
                        .AddService("x21-vsto-addin")
                        .AddAttributes(new Dictionary<string, object>
                        {
                            ["deployment.environment"] = environment,
                            ["host.name"] = Environment.MachineName,
                        });

                    _loggerFactory = LoggerFactory.Create(builder =>
                    {
                        builder.AddOpenTelemetry(options =>
                        {
                            options.SetResourceBuilder(resourceBuilder)
                                .AddOtlpExporter(otlpOptions =>
                                {
                                    otlpOptions.Endpoint = new Uri(GetPostHogLogsEndpoint());
                                    otlpOptions.Headers = $"Authorization=Bearer {apiKey}";
                                    otlpOptions.Protocol = OtlpExportProtocol.HttpProtobuf;
                                });
                        });
                    });

                    _otelLogger = _loggerFactory.CreateLogger("x21-vsto");
                    return true;
                }
                catch (Exception ex)
                {
                    if (IsFatal(ex))
                    {
                        throw;
                    }
                    Logger.LogException(ex);
                    Logger.Info($"PostHog OTLP logger initialization failed: {ex.Message}");
                    return false;
                }
            }
        }

        private void Log(MsLogLevel level, string message, Exception ex)
        {
            if (string.IsNullOrWhiteSpace(message)) return;
            if (!EnsureLogger()) return;

            try
            {
                var userEmail = UserUtils.GetUserEmail();
                var userName = string.IsNullOrWhiteSpace(userEmail)
                    ? Environment.UserName
                    : userEmail;
                var scope = new Dictionary<string, object>
                {
                    ["source"] = "vsto",
                    ["user.name"] = userName,
                    ["deployment.environment"] = EnvironmentHelper.GetEnvironmentName(),
                    ["host.name"] = Environment.MachineName,
                };
                if (!string.IsNullOrWhiteSpace(userEmail))
                {
                    scope["user.email"] = userEmail;
                }

                using (_otelLogger.BeginScope(scope))
                {
                    if (ex != null)
                    {
                        _otelLogger.Log(level, ex, message);
                    }
                    else
                    {
                        _otelLogger.Log(level, message);
                    }
                }
            }
            catch (Exception logEx)
            {
                if (IsFatal(logEx))
                {
                    throw;
                }
                System.Diagnostics.Trace.WriteLine($"PostHog logging failed: {logEx}");
            }
        }

        private static bool IsFatal(Exception ex)
        {
            return ex is OutOfMemoryException
                || ex is StackOverflowException
                || ex is System.Threading.ThreadAbortException
                || ex is System.Threading.ThreadInterruptedException
                || ex is AppDomainUnloadedException;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                lock (_lock)
                {
                    _loggerFactory?.Dispose();
                    _loggerFactory = null;
                    _otelLogger = null;
                }
            }

            _disposed = true;
        }
    }
}
