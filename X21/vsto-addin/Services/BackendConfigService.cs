using System;
using System.IO;
using System.Threading.Tasks;
using X21.Logging;
using X21.Utils;

namespace X21.Services
{
    public class BackendConfigService
    {
        private static readonly object _lock = new object();
        private static BackendConfigService _instance;
        private string _baseUrl;
        private string _websocketUrl;
        private const int DefaultPort = 8000;
        private const int DefaultWebSocketPort = 8000;
        private const string PortFilePrefix = "deno-server-port-";
        private const string WebSocketPortFilePrefix = "deno-websocket-port-";

        private BackendConfigService()
        {
            _baseUrl = $"http://localhost:{DefaultPort}";
            _websocketUrl = $"ws://localhost:{DefaultWebSocketPort}/ws";
        }

        public static BackendConfigService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new BackendConfigService();
                        }
                    }
                }
                return _instance;
            }
        }

        public string BaseUrl
        {
            get
            {
                TryUpdateBaseUrl();
                return _baseUrl;
            }
        }

        public string WebSocketUrl
        {
            get
            {
                TryUpdateWebSocketUrl();
                return _websocketUrl;
            }
        }

        private string GetCurrentEnvironment()
        {
            // Get the actual environment from the assembly name (e.g., "Debug", "Dev", "Staging", "Production")
            return EnvironmentHelper.GetEnvironmentName();
        }

        private void TryUpdateBaseUrl()
        {
            var port = TryGetPortFromFile(PortFilePrefix, DefaultPort, "backend");
            var newBaseUrl = $"http://localhost:{port}";
            if (_baseUrl != newBaseUrl)
            {
                _baseUrl = newBaseUrl;
                Logger.Info($"Updated backend base URL to: {_baseUrl}");
            }
        }

        private void TryUpdateWebSocketUrl()
        {
            var port = TryGetPortFromFile(WebSocketPortFilePrefix, DefaultWebSocketPort, "WebSocket");
            var newWebSocketUrl = $"ws://localhost:{port}/ws";
            if (_websocketUrl != newWebSocketUrl)
            {
                _websocketUrl = newWebSocketUrl;
                Logger.Info($"Updated WebSocket URL to: {_websocketUrl}");
            }
        }

        private int TryGetPortFromFile(string portFilePrefix, int defaultPort, string serviceName)
        {
            try
            {
                var appDataDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "X21"
                );

                if (Directory.Exists(appDataDir))
                {
                    var portFiles = Directory.GetFiles(appDataDir, $"{portFilePrefix}*");
                    var bestPortFile = GetBestMatchingPortFile(portFiles, portFilePrefix);

                    if (bestPortFile != null)
                    {
                        var port = ReadPortFromFile(bestPortFile);
                        if (port.HasValue)
                        {
                            Logger.Info($"Found {serviceName} port {port.Value} from {Path.GetFileName(bestPortFile)}");
                            return port.Value;
                        }
                    }
                }

                var backendDir = Path.GetDirectoryName(PathResolver.GetBackendExecutablePath());
                if (backendDir != null)
                {
                    var fallbackFiles = Directory.GetFiles(backendDir, $"{portFilePrefix}*");
                    if (fallbackFiles.Length > 0)
                    {
                        var port = ReadPortFromFile(fallbackFiles[0]);
                        if (port.HasValue)
                        {
                            Logger.Info($"Found {serviceName} port {port.Value} from backend directory");
                            return port.Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Could not read {serviceName} port file: {ex.Message}");
            }

            Logger.Info($"Using default {serviceName} port: {defaultPort}");
            return defaultPort;
        }

        private string GetBestMatchingPortFile(string[] portFiles, string portFilePrefix)
        {
            if (portFiles.Length == 0) return null;

            var currentEnv = GetCurrentEnvironment();

            foreach (var portFile in portFiles)
            {
                var fileName = Path.GetFileName(portFile);
                if (fileName.EndsWith($"-{currentEnv}") || fileName.Equals($"{portFilePrefix}{currentEnv}"))
                {
                    return portFile;
                }
            }

            var mostRecentFile = portFiles[0];
            var mostRecentTime = File.GetLastWriteTime(mostRecentFile);

            for (int i = 1; i < portFiles.Length; i++)
            {
                var writeTime = File.GetLastWriteTime(portFiles[i]);
                if (writeTime > mostRecentTime)
                {
                    mostRecentTime = writeTime;
                    mostRecentFile = portFiles[i];
                }
            }

            return mostRecentFile;
        }

        private int? ReadPortFromFile(string filePath)
        {
            try
            {
                var portText = File.ReadAllText(filePath).Trim();
                int port;
                if (int.TryParse(portText, out port))
                {
                    return port;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"Failed to read port file '{filePath}': {ex}");
                return null;
            }
        }
    }
}
