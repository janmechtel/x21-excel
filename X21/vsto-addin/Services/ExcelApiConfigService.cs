using System;
using System.IO;
using System.Net;
using X21.Logging;
using X21.Utils;

namespace X21.Services
{
    public class ExcelApiConfigService
    {
        private static readonly object _lock = new object();
        private static ExcelApiConfigService _instance;
        private string _baseUrl;
        private int _currentPort;
        private const int DefaultPort = 8080;
        private const string PortFilePrefix = "excel-api-port-";

        private ExcelApiConfigService()
        {
            _baseUrl = $"http://localhost:{DefaultPort}/";
            _currentPort = DefaultPort;
        }

        public static ExcelApiConfigService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ExcelApiConfigService();
                        }
                    }
                }
                return _instance;
            }
        }

        public string BaseUrl => _baseUrl;
        public int CurrentPort => _currentPort;

        public int FindAvailablePort(int startPort = DefaultPort)
        {
            const int maxAttempts = 100;

            for (int i = 0; i < maxAttempts; i++)
            {
                int port = startPort + i;
                if (IsPortAvailable(port))
                {
                    return port;
                }
            }

            throw new Exception($"No available port found starting from {startPort}");
        }

        private bool IsPortAvailable(int port)
        {
            try
            {
                using (var listener = new HttpListener())
                {
                    listener.Prefixes.Add($"http://localhost:{port}/");
                    listener.Start();
                    listener.Stop();
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"Port availability check failed for {port}: {ex.Message}");
                return false;
            }
        }

        public void SetPort(int port)
        {
            _currentPort = port;
            _baseUrl = $"http://localhost:{port}/";
            WritePortToFile(port);
        }

        private void WritePortToFile(int port)
        {
            try
            {
                // Write to the same X21 app data directory as the Deno server
                var appDataDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "X21"
                );

                // Ensure the directory exists
                if (!Directory.Exists(appDataDir))
                {
                    Directory.CreateDirectory(appDataDir);
                }

                var environment = EnvironmentHelper.GetEnvironmentName();
                var fileName = $"{PortFilePrefix}{environment}";
                var filePath = Path.Combine(appDataDir, fileName);

                File.WriteAllText(filePath, port.ToString());
                Logger.Info($"Excel API port {port} written to: {filePath}");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Failed to write Excel API port file: {ex.Message}");
            }
        }
    }
}
