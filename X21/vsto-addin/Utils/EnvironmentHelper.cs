using System;
using System.IO;
using System.Reflection;

namespace X21.Utils
{
    /// <summary>
    /// Helper class to get environment-specific paths and names
    /// </summary>
    public static class EnvironmentHelper
    {
        private static string _environmentName;

        /// <summary>
        /// Gets the environment name from the assembly name (e.g., "Dev", "Staging", "Production")
        /// </summary>
        public static string GetEnvironmentName()
        {
            if (string.IsNullOrEmpty(_environmentName))
            {
                // Get assembly name which will be like "X21-Dev", "X21-Staging", etc.
                var assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

                // Extract environment from assembly name
                // Format: X21-{Environment}
                if (assemblyName.StartsWith("X21-"))
                {
                    _environmentName = assemblyName.Substring(4); // Remove "X21-" prefix
                }
                else
                {
                    // Fallback for legacy builds without environment suffix
                    _environmentName = "Production";
                }
            }
            return _environmentName;
        }

        /// <summary>
        /// Gets the environment-specific Excel log path
        /// </summary>
        public static string GetExcelLogPath()
        {
            var env = GetEnvironmentName();
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "X21",
                $"X21-xls-{env}",
                "Logs"
            );
        }

        /// <summary>
        /// Gets the environment-specific Deno log path
        /// </summary>
        public static string GetDenoLogPath()
        {
            var env = GetEnvironmentName();
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "X21",
                $"X21-deno-{env}",
                "Logs"
            );
        }

        /// <summary>
        /// Gets a display-friendly name for the current environment
        /// </summary>
        public static string GetDisplayName()
        {
            return $"X21-{GetEnvironmentName()}";
        }

        /// <summary>
        /// Checks if running in Debug mode (F5 from Visual Studio)
        /// Returns true for Debug builds, false for Release/Published builds
        /// </summary>
        public static bool IsDebugMode()
        {
#if DEBUG
            return true;
#else
            return false;
#endif
        }

        /// <summary>
        /// Gets the current version of the X21 add-in
        /// </summary>
        public static string GetVersion()
        {
            try
            {
                if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
                {
                    return System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"Failed to read deployed version: {ex}");
            }

            // Fallback to assembly version
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                return assembly.GetName().Version.ToString();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"Failed to read assembly version: {ex}");
                return "1.0.0.0";
            }
        }
    }
}
