using System;
using System.IO;
using System.Deployment.Application;
using X21.Logging;

namespace X21.Utils
{
    /// <summary>
    /// Utility class to resolve file paths in both ClickOnce deployments and regular assembly locations
    /// </summary>
    public static class PathResolver
    {
        /// <summary>
        /// Gets the path to a file, trying multiple possible locations for ClickOnce deployments and regular assemblies
        /// </summary>
        /// <param name="relativePath">Relative path to the file (e.g., "WebAssets/index.html")</param>
        /// <returns>Full path to the file if found, or the first attempted path if not found</returns>
        public static string GetFilePath(string relativePath)
        {
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var assemblyDirectory = Path.GetDirectoryName(assemblyLocation);

            // Try multiple possible locations for ClickOnce deployments and regular assemblies
            string[] possiblePaths = {
                // 1. Shadow copy location (where assembly is actually running from)
                Path.Combine(assemblyDirectory, relativePath),

                // 2. Original deployment location (one level up from shadow copy)
                Path.Combine(Directory.GetParent(assemblyDirectory)?.FullName ?? "", relativePath),

                // 3. AppDomain base directory
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath),

                // 4. Current working directory
                Path.Combine(Environment.CurrentDirectory, relativePath),

                // 5. Try to get from ApplicationDeployment if available (ClickOnce)
                GetClickOnceFilePath(relativePath)
            };

            Logger.Info($"Looking for file '{relativePath}' in the following locations:");
            Logger.Info($"Assembly location: {assemblyLocation}");
            Logger.Info($"Assembly directory: {assemblyDirectory}");

            foreach (string path in possiblePaths)
            {
                if (!string.IsNullOrEmpty(path))
                {
                    Logger.Info($"  Checking: {path}");
                    Logger.Info($"  Exists: {File.Exists(path)}");

                    if (File.Exists(path))
                    {
                        Logger.Info($"  Found file at: {path}");
                        return path;
                    }
                }
            }

            // Return the first path as default (shadow copy location)
            return possiblePaths[0];
        }

        /// <summary>
        /// Gets the directory path for a relative directory, trying multiple possible locations
        /// </summary>
        /// <param name="relativeDirectory">Relative directory path (e.g., "WebAssets")</param>
        /// <returns>Full path to the directory if found, or the first attempted path if not found</returns>
        public static string GetDirectoryPath(string relativeDirectory)
        {
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var assemblyDirectory = Path.GetDirectoryName(assemblyLocation);

            // Try multiple possible locations for ClickOnce deployments and regular assemblies
            string[] possiblePaths = {
                // 1. Shadow copy location (where assembly is actually running from)
                Path.Combine(assemblyDirectory, relativeDirectory),

                // 2. Original deployment location (one level up from shadow copy)
                Path.Combine(Directory.GetParent(assemblyDirectory)?.FullName ?? "", relativeDirectory),

                // 3. AppDomain base directory
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativeDirectory),

                // 4. Current working directory
                Path.Combine(Environment.CurrentDirectory, relativeDirectory),

                // 5. Try to get from ApplicationDeployment if available (ClickOnce)
                GetClickOnceDirectoryPath(relativeDirectory)
            };

            Logger.Info($"Looking for directory '{relativeDirectory}' in the following locations:");
            Logger.Info($"Assembly location: {assemblyLocation}");
            Logger.Info($"Assembly directory: {assemblyDirectory}");

            foreach (string path in possiblePaths)
            {
                if (!string.IsNullOrEmpty(path))
                {
                    Logger.Info($"  Checking: {path}");
                    Logger.Info($"  Exists: {Directory.Exists(path)}");

                    if (Directory.Exists(path))
                    {
                        Logger.Info($"  Found directory at: {path}");
                        return path;
                    }
                }
            }

            // Return the first path as default (shadow copy location)
            return possiblePaths[0];
        }

        /// <summary>
        /// Gets file path from ClickOnce deployment if available
        /// </summary>
        /// <param name="relativePath">Relative path to the file</param>
        /// <returns>Path to file or null if not available</returns>
        private static string GetClickOnceFilePath(string relativePath)
        {
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var deployment = ApplicationDeployment.CurrentDeployment;
                    var dataDirectory = deployment.DataDirectory;

                    // Try the data directory
                    var dataPath = Path.Combine(dataDirectory, relativePath);
                    if (File.Exists(dataPath))
                    {
                        return dataPath;
                    }

                    // Try one level up from data directory
                    var parentDataPath = Path.Combine(
                        Directory.GetParent(dataDirectory)?.FullName ?? "",
                        relativePath);
                    if (File.Exists(parentDataPath))
                    {
                        return parentDataPath;
                    }

                    // Try the deployment directory
                    var deploymentPath = Path.Combine(
                        Directory.GetParent(dataDirectory)?.FullName ?? "",
                        relativePath);
                    if (File.Exists(deploymentPath))
                    {
                        return deploymentPath;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting ClickOnce file path for '{relativePath}': {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Gets directory path from ClickOnce deployment if available
        /// </summary>
        /// <param name="relativeDirectory">Relative directory path</param>
        /// <returns>Path to directory or null if not available</returns>
        private static string GetClickOnceDirectoryPath(string relativeDirectory)
        {
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var deployment = ApplicationDeployment.CurrentDeployment;
                    var dataDirectory = deployment.DataDirectory;

                    // Try the data directory
                    var dataPath = Path.Combine(dataDirectory, relativeDirectory);
                    if (Directory.Exists(dataPath))
                    {
                        return dataPath;
                    }

                    // Try one level up from data directory
                    var parentDataPath = Path.Combine(
                        Directory.GetParent(dataDirectory)?.FullName ?? "",
                        relativeDirectory);
                    if (Directory.Exists(parentDataPath))
                    {
                        return parentDataPath;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting ClickOnce directory path for '{relativeDirectory}': {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Gets the path to the backend executable (x21-backend-{Environment}.exe)
        /// Searches for any file matching the pattern x21-backend*.exe
        /// </summary>
        /// <returns>Full path to the backend executable if found</returns>
        public static string GetBackendExecutablePath()
        {
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var assemblyDirectory = Path.GetDirectoryName(assemblyLocation);

            // Try multiple possible locations for ClickOnce deployments and regular assemblies
            string[] possibleDirectories = {
                // 1. Shadow copy location (where assembly is actually running from)
                assemblyDirectory,

                // 2. Original deployment location (one level up from shadow copy)
                Directory.GetParent(assemblyDirectory)?.FullName,

                // 3. AppDomain base directory
                AppDomain.CurrentDomain.BaseDirectory,

                // 4. Current working directory
                Environment.CurrentDirectory,

                // 5. ClickOnce data directory
                GetClickOnceDeploymentDirectory()
            };

            Logger.Info("Looking for backend executable (x21-backend*.exe) in the following locations:");

            foreach (string directory in possibleDirectories)
            {
                if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                    continue;

                Logger.Info($"  Checking directory: {directory}");

                try
                {
                    var matchingFiles = Directory.GetFiles(directory, "x21-backend*.exe");
                    if (matchingFiles.Length > 0)
                    {
                        var backendPath = matchingFiles[0];
                        Logger.Info($"  Found backend executable at: {backendPath}");
                        return backendPath;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Info($"  Error searching directory: {ex.Message}");
                }
            }

            // Fallback to hardcoded name for backwards compatibility
            Logger.Info("Backend executable not found, falling back to x21-backend.exe");
            return GetFilePath("x21-backend.exe");
        }

        /// <summary>
        /// Checks if the application is running as a ClickOnce deployment
        /// </summary>
        /// <returns>True if running as ClickOnce deployment, false otherwise</returns>
        public static bool IsClickOnceDeployment()
        {
            try
            {
                return ApplicationDeployment.IsNetworkDeployed;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the ClickOnce deployment directory if available
        /// </summary>
        /// <returns>Deployment directory path or null if not available</returns>
        public static string GetClickOnceDeploymentDirectory()
        {
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var deployment = ApplicationDeployment.CurrentDeployment;
                    return deployment.DataDirectory;
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting ClickOnce deployment directory: {ex.Message}");
            }

            return null;
        }
    }
}
