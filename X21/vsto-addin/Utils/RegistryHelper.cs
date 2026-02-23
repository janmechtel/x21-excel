using Microsoft.Win32;
using System;
using X21.Logging;

namespace X21.Utils
{
    /// <summary>
    /// Centralized registry helper for managing X21 preferences
    /// </summary>
    public static class RegistryHelper
    {
        private const string ROOT_KEY = @"SOFTWARE\X21";
        private const string OLD_FIRST_RUN_KEY = @"SOFTWARE\X21\FirstRun"; // Legacy nested key

        /// <summary>
        /// Gets a boolean value from the registry. Returns defaultValue if not found.
        /// Supports both new DWORD format (0/1) and old string format ("false").
        /// For FirstRun specifically, also checks the old nested key location for backward compatibility.
        /// </summary>
        public static bool GetBool(string valueName, bool defaultValue)
        {
            try
            {
                // Special handling for FirstRun to support migration from old format
                if (valueName == "FirstRun")
                {
                    return GetFirstRunValue(defaultValue);
                }

                // Try to read from the new format (DWORD)
                using (var key = Registry.CurrentUser.OpenSubKey(ROOT_KEY))
                {
                    if (key != null)
                    {
                        var value = key.GetValue(valueName);
                        if (value != null)
                        {
                            // Handle DWORD format (0 = false, 1 = true)
                            if (value is int intValue)
                            {
                                return intValue != 0;
                            }

                            // Handle old string format for backward compatibility
                            if (value is string stringValue)
                            {
                                if (stringValue.Equals("false", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Migrate to new format
                                    SetBool(valueName, false);
                                    return false;
                                }
                                if (stringValue.Equals("true", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Migrate to new format
                                    SetBool(valueName, true);
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error reading registry value {valueName}: {ex.Message}");
            }

            return defaultValue;
        }

        /// <summary>
        /// Sets a boolean value in the registry as a DWORD (0 for false, 1 for true)
        /// </summary>
        public static void SetBool(string valueName, bool value)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(ROOT_KEY))
                {
                    if (key != null)
                    {
                        key.SetValue(valueName, value ? 1 : 0, RegistryValueKind.DWord);
                        Logger.Info($"Set registry value {valueName} to {(value ? 1 : 0)}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error setting registry value {valueName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Special handling for FirstRun to support migration from old nested key format.
        /// Old format: SOFTWARE\X21\FirstRun\FirstRun (string "false")
        /// New format: SOFTWARE\X21\FirstRun (DWORD 0)
        /// </summary>
        private static bool GetFirstRunValue(bool defaultValue)
        {
            try
            {
                // First check new format: SOFTWARE\X21\FirstRun as DWORD
                using (var key = Registry.CurrentUser.OpenSubKey(ROOT_KEY))
                {
                    if (key != null)
                    {
                        var value = key.GetValue("FirstRun");
                        if (value != null)
                        {
                            // New format exists - DWORD where 0 = not first run, 1 = first run
                            if (value is int intValue)
                            {
                                return intValue != 0;
                            }
                        }
                    }
                }

                // Check old nested format: SOFTWARE\X21\FirstRun\FirstRun as string
                using (var key = Registry.CurrentUser.OpenSubKey(OLD_FIRST_RUN_KEY))
                {
                    if (key != null)
                    {
                        var value = key.GetValue("FirstRun");
                        if (value != null && value.ToString().Equals("false", StringComparison.OrdinalIgnoreCase))
                        {
                            // Found old format with "false" - migrate to new format
                            Logger.Info("Migrating FirstRun from old string format to new DWORD format");
                            SetBool("FirstRun", false); // Set to 0 (not first run)

                            // Clean up old nested key structure
                            try
                            {
                                Registry.CurrentUser.DeleteSubKeyTree(OLD_FIRST_RUN_KEY, false);
                                Logger.Info("Deleted legacy FirstRun registry key");
                            }
                            catch (Exception deleteEx)
                            {
                                Logger.Info($"Could not delete legacy FirstRun key: {deleteEx.Message}");
                            }

                            return false; // Not first run
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error reading FirstRun value: {ex.Message}");
            }

            return defaultValue;
        }

        /// <summary>
        /// Gets an integer value from the registry. Returns defaultValue if not found.
        /// </summary>
        public static int GetInt(string valueName, int defaultValue)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(ROOT_KEY))
                {
                    if (key != null)
                    {
                        var value = key.GetValue(valueName);
                        if (value is int intValue)
                        {
                            return intValue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error reading registry value {valueName}: {ex.Message}");
            }

            return defaultValue;
        }

        /// <summary>
        /// Sets an integer value in the registry as a DWORD
        /// </summary>
        public static void SetInt(string valueName, int value)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(ROOT_KEY))
                {
                    if (key != null)
                    {
                        key.SetValue(valueName, value, RegistryValueKind.DWord);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error setting registry value {valueName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets a string value from the registry. Returns defaultValue if not found.
        /// </summary>
        public static string GetString(string valueName, string defaultValue)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(ROOT_KEY))
                {
                    if (key != null)
                    {
                        var value = key.GetValue(valueName);
                        if (value != null)
                        {
                            return value.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error reading registry value {valueName}: {ex.Message}");
            }

            return defaultValue;
        }

        /// <summary>
        /// Sets a string value in the registry
        /// </summary>
        public static void SetString(string valueName, string value)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(ROOT_KEY))
                {
                    if (key != null)
                    {
                        key.SetValue(valueName, value, RegistryValueKind.String);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error setting registry value {valueName}: {ex.Message}");
            }
        }
    }
}
