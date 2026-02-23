using System;

namespace X21.Logging
{
    // Order of levels:
    // Trace
    // Debug
    // Info
    // Warn
    // Error
    // Fatal
    // Off
    public interface ILogger
    {
        void LogException(Exception ex);
        void LogExceptionUI(Exception ex);
        void Debug(object obj);
        void Debug(string message, params object[] args);
        void Info(string message, params object[] args);
        void Info(object obj);
        void Warn(string message, params object[] args);
        void Warn(object obj);
        void Error(string message, params object[] args);
        void Error(object obj);
    }
}
