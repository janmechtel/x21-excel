using System;

namespace X21.Logging
{
    public class VoidLogger : ILogger
    {
        public void LogException(Exception ex) { }
        public void LogExceptionUI(Exception ex) { }
        public void Debug(object obj) { }
        public void Debug(string message, params object[] args) { }
        public void Info(string message, params object[] args) { }
        public void Info(object obj) { }
        public void Warn(string message, params object[] args) { }
        public void Warn(object obj) { }
        public void Error(string message, params object[] args) { }
        public void Error(object obj) { }
        public void IncIndent() { }
        public void DecIndent() { }
    }
}
