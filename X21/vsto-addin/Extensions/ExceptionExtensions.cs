using X21.Logging;
using System;

namespace X21.Extensions
{
    public static class ExceptionExtensions
    {
        public static void TraceToDebugLog(this Exception ex)
        {
            Logger.Info($"{ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
        }
    }
}
