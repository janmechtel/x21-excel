using X21.Extensions;
using X21.Logging;
using System;

namespace X21.Utils
{
    public static class Execute
    {
        public enum CatchMode
        {
            // Will not catch exceptions at all, rarely used.
            Fail,

            // Default case, log and show the ui message
            LogShowUi,

            // Normal exception that will be logged
            LogFileOnly,

            // For high number exceptions that are expected and where we can't do anything about.
            DontLog,
        }

        public static void Call(System.Action action, CatchMode mode = CatchMode.LogShowUi)
        {
            switch (mode)
            {
                case CatchMode.Fail:
                    action();
                    break;
                case CatchMode.LogShowUi:
                    try
                    {
                        action();
                    }
                    catch (System.Threading.ThreadAbortException)
                    {
                        // Ignore
                    }
                    catch (Exception ex)
                    {
                        Logger.Instance.LogException(ex);
                    }
                    break;
                case CatchMode.LogFileOnly:
                    try
                    {
                        action();
                    }
                    catch (System.Threading.ThreadAbortException)
                    {
                        // Ignore
                    }
                    catch (Exception ex)
                    {
                        ex.TraceToDebugLog();
                    }
                    break;
                case CatchMode.DontLog:
                    try
                    {
                        action();
                    }
                    catch
                    {
                        // ignore
                    }
                    break;
            }
        }

        public static T Call<T>(Func<T> func, CatchMode mode = CatchMode.LogShowUi)
        {
            switch (mode)
            {
                case CatchMode.Fail:
                    return func();
                case CatchMode.LogShowUi:
                    try
                    {
                        return func();
                    }
                    catch (Exception ex)
                    {
                        Logger.Instance.LogException(ex);
                    }
                    break;
                case CatchMode.DontLog:
                    try
                    {
                        return func();
                    }
                    catch
                    {
                    }
                    break;
            }

            return default(T);
        }
    }
}
