using System;
using System.Runtime.InteropServices;
using X21.Logging;

namespace X21.Utils
{
    [ComImport]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    /// <summary>
    /// COM message filter to handle "Excel is busy" (RPC_E_CALL_REJECTED) retry behavior.
    /// </summary>
    internal sealed class OleMessageFilter : IOleMessageFilter
    {
        private const int SERVERCALL_ISHANDLED = 0;
        private const int PENDINGMSG_WAITDEFPROCESS = 2;
        private const int SERVERCALL_RETRYLATER = 2;

        public static void Register()
        {
            try
            {
                CoRegisterMessageFilter(new OleMessageFilter(), out _);
                Logger.Info("[OleMessageFilter] Registered COM message filter");
            }
            catch (COMException ex)
            {
                Logger.Info($"[OleMessageFilter] Failed to register (COM exception): {ex.Message}");
            }
            catch (SEHException ex)
            {
                Logger.Info($"[OleMessageFilter] Failed to register (structured exception handling): {ex.Message}");
            }
            catch (ExternalException ex)
            {
                Logger.Info($"[OleMessageFilter] Failed to register (external exception): {ex.Message}");
            }
            catch (Exception ex)
            {
                Logger.Error($"[OleMessageFilter] Unexpected error during registration: {ex}");
                throw;
            }
        }

        public static void Revoke()
        {
            try
            {
                CoRegisterMessageFilter(null, out _);
                Logger.Info("[OleMessageFilter] Revoked COM message filter");
            }
            catch (COMException ex)
            {
                Logger.Info($"[OleMessageFilter] Failed to revoke (COM exception): {ex.Message}");
            }
            catch (SEHException ex)
            {
                Logger.Info($"[OleMessageFilter] Failed to revoke (structured exception handling): {ex.Message}");
            }
            catch (ExternalException ex)
            {
                Logger.Info($"[OleMessageFilter] Failed to revoke (external exception): {ex.Message}");
            }
            catch (Exception ex)
            {
                Logger.Error($"[OleMessageFilter] Unexpected error during revoke: {ex}");
                throw;
            }
        }

        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return SERVERCALL_ISHANDLED;
        }

        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == SERVERCALL_RETRYLATER)
            {
                // Retry in 100ms.
                return 100;
            }

            // Cancel the call.
            return -1;
        }

        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return PENDINGMSG_WAITDEFPROCESS;
        }

        [DllImport("ole32.dll", PreserveSig = true, SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    }
}
