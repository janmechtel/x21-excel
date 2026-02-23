using System;

namespace X21.Utils
{
    public static class UserUtils
    {
        private static readonly object _lock = new object();
        private static string _userEmail;

        public static void SetUserEmail(string email)
        {
            lock (_lock)
            {
                _userEmail = string.IsNullOrWhiteSpace(email) ? null : email;
            }
        }

        public static string GetUserEmail()
        {
            lock (_lock)
            {
                return _userEmail;
            }
        }

        public static string GetUserId()
        {
            return $"{Environment.UserName ?? "unknown"}@{Environment.MachineName ?? "unknown"}";
        }
    }
}
