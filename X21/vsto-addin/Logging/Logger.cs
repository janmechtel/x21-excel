using X21.Utils;
using System;
using System.IO;
using System.Reflection;

namespace X21.Logging
{
    public static class Logger
    {
        static Logger()
        {
            Instance = new VoidLogger();
        }

        public static void LogException(Exception ex)
        {
            Instance.LogException(ex);
        }
        public static void LogExceptionUI(Exception ex)
        {
            Instance.LogExceptionUI(ex);
        }
        public static void Debug(object obj)
        {
            Instance.Debug(obj);
        }
        public static void Debug(string message, params object[] args)
        {
            Instance.Debug(message, args);
        }

        public static void Info(string message, params object[] args)
        {
            Instance.Info(message, args);
        }

        public static void Info(object obj)
        {
            Instance.Info(obj);
        }

        public static void Warn(string message, params object[] args)
        {
            Instance.Warn(message, args);
        }

        public static void Warn(object obj)
        {
            Instance.Warn(obj);
        }

        public static void Error(string message, params object[] args)
        {
            Instance.Error(message, args);
        }

        public static void Error(object obj)
        {
            Instance.Error(obj);
        }

        public static void LogExceptionToDesktop(Exception ex, string fileName)
        {
            Execute.Call(() =>
            {
                var dir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                var file = Path.Combine(dir, fileName);

                using (var writer = File.AppendText(file))
                {
                    writer.WriteLine("----------------------------------------------");
                    writer.WriteLine("");
                    writer.WriteLine("It seems our product didn't start properly :(");
                    writer.WriteLine("Please email this file to support@kontext21.com");
                    writer.WriteLine("");
                    writer.WriteLine("Sorry for the inconvenience.");
                    writer.WriteLine("Your Team kontext21");
                    writer.WriteLine("");
                    writer.WriteLine("----------------------------------------------");

                    Write(writer, ex);

                    if (ex is TargetInvocationException exception)
                    {
                        Write(writer, exception.InnerException);
                    }
                    else if (ex is ReflectionTypeLoadException loadException)
                    {
                        Write(writer, loadException.InnerException);
                    }
                }
            });
        }

        private static void Write(TextWriter writer, Exception ex)
        {
            writer.WriteLine(ex.GetType().Name + ": " + ex.Message);
            if (!string.IsNullOrEmpty(ex.StackTrace))
            {
                writer.WriteLine(ex.StackTrace);
            }
        }

        public static ILogger Instance { get; set; }
    }
}
