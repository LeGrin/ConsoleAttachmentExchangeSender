using System;
using System.Diagnostics;
using System.IO;

namespace SimpleExchangeEmailConsole
{
    public static class Logger
    {
        public static bool IsLoggingEnabed { get; set; }

        private static readonly string LogFolder =
            Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\Logs\";

        public static void Info(string message, bool important = false, bool wait = false)
        {
            message = $"INFO      {DateTime.Now} : {message}";
            FileLogger(message);
            if (IsLoggingEnabed || important)
            {
                Console.WriteLine($"{message}");
            }

            if (wait)
            {
                Console.ReadLine();
            }
        }

        public static void Exception(string message, bool wait = false)
        {
            Console.WriteLine($"EXCEPTION {DateTime.Now} : {message}");
            if (wait)
            {
                Console.ReadLine();
            }
        }

        private static void FileLogger(string message)
        {
            string path = $"{LogFolder}/Log {DateTime.Now:yyyy MMMM dd}.txt";
            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine($"INFO      {DateTime.Now} : Created new Log File");
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine($"{message}");
            }
        }

    }
}
