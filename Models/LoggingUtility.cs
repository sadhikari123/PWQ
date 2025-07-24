using System;
using System.IO;

namespace PWQ.Models
{
    public static class LoggingUtility
    {
        private static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NPIReportConfigEditor",
            "logs",
            $"app_{DateTime.Now:yyyyMMdd}.log"
        );
        
        static LoggingUtility()
        {
            // Ensure log directory exists
            var logDir = Path.GetDirectoryName(LogFilePath);
            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }
        }
        
        public static void LogInfo(string message)
        {
            Log("INFO", message);
        }
        
        public static void LogError(string message, Exception ex = null)
        {
            Log("ERROR", message + (ex != null ? $" Exception: {ex}" : ""));
        }
        
        public static void LogWarning(string message)
        {
            Log("WARNING", message);
        }
        
        private static void Log(string level, string message)
        {
            try
            {
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}";
                
                // Output to console for debugging
                Console.WriteLine(logEntry);
                
                // Write to log file
                File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently fail if logging fails
            }
        }
    }
}
