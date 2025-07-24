using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;

namespace PWQ.Models
{
    public static class VersionHelper
    {
        public static string GetCurrentVersion()
        {
            // Try both base directory and current directory
            string[] possiblePaths = {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VERSION_HISTORY.md"),
                Path.Combine(Directory.GetCurrentDirectory(), "VERSION_HISTORY.md")
            };
            string versionFile = possiblePaths.FirstOrDefault(File.Exists);
            if (versionFile == null)
                return "Unknown";
            var lines = File.ReadAllLines(versionFile);
            foreach (var line in lines)
            {
                // Match version entry: | 1.2.3   | 2025-07-08 | ... |
                var match = Regex.Match(line, @"\|\s*(\d+\.\d+\.\d+)\s*\|\s*(\d{4}-\d{2}-\d{2})\s*\|");
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }
            }
            return "Unknown";
        }
    }
}
