using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using CsvHelper;
using System.Globalization;


namespace PWQ.Models
{
    public static class HistoryManager
    {
        private static string GetHistoryFilePath()
        {
            return Path.Combine(PWQ.Models.Constants.ArchivePath, PWQ.Models.Constants.HistoryFileName);
        }

        public static void LogAdd(string configFile, Dictionary<string, string> newRow)
        {
            try
            {
                var entry = new HistoryEntry
                {
                    Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    UserID = Environment.UserName,
                    ConfigFile = configFile,
                    Operation = "ADD",
                    RowKey = newRow.ContainsKey("KEY") ? newRow["KEY"] : "N/A",
                    ChangeSummary = "New row added",
                    OldValues = "",
                    NewValues = JsonSerializer.Serialize(newRow)
                };
                
                WriteHistoryEntry(entry);
            }
            catch (Exception ex)
            {
                LoggingUtility.LogError($"Failed to log ADD operation: {ex.Message}");
            }
        }

        public static void LogEdit(string configFile, Dictionary<string, string> oldRow, Dictionary<string, string> newRow)
        {
            try
            {
                var changes = new List<string>();
                var allKeys = oldRow.Keys.Union(newRow.Keys).ToList();
                
                foreach (var key in allKeys)
                {
                    var oldVal = oldRow.ContainsKey(key) ? oldRow[key] : "";
                    var newVal = newRow.ContainsKey(key) ? newRow[key] : "";
                    
                    if (oldVal != newVal && !PWQ.Models.Constants.SystemFields.Contains(key))
                    {
                        changes.Add($"{key}: '{oldVal}' → '{newVal}'");
                    }
                }

                var entry = new HistoryEntry
                {
                    Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    UserID = Environment.UserName,
                    ConfigFile = configFile,
                    Operation = "EDIT",
                    RowKey = newRow.ContainsKey("KEY") ? newRow["KEY"] : (oldRow.ContainsKey("KEY") ? oldRow["KEY"] : "N/A"),
                    ChangeSummary = changes.Count > 0 ? string.Join("; ", changes) : "No significant changes",
                    OldValues = JsonSerializer.Serialize(oldRow),
                    NewValues = JsonSerializer.Serialize(newRow)
                };
                
                WriteHistoryEntry(entry);
            }
            catch (Exception ex)
            {
                LoggingUtility.LogError($"Failed to log EDIT operation: {ex.Message}");
            }
        }

        public static void LogDelete(string configFile, Dictionary<string, string> deletedRow)
        {
            try
            {
                var entry = new HistoryEntry
                {
                    Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    UserID = Environment.UserName,
                    ConfigFile = configFile,
                    Operation = "DELETE",
                    RowKey = deletedRow.ContainsKey("KEY") ? deletedRow["KEY"] : "N/A",
                    ChangeSummary = "Row deleted",
                    OldValues = JsonSerializer.Serialize(deletedRow),
                    NewValues = ""
                };
                
                WriteHistoryEntry(entry);
            }
            catch (Exception ex)
            {
                LoggingUtility.LogError($"Failed to log DELETE operation: {ex.Message}");
            }
        }

        private static void WriteHistoryEntry(HistoryEntry entry)
        {
            try
            {
                string historyPath = GetHistoryFilePath();
                
                // Ensure archive directory exists
                Directory.CreateDirectory(PWQ.Models.Constants.ArchivePath);
                
                bool fileExists = File.Exists(historyPath);
                
                using (var writer = new StreamWriter(historyPath, append: true))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    // Write header if file doesn't exist
                    if (!fileExists)
                    {
                        csv.WriteField("Timestamp");
                        csv.WriteField("UserID");
                        csv.WriteField("ConfigFile");
                        csv.WriteField("Operation");
                        csv.WriteField("RowKey");
                        csv.WriteField("ChangeSummary");
                        csv.WriteField("OldValues");
                        csv.WriteField("NewValues");
                        csv.NextRecord();
                    }
                    
                    // Write the entry
                    csv.WriteField(entry.Timestamp);
                    csv.WriteField(entry.UserID);
                    csv.WriteField(entry.ConfigFile);
                    csv.WriteField(entry.Operation);
                    csv.WriteField(entry.RowKey);
                    csv.WriteField(entry.ChangeSummary);
                    csv.WriteField(entry.OldValues);
                    csv.WriteField(entry.NewValues);
                    csv.NextRecord();
                }
            }
            catch (Exception ex)
            {
                LoggingUtility.LogError($"Failed to write history entry: {ex.Message}");
            }
        }

        public static List<HistoryEntry> LoadHistoryEntries()
        {
            var entries = new List<HistoryEntry>();
            
            try
            {
                string historyPath = GetHistoryFilePath();
                Console.WriteLine($"Loading history from: {historyPath}");
                
                if (!File.Exists(historyPath))
                {
                    Console.WriteLine("History file does not exist, returning empty list");
                    return entries; // Return empty list if file doesn't exist
                }
                
                Console.WriteLine("History file exists, attempting to read...");
                
                // Check if file needs repair by reading first few lines
                var firstLines = File.ReadLines(historyPath).Take(5).ToList();
                bool needsRepair = false;
                
                if (firstLines.Count > 0)
                {
                    var firstLine = firstLines[0];
                    // Check if first line is proper CSV header
                    if (!firstLine.StartsWith("Timestamp,UserID,ConfigFile") && 
                        (firstLine.StartsWith("======") || firstLine.StartsWith("Config:") || !firstLine.Contains(",")))
                    {
                        Console.WriteLine("History file appears corrupted, attempting repair...");
                        needsRepair = true;
                    }
                }
                
                if (needsRepair)
                {
                    RepairHistoryFile();
                }
                
                using (var reader = new StreamReader(historyPath))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    try
                    {
                        csv.Read();
                        csv.ReadHeader();
                        Console.WriteLine("CSV header read successfully");
                        
                        int rowCount = 0;
                        while (csv.Read())
                        {
                            rowCount++;
                            try
                            {
                                var entry = new HistoryEntry
                                {
                                    Timestamp = csv.GetField("Timestamp") ?? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                    UserID = csv.GetField("UserID") ?? "Unknown",
                                    ConfigFile = csv.GetField("ConfigFile") ?? "Unknown",
                                    Operation = csv.GetField("Operation") ?? "Unknown",
                                    RowKey = csv.GetField("RowKey") ?? "N/A",
                                    ChangeSummary = csv.GetField("ChangeSummary") ?? "No details",
                                    OldValues = csv.GetField("OldValues") ?? "",
                                    NewValues = csv.GetField("NewValues") ?? ""
                                };
                                entries.Add(entry);
                            }
                            catch (Exception ex)
                            {
                                LoggingUtility.LogError($"Failed to parse history entry at row {rowCount}: {ex.Message}");
                                Console.WriteLine($"Failed to parse history entry at row {rowCount}: {ex.Message}");
                                // Continue to next entry instead of failing completely
                            }
                        }
                        Console.WriteLine($"Successfully processed {entries.Count} history entries");
                    }
                    catch (Exception csvEx)
                    {
                        LoggingUtility.LogError($"Error reading CSV: {csvEx.Message}");
                        Console.WriteLine($"Error reading CSV: {csvEx.Message}");
                        return entries; // Return what we have so far
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingUtility.LogError($"Failed to load history entries: {ex.Message} - Stack: {ex.StackTrace}");
                Console.WriteLine($"Failed to load history entries: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            
            // Sort by timestamp descending (most recent first) with null safety
            try
            {
                var sortedEntries = entries.Where(e => e != null && !string.IsNullOrEmpty(e.Timestamp))
                                          .OrderByDescending(e => e.Timestamp)
                                          .ToList();
                Console.WriteLine($"Returning {sortedEntries.Count} sorted entries");
                return sortedEntries;
            }
            catch (Exception sortEx)
            {
                LoggingUtility.LogError($"Error sorting entries: {sortEx.Message}");
                Console.WriteLine($"Error sorting entries: {sortEx.Message}");
                return entries; // Return unsorted if sorting fails
            }
        }

        public static void RepairHistoryFile()
        {
            try
            {
                string historyPath = GetHistoryFilePath();
                
                if (!File.Exists(historyPath))
                {
                    Console.WriteLine("History file does not exist, nothing to repair");
                    return;
                }
                
                Console.WriteLine("Repairing history file...");
                
                // Create backup
                string backupPath = historyPath + $".backup_{DateTime.Now:yyyyMMddHHmmss}";
                File.Copy(historyPath, backupPath);
                Console.WriteLine($"Backup created at: {backupPath}");
                
                var validEntries = new List<HistoryEntry>();
                var allLines = File.ReadAllLines(historyPath);
                
                foreach (var line in allLines)
                {
                    // Skip empty lines and corrupted format lines
                    if (string.IsNullOrWhiteSpace(line) || 
                        line.StartsWith("======") || 
                        line.StartsWith("Config:") ||
                        line.StartsWith("Action:") ||
                        line.StartsWith("Key:") ||
                        line.StartsWith("User:") ||
                        line.StartsWith("Changes:") ||
                        line.Contains("=>") ||
                        line.Trim().StartsWith("Operations:") ||
                        line == "Timestamp,UserID,ConfigFile,Operation,RowKey,ChangeSummary,OldValues,NewValues")
                    {
                        continue;
                    }
                    
                    // Try to parse as CSV line
                    try
                    {
                        var parts = ParseCsvLine(line);
                        if (parts.Length >= 6) // At least have the basic fields
                        {
                            var entry = new HistoryEntry
                            {
                                Timestamp = parts[0],
                                UserID = parts[1],
                                ConfigFile = parts[2],
                                Operation = parts[3],
                                RowKey = parts[4],
                                ChangeSummary = parts[5],
                                OldValues = parts.Length > 6 ? parts[6] : "",
                                NewValues = parts.Length > 7 ? parts[7] : ""
                            };
                            validEntries.Add(entry);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Skipping invalid line: {line.Substring(0, Math.Min(50, line.Length))}... Error: {ex.Message}");
                    }
                }
                
                Console.WriteLine($"Found {validEntries.Count} valid entries, rewriting file");
                
                // Rewrite the file with proper format
                using (var writer = new StreamWriter(historyPath, append: false))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    // Write header
                    csv.WriteField("Timestamp");
                    csv.WriteField("UserID");
                    csv.WriteField("ConfigFile");
                    csv.WriteField("Operation");
                    csv.WriteField("RowKey");
                    csv.WriteField("ChangeSummary");
                    csv.WriteField("OldValues");
                    csv.WriteField("NewValues");
                    csv.NextRecord();
                    
                    // Write valid entries
                    foreach (var entry in validEntries)
                    {
                        csv.WriteField(entry.Timestamp);
                        csv.WriteField(entry.UserID);
                        csv.WriteField(entry.ConfigFile);
                        csv.WriteField(entry.Operation);
                        csv.WriteField(entry.RowKey);
                        csv.WriteField(entry.ChangeSummary);
                        csv.WriteField(entry.OldValues);
                        csv.WriteField(entry.NewValues);
                        csv.NextRecord();
                    }
                }
                
                Console.WriteLine($"History file repaired successfully. {validEntries.Count} entries preserved.");
            }
            catch (Exception ex)
            {
                LoggingUtility.LogError($"Failed to repair history file: {ex.Message}");
                Console.WriteLine($"Failed to repair history file: {ex.Message}");
            }
        }
        
        private static string[] ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var currentField = new System.Text.StringBuilder();
            
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // Escaped quote
                        currentField.Append('"');
                        i++; // Skip next quote
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }
            
            result.Add(currentField.ToString());
            return result.ToArray();
        }
    }
}
