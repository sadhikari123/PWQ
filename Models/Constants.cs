using System.Collections.Generic;

namespace PWQ.Models
{
    public static class Constants
    {
        public static readonly string[] SystemFields = { "KEY", "userID", "timestamp" };
        
        public static readonly Dictionary<string, string> CsvFiles = new Dictionary<string, string>
        {
            { "PWQ Litho Layers", @"\\ORSHFS.intel.com\ORAnalysis$\1278_MAODATA\LITHO_LAYER\PWQ\PWQ_litho_layers.csv" }
        };
        
        // Archive and history file for PWQ
        public static readonly string ArchivePath = @"\\ORSHFS.intel.com\ORAnalysis$\1278_MAODATA\LITHO_LAYER\PWQ\Archive";
        public static readonly string HistoryFileName = "edit_history.csv";
    }
}
