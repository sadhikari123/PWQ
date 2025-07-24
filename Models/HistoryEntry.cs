using System;
using System.Collections.Generic;

namespace PWQ.Models
{
    public class HistoryEntry
    {
        public string Timestamp { get; set; }
        public string UserID { get; set; }
        public string ConfigFile { get; set; }
        public string Operation { get; set; } // "ADD", "EDIT", "DELETE"
        public string RowKey { get; set; }
        public string ChangeSummary { get; set; }
        public string OldValues { get; set; } // JSON string of old row data (for EDIT and DELETE)
        public string NewValues { get; set; } // JSON string of new row data (for ADD and EDIT)
    }
}
