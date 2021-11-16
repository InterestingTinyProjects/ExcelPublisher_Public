using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenApi.Cms.TestTools.Client.Models
{
    public class GenericCellData
    {
        [JsonIgnore]
        public DateTime Timestamp { get; set; }

        [JsonIgnore]
        public int Rows { get; set; }

        [JsonIgnore]
        public int Columns { get; set; }

        [JsonIgnore]
        public string SheetName { get; set; }

        public object[,] Data { get; set; }

        public string[] Formatter { get; set; }
        
        public long? InactiveTimeout { get; set; }

        [JsonIgnore]
        public long? MaxStoredRecords { get; set; }

    }
}
