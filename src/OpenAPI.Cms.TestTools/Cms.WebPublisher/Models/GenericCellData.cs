using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Cms.WebPublisher.Models
{
    public class GenericCellData
    {
        public string SheetName { get; set; }

        public DateTime Timestamp { get; set; }

        public int Rows { get; set; }

        public int Columns { get; set; }


        public object[,] Data { get; set; }

        public string[] Formatter { get; set; }

        public long? InactiveTimeout { get; set; }

        public string DataTimeTag { get; set; }

        public bool HasValue
        {
            get
            {
                return Data != null;
            }
        }

    }
}
