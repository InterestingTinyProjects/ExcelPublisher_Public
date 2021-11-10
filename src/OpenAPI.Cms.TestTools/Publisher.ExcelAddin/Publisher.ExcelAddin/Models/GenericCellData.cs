using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenApi.Cms.TestTools.Client.Models
{
    public class GenericCellData
    {
        public DateTime Timestamp { get; set; }

        public int Rows { get; set; }

        public int Columns { get; set; }

        public string SheetName { get; set; }

        public object[,] Data { get; set; }

    }
}
