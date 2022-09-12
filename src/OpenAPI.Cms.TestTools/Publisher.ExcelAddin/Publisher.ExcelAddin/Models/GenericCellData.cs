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

        public string DataTimeTag { get; set; }

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

    public static class GenericCellDataExt
    {
        public static object[,] ToExcelValue(this GenericCellData data, bool includeDataTimeTag = false, params string[] columnNames)
        {
            if (data == null)
                return new object[0, 0];

            var rows = data.Rows;
            // For the latest data time tag
            if(includeDataTimeTag)
                rows = rows + 1;

            // Get columns
            var columnIndex = new List<int>();
            for (int j = 0; j < data.Columns; j++)
               if(columnNames == null || !columnNames.Any() || columnNames.Contains(data.Data[0,j]))
                    columnIndex.Add(j);

            // create retrun data
            var ret = new object[rows, columnIndex.Count];

            // Set data time tag
            int index = 0;
            if (includeDataTimeTag)
            {
                index = 1;
                ret[0, 0] = data.DataTimeTag;
            }

            // Set values
            for (int i = 0; i < data.Rows; i++)
            {
                for (int j = 0; j < columnIndex.Count; j++)
                    ret[index, j] = data.Data[i, columnIndex[j]];
                index++;
            }
            
            return ret;
        }

        public static object[,] ToExcelValue(this GenericCellData[] data)
        {
            if (!data.Any())
                return new object[0, 0];

            var columns = data.First().Columns + 1;
            var rows = data.Sum(c => c.Rows);
            // For the latest data time tag
            rows = rows + 1;

            var ret = new object[rows, columns];

            // Set data time tag
            ret[0, 0] = "Data Time Tag:";
            ret[0, 1] = data.First().DataTimeTag;

            // Set Values
            int index = 1;
            foreach (var cellData in data)
            {
                for (int i = 0; i < cellData.Rows; i++)
                {
                    //ret[index, 0] = cellData.Timestamp;
                    for (int j = 0; j < cellData.Columns; j++)
                        ret[index, j] = cellData.Data[i, j];
                    index ++;
                }
            }

            return ret;
        }

        public static string[] ToCsvLines(this GenericCellData[] data)
        {
            if (!data.Any())
                return null;

            var columns = data.First().Columns + 1;
            var rows = data.Sum(c => c.Rows);
            string[] lines = new string[rows];
            int index = 0;
            foreach (var cellData in data)
            {
                //skip the first title line
                for (int i = 1; i < cellData.Rows; i++)
                {
                    StringBuilder sb = new StringBuilder(cellData.Timestamp.ToString("yyyy/MM/dd HH:mm:ss.ffff"));
                    for (int j = 0; j < cellData.Columns; j++)
                    {
                        sb.Append(",");
                        sb.Append(cellData.Data[i, j]);
                    }
                    lines[index] = sb.ToString();
                    index++;
                }
            }

            return lines;
        }

    }
}
