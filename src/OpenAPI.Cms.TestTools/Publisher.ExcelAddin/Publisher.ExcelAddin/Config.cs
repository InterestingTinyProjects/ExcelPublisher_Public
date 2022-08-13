using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Publisher.ExcelAddin.Models;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenApi.Cms.TestTools.Client
{
    public class Config
    {
        private static readonly string StatusCell = "A1";
        private static readonly int RowNumerColumn = 1;
        private static readonly int ReportNameColumn = 2;
        private static readonly int SheetRangeColumn = 4;
        private static readonly int FormatterColumn = 5;
        private static readonly int PublishColumn = 6;
        private static readonly int IntervalColumn = 7;
        private static readonly int DBServerColumn = 8;
        private static readonly int DBNameColumn = 9;
        private static readonly int DataPersistenceSecondsColumn = 10;
        private static readonly int InactiveTimeoutColumn = 11;
        private static readonly int FilenameColumn = 12;

        public string ConfigSheetName
        {
            get
            {
                var val = ConfigurationManager.AppSettings["ConfigSheet"];
                if (string.IsNullOrEmpty(val))
                    throw new Exception("Cannot find 'ConfigSheet' in the .confg file");

                return val;
            }
        }

        public LogEventLevel LogLevel
        {
            get
            {
                var val = ConfigurationManager.AppSettings["LogLevel"];
                var level = LogEventLevel.Error;
                if (!Enum.TryParse(val, out level))
                    return LogEventLevel.Error;

                return level;
            }
        }

        public ReportConfig[] GetConfigurations(Worksheet configSheet)
        {
            var val = configSheet.UsedRange.Value2 as object[,];
            var cols = val.GetLength(1);
            var rows = val.Length / cols;

            var reports = new List<ReportConfig>();
            for (int i = 1; i <= rows; i++)
            {
                int rowNumber;
                var strRowNumber = val[i, RowNumerColumn]?.ToString();
                if (!int.TryParse(strRowNumber, out rowNumber))
                    continue;

                reports.Add(new ReportConfig
                {
                    ConfigSheet = configSheet,
                    RowNumber = rowNumber,
                    ReportName = val[i, ReportNameColumn]?.ToString(),
                    SheetRange = val[i, SheetRangeColumn]?.ToString(),
                    FormatterString = val[i, FormatterColumn]?.ToString(),
                    Interval = Parse(val[i, IntervalColumn]),
                    DBServer = val[i, DBServerColumn]?.ToString(),
                    DBName = val[i, DBNameColumn]?.ToString(),
                    MaxStoredRecords = Parse(val[i, DataPersistenceSecondsColumn]),
                    InactiveTimeout = Parse(val[i, InactiveTimeoutColumn]),
                    Filename = val[i, FilenameColumn]?.ToString()
                });
            }

            return reports.ToArray();
        }


        public void PublishStarted(Worksheet configSheet, int countOfRunning)
        {
            configSheet.Range[StatusCell].Value2 = $"RUNNING({countOfRunning})...";
        }

        public void PublishStopped(Worksheet configSheet)
        {
            configSheet.Range[StatusCell].Value2 = "Stopped";
        }


        private long? Parse(object val)
        {
            if (val == null)
                return null;

            long longVal;
            if (!long.TryParse(val.ToString(), out longVal))
                return null;

            return longVal;

        }
    }
}
