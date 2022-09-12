using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using OpenApi.Cms.TestTools.Client;
using OpenApi.Cms.TestTools.Client.DB;
using OpenApi.Cms.TestTools.Client.Models;
using Publisher.ExcelAddin.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Publisher.ExcelAddin.UI
{
    public static class Functions
    {
        private static readonly Config _config = new Config();

        //[ExcelFunction(Description = "Get Fairs from DB")]
        //public static object __GetFairs(string reportName, string columns)
        //{
        //    if (string.IsNullOrEmpty(reportName))
        //        reportName = "Position";

        //    if (string.IsNullOrEmpty(columns))
        //        columns =  "-, Fair";

        //    GenericCellData data = LoadReport(reportName);
        //    return data.ToExcelValue(true, columns);
        //    //return data;
        //}


        [ExcelFunction(Description = "Get report from DB")]
        public static object __GetReport(string reportName, string columnsObj)
        {
            GenericCellData data = LoadReport(reportName);
            var columns = Array.Empty<string>();
            if(!string.IsNullOrEmpty(columnsObj))
                columns = columnsObj?.Split(',').Select(s => s.Trim())
                                                .Where(s => !string.IsNullOrEmpty(s))
                                                .ToArray();

            return data.ToExcelValue(true, columns);
        }

        //[ExcelFunction(Description = "Load a report from DB to excel")]
        //public static object __LoadReportToExcel(string reportName, params object[] columnsObj)
        //{
        //    try
        //    {
        //        var columns = columnsObj.OfType<string>().ToList();

        //        // Get Report data           
        //        GenericCellData data = LoadReport(reportName);
        //        if(data == null)
        //            return reportName + ": Not Found";

        //        // Show report
        //        var cellValues = data.ToExcelValue(false, columns.ToArray());

        //        var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;

        //        var reference = new ExcelReference(
        //            caller.RowFirst + 1, caller.RowFirst + data.Rows,
        //            caller.ColumnFirst, caller.ColumnFirst + data.Columns - 1);

        //        // Cells are written via this async task
        //        ExcelAsyncUtil.QueueAsMacro(() => { reference.SetValue(cellValues); });

        //        return data.DataTimeTag;
        //    }
        //    catch (Exception ex)
        //    {
        //        return $"Error: {ex.Message} {ex.StackTrace}";
        //    }
        //}



        /// <summary>
        /// Load report data from DB
        /// </summary>
        /// <param name="takes">how many to take</param>
        /// <param name="reportName">Report name.</param>
        /// <param name="dbServer">DB Server</param>
        /// <param name="dbName">DB Name</param>
        /// <returns></returns>
        [ExcelFunction(Description = "Load Report Data from DB")]
        public static object __LoadReport(int takes, string reportName, string dbServer, string dbName)
        {
            try
            {
                GenericCellData[] data = LoadReport(takes, reportName, dbServer, dbName);
                return data.ToExcelValue();
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message} {ex.StackTrace}";
            }
        }

        public static string[] ReportCsvLines(int takes, string reportName, string dbServer, string dbName)
        {
            try
            {
                GenericCellData[] data = LoadReport(takes, reportName, dbServer, dbName);
                return data.ToCsvLines();
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        public static GenericCellData LoadReport(string reportName)
        {
            var repo = new WebPublisherRepository(_config.WebPublisherDB);
            GenericCellData[] data = repo.GetReportData(reportName, 1, false);

            return data.FirstOrDefault();
        }


        public static GenericCellData[] LoadReport(int takes, string reportName, string dbServer, string dbName)
        {
            if (string.IsNullOrEmpty(reportName))
                return null;

            if (takes <= 0)
                takes = 1;

            // Load Config
            var reportConfig = new ReportConfig
            {
                ReportName = reportName,
                DBServer = dbServer,
                DBName = dbName
            };

            var repo = new WebPublisherRepository(reportConfig.DbConnectionString);
            GenericCellData[] data = repo.GetReportData(reportConfig.ReportName, takes);

            return data;
        }
    }
}
