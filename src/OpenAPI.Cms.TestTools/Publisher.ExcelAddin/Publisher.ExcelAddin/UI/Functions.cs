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
using System.Text;
using System.Threading.Tasks;

namespace Publisher.ExcelAddin.UI
{
    public static class Functions
    {
        [ExcelFunction(Description = "Load Report Data from DB")]
        public static object P_LoadReport(int takes, string reportName, string dbServer, string dbName)
        {
            try
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

                return data.ToExcelValue();
            }
            catch(Exception ex)
            {
                return $"Error: {ex.Message} {ex.StackTrace}";
            }
        }
    }
}
