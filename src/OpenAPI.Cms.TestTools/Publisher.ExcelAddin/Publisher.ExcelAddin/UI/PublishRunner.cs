using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using OpenApi.Cms.TestTools.Client.DB;
using OpenApi.Cms.TestTools.Client.Models;
using Publisher.ExcelAddin.Models;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Publisher.ExcelAddin.UI
{
    public class PublishRunner: IDisposable
    {
        private System.Timers.Timer _timer;
        private bool _isRunning;
        private ReportConfig _reportConfig;
        private Worksheet _sheet;

        public PublishRunner(ReportConfig reportConfig, Worksheet sheet)
        {
            _reportConfig = reportConfig;
            _sheet = sheet;
        }

        public void Start()
        {
            if ( _timer != null && _timer.Enabled == true)
                return;

            _isRunning = false;
            SetTimer();

            _timer.Enabled = true;
            _timer.Start();
        }


        public void Stop()
        {
            _timer.Stop();
            _timer.Enabled = false;
        }

        /// <summary>
        /// Set the timer
        /// </summary>
        private void SetTimer()
        {
            var interval = _reportConfig.Interval.HasValue? Convert.ToDouble(_reportConfig.Interval.Value): double.MaxValue;
            _timer = new System.Timers.Timer();
            _timer.Enabled = false;
            _timer.Interval = interval;       
            _timer.Elapsed += (s, e) => {               
                PublishSheets();
            };
        }

        private void PublishSheets()
        {
            var reportName = this._reportConfig.ReportName;
            try
            {
                // publishing data directly.
                // COM may reject access and throw an Exception when COM is no Ready when cursor is outside Excel ("0x8001010A Apllication is busy. "). In this case, a message will be sent to Excel to register a callback to publish data in the below catch Exception section
                Log.Information($"Publishing {reportName}... ");
                PublishData();
                Log.Information("Done.");
            }
            catch (COMException comEx)
            {
                // In case of COM rejects access, push a WM_SYNCMACRO windows event to Excel and Excel will callback the PublishData function when COM is ready
                // "0x8001010A Apllication is busy" errors were happened when accessing COM at the time Excel is not "Ready". And the only suggestion by COM is "RETRY LATER"...
                // To avoid such conflicts, COM actions are synchronized by posting a WM_SYNCMACRO windows message. Excel will then receive the message from Windows in its event loop when it is "Ready" and perform the action.
                // This feature is avaible in ExcelDNA - ExcelAsyncUtil.QueueAsMacro. Details: https://docs.excel-dna.net/performing-asynchronous-work/
                // https://stackoverflow.com/questions/25434845/disposing-of-exceldnautil-application-from-new-thread

                ExcelAsyncUtil.QueueAsMacro((state) =>
                {
                    try
                    {
                        Log.Information($"WM_SYNCMACRO Publishing {state}...");
                        PublishData();
                        Log.Information("Done.");
                    }
                    catch (Exception e)
                    {
                        _isRunning = false;
                        Log.Logger.Error("Publish by Excel Message Error: " + e.Message + e.StackTrace);
                    }
                }, reportName);
                Log.Logger.Information($"Sent message WM_SYNCMACRO to Excel... {reportName}");
            }
            catch (Exception ex)
            {
                _isRunning = false;
                Log.Logger.Error("Publish Error: " + ex.Message + ex.StackTrace);
            }
        }

        private void PublishData()
        {
            try
            {
                // Check Can Publish
                if (!_reportConfig.CanPublish())
                    return;

                // Collect data and publish
                _isRunning = true;
                var data = GetCellData(_sheet, _reportConfig.Range, _reportConfig.DataTimeTag);

                // Publish to DB in an async way
                PublishToDBAsync(data);
                _isRunning = false;
            }
            catch
            {
                _isRunning = false;
                throw;
            }
        }



        /// <summary>
        /// Get data ona sheet
        /// </summary>
        /// <param name="fromSheet"></param>
        /// <returns></returns>
        private GenericCellData GetCellData(Worksheet fromSheet, string rangeName, string dataTimeTagAddress)
        {
            var dataTimeTag = string.Empty;
            if (!string.IsNullOrEmpty(dataTimeTagAddress))
            {
                var dataTimeTagRange = fromSheet.Range[dataTimeTagAddress];
                dataTimeTag = dataTimeTagRange.Text;
            }

            // read report data
            var usedRange = string.IsNullOrEmpty(rangeName) ? fromSheet.UsedRange : fromSheet.Range[rangeName];
            var val = usedRange.Value2 as object[,];

            var cols = val.GetLength(1);
            var rows = val.Length / cols;

            return new GenericCellData
            {
                Timestamp = DateTime.Now,
                Rows = rows,
                Columns = cols,
                Data = val,
                SheetName = _reportConfig.ReportName,
                Formatter = _reportConfig.Formatter,
                InactiveTimeout = _reportConfig.InactiveTimeout,
                MaxStoredRecords = _reportConfig.MaxStoredRecords,
                DataTimeTag = dataTimeTag
            };

            //var form = new FormCopyData();
            //form.Data = JsonConvert.SerializeObject(data, Formatting.Indented);
            //form.CellData = data;
            //form.ShowDialog();
        }

        /// <summary>
        /// Publish to DB
        /// </summary>
        /// <param name="cellData"></param>
        /// <returns></returns>
        /// 
        private void PublishToDBAsync(GenericCellData cellData)
        {
            try
            {
                var dbConn = _reportConfig.DbConnectionString;
                var repo = new WebPublisherRepository(dbConn);
                repo.PublishPositions(cellData);

            }
            catch(Exception ex)
            {
                Log.Logger.Error("Save to Database Error. " + ex.Message);
                throw;
            };
        }

        public void Dispose()
        {
            if (_timer != null)
            {
                _timer.Dispose();
                _timer = null;
            }

            GC.SuppressFinalize(this);
        }
    }
}
