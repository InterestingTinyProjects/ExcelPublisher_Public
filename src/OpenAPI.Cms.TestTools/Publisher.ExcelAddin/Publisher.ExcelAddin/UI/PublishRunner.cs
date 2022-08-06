using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using OpenApi.Cms.TestTools.Client.DB;
using OpenApi.Cms.TestTools.Client.Models;
using Publisher.ExcelAddin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            _timer = new System.Timers.Timer(interval);
            _timer.Enabled = false;
            _timer.Elapsed += (s, e) => {
                PublishSheets();
            };
        }

        private void PublishSheets()
        {
            if (_isRunning)
                return;

            // 0x8001010A errors were happened when accessing COM at the time Excel is not "Ready". And the only suggestion by COM is "RETRY LATER"...
            // To avoid such conflicts, COM actions are synchronized by posting a WM_SYNCMACRO windows message. Excel will then receive the message from Windows in its event loop when it is "Ready" and perform the action. 
            // This feature is avaible in ExcelDNA - ExcelAsyncUtil.QueueAsMacro. Details: https://docs.excel-dna.net/performing-asynchronous-work/
            // https://stackoverflow.com/questions/25434845/disposing-of-exceldnautil-application-from-new-thread
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    // Check Can Publish
                    if (!_reportConfig.CanPublish())
                        return;

                    // Collect data and publish
                    _isRunning = true;
                    var data = GetCellData(_sheet, _reportConfig.Range);

                    // Publish to DB in an async way
                    PublishToDBAsync(data);
                    _isRunning = false;
                }
                catch (Exception ex)
                {
                    _isRunning = false;
                }
            });

        }

        /// <summary>
        /// Get data ona sheet
        /// </summary>
        /// <param name="fromSheet"></param>
        /// <returns></returns>
        private GenericCellData GetCellData(Worksheet fromSheet, string rangeName)
        {
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
                MaxStoredRecords = _reportConfig.MaxStoredRecords
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
            Task.Run(() =>
            {
                var dbConn = _reportConfig.DbConnectionString;
                var repo = new WebPublisherRepository(dbConn);
                repo.PublishPositions(cellData);
            });
        }

        public void Dispose()
        {
            if (_timer != null)
            {
                _timer.Dispose();
                _timer = null;
            }
            _reportConfig = null;
            _sheet = null;

            GC.SuppressFinalize(this);
        }
    }
}
