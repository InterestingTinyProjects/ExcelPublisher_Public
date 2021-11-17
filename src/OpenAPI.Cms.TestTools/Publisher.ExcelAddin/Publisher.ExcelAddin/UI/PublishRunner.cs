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
            try
            {
                if (_isRunning)
                    return;

                if (!_reportConfig.CanPublish())
                    return;

                _isRunning = true;
                 var data = GetCellData(_sheet, _reportConfig.Range);
                 PublishToDB(data);
                _isRunning = false;
            }
            catch (Exception ex)
            {
                _isRunning = false;
            }
        }

        /// <summary>
        /// Get data ona sheet
        /// </summary>
        /// <param name="fromSheet"></param>
        /// <returns></returns>
        private GenericCellData GetCellData(Worksheet fromSheet, string rangeName)
        {
            var sheetName = fromSheet.Name;
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
        private int PublishToDB(GenericCellData cellData)
        {
            var dbConn = _reportConfig.DbConnectionString;
            var repo = new WebPublisherRepository(dbConn);
            return repo.PublishPositions(cellData);
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
