using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using OpenApi.Cms.TestTools.Client;
using OpenApi.Cms.TestTools.Client.DB;
using OpenApi.Cms.TestTools.Client.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Publisher.ExcelAddin
{
    [ComVisible(true)]
    public class PublishTab : ExcelRibbon
    {
        private IRibbonUI _ribbon;
        private Application _app;
        private List<Worksheet> _fromSheets;
        private Dictionary<string, string> _sheetsToPublish;
        private System.Timers.Timer _timer;
        private bool _isRunning;

        /// <summary>
        /// Create the Tab
        /// </summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public override string GetCustomUI(string RibbonID)
        {
            var resourceName = typeof(PublishTab).Namespace + ".UI.PublishTab.xml";
            using (var stream = typeof(PublishTab).Assembly.GetManifestResourceStream(resourceName))
            {
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Initialize the UI
        /// </summary>
        /// <param name="ribbon"></param>
        public void OnLoad(IRibbonUI ribbon)
        {
            // Setup UI
            _ribbon = ribbon;
            _ribbon.ActivateTab(ControlID: "tabPublish");

            // Set Application
            _app = ExcelDnaUtil.Application as Application;
            if (_app.Workbooks.Count == 0)
                _app.Workbooks.Add();

            // Set Timer
            SetTimer();
        }

        /// <summary>
        /// Click the Start button
        /// </summary>
        /// <param name="control"></param>
        public void StartPublishPositions(IRibbonControl control)
        {
            // Load Sheets
            if (_fromSheets == null || !_fromSheets.Any())
            {
                _fromSheets = new List<Worksheet>();
                _sheetsToPublish = Config.SheetsToPublish;
                foreach (Worksheet sheet in _app.Sheets)
                {
                    if (_sheetsToPublish.ContainsKey(sheet.Name))
                        _fromSheets.Add(sheet);
                }
            }

            _timer.Enabled = true;
            _timer.Start();  
        }

        /// <summary>
        /// Click the stop button
        /// </summary>
        /// <param name="control"></param>
        public void StopPublishPositions(IRibbonControl control)
        {
            _timer.Stop();
            _timer.Enabled = false;
            MessageBox.Show($"Publish stopped. ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        public void TestDbConnection(IRibbonControl control)
        {
            try
            {
                var dbConn = Config.DbConnectionString;
                var repo = new WebPublisherRepository(dbConn);
                var ret = repo.TestDBCOnnection();
                MessageBox.Show($"DB Connected. Found {ret} records. \r\n Connection: {Config.DbConnectionString}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(Exception ex)
            {
                MessageBox.Show($"DB connection failed. Connection: {Config.DbConnectionString}, Info: {ex.Message}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        /// <summary>
        /// Set the timer
        /// </summary>
        private void SetTimer()
        {
            var interval = Config.Interval;
            _timer = new System.Timers.Timer(interval);
            _timer.Enabled = false;
            _timer.Elapsed += (s, e) => {
                PublishSheets();
            };
        }

        /// <summary>
        /// Log
        /// </summary>
        /// <param name="message"></param>
        private void ShowMessage(string message)
        {
            //_fromSheet..Cells[0, 0].Value2 = message;
        }


        private void PublishSheets()
        {
            try
            {
                if (_fromSheets == null || !_fromSheets.Any() || _isRunning)
                    return;

                _isRunning = true;
                foreach (var sheet in _fromSheets)
                {
                    string rangeName;
                    if (!_sheetsToPublish.TryGetValue(sheet.Name, out rangeName))
                        rangeName = string.Empty;

                    var data = GetCellData(sheet, rangeName);
                    PublishToDB(data);
                }
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
            var usedRange = string.IsNullOrEmpty(rangeName)?fromSheet.UsedRange : fromSheet.Range[rangeName];
            var val = usedRange.Value2 as object[,];

            var cols = val.GetLength(1);
            var rows = val.Length / cols;

            return new GenericCellData
            {
                Timestamp = DateTime.Now,
                Rows = rows,
                Columns = cols,
                Data = val,
                SheetName = sheetName 
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
            var dbConn = Config.DbConnectionString;
            var repo = new WebPublisherRepository(dbConn);
            return repo.PublishPositions(cellData);
        }
    }
}
