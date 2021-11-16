using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using OpenApi.Cms.TestTools.Client;
using OpenApi.Cms.TestTools.Client.DB;
using OpenApi.Cms.TestTools.Client.Models;
using Publisher.ExcelAddin.UI;
using System;
using System.Collections.Concurrent;
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
        private ConcurrentDictionary<string, PublishRunner> _runners = new ConcurrentDictionary<string, PublishRunner>();
        private Config _config = new Config();

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
        }

        /// <summary>
        /// Click the Start button
        /// </summary>
        /// <param name="control"></param>
        public void StartPublishPositions(IRibbonControl control)
        {
            if(_runners.Any())
            {
                MessageBox.Show($"Publishing {_runners.Count} reports. Please press the stop button first before restart.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // create runners
                var configSheet = _app.Sheets[_config.ConfigSheetName];
                var reportConfigs = _config.GetConfigurations(configSheet);
                foreach (var reportConfig in reportConfigs)
                {
                    var sheet = _app.Sheets[reportConfig.Sheet];
                    _runners.TryAdd(reportConfig.ReportName, new PublishRunner(reportConfig, sheet));
                }

                _config?.PublishStarted(configSheet, _runners.Count);

                // Start running
                foreach (var runner in _runners.Values)
                    runner.Start();
            }
            catch (Exception ex)
            {
                PopupWarning($"{ex.Message} - {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Click the stop button
        /// </summary>
        /// <param name="control"></param>
        public void StopPublishPositions(IRibbonControl control)
        {
            try
            {
                foreach (var runner in _runners.Values)
                {
                    runner.Stop();
                    runner.Dispose();
                }
                _runners.Clear();

                var configSheet = _app.Sheets[_config.ConfigSheetName];
                _config?.PublishStopped(configSheet);
                MessageBox.Show($"Publish stopped. ", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(Exception ex)
            {
                PopupWarning($"{ex.Message} - {ex.StackTrace}");
            }
        }


        public void TestDbConnection(IRibbonControl control)
        {
            try
            {
                var configSheet = _app.Sheets[_config.ConfigSheetName];
                var message = new StringBuilder();
                var reportConfigs = _config.GetConfigurations(configSheet);
                foreach (var reportConfig in reportConfigs)
                {
                    message.Clear();
                    var validateResult = reportConfig.CheckDb(message);
                    if (validateResult)
                        MessageBox.Show($"{reportConfig.RowNumber}.{reportConfig.ReportName}. \r\n {message}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show($"FAILED. {reportConfig.RowNumber}.{reportConfig.ReportName} \r\n {message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch(Exception ex)
            {
                PopupWarning($"{ex.Message} - {ex.StackTrace}");
            }
        }

        public void CheckConfig(IRibbonControl control)
        {
            try
            {
                var configSheet = _app.Sheets[_config.ConfigSheetName];
                var message = new StringBuilder();
                var reportConfigs = _config.GetConfigurations(configSheet);
                foreach (var reportConfig in reportConfigs)
                {
                    message.Clear();
                    var validateResult = reportConfig.Validate(message, _app);
                    if (validateResult)
                        MessageBox.Show($"{reportConfig.RowNumber}.{reportConfig.ReportName} is ready. \r\n - Sheet {_config.ConfigSheetName}. \r\n {message}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show($"FAILED. {reportConfig.RowNumber}.{reportConfig.ReportName} \r\n - Sheet {_config.ConfigSheetName} \r\n {message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                PopupWarning($"{ex.Message} - {ex.StackTrace}");
            }
        }

        private void PopupWarning(string message)
        {
            MessageBox.Show(message, "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Log
        /// </summary>
        /// <param name="message"></param>
        private void ShowMessage(string message)
        {
            //_fromSheet..Cells[0, 0].Value2 = message;
        } 
    }
}
