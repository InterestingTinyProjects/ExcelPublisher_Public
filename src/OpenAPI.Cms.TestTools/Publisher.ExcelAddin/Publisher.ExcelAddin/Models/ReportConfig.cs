using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using OpenApi.Cms.TestTools.Client.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Publisher.ExcelAddin.Models
{
    /// <summary>
    /// Stores a row in the configuration sheet
    /// </summary>
    public class ReportConfig
    {
        public Worksheet ConfigSheet { get; set; }

        public int RowNumber { get; set; }

        public string ReportName { get; set; }

        public string SheetRange { get; set; }

        public string Sheet 
        {
            get
            {
                var index = SheetRange.IndexOf("!");
                if (index <= 0)
                    return SheetRange;

                return SheetRange.Substring(0, index).Trim();
            }
        }

        public string Range 
        {
            get
            {
                var index = SheetRange.IndexOf("!");
                if (index <= 0)
                    return string.Empty;

                return SheetRange.Substring(index + 1, SheetRange.Length - index - 1).Trim();
            }
        }

        private string _formatterString;
        private string[] _formatter;

        public string FormatterString 
        {
            get { return _formatterString; } 
            set
            {
                _formatterString = value;
                _formatter = null;
            } 
        }

        public string[] Formatter 
        { 
            get
            {
                if (_formatter == null)
                {
                    if (string.IsNullOrEmpty(_formatterString))
                        _formatter = null;
                    else
                        _formatter = JsonConvert.DeserializeObject<string[]>(_formatterString);
                }

                return _formatter;
            }
        }

        public Func<bool> PublishFunc { get; set; }

        public long? Interval { get; set; }

        public string DBServer { get; set; }

        public string DBName { get; set; }

        public long? MaxStoredRecords { get; set; }

        public long? InactiveTimeout { get; set; }

        public string DbConnectionString
        {
            get
            {
                return $"data source={DBServer};database={DBName};Integrated Security=SSPI;persist security info=True; Pooling=true;";
            }
        }

        public bool CanPublish()
        {
            if (this.ConfigSheet == null)
                return false;

            var sheetRowNumber = this.RowNumber + 2;
            var cellVal = this.ConfigSheet.Range[$"F{sheetRowNumber}"].Value2 as object;
            if (cellVal == null)
                return false;

            return cellVal.ToString().Equals("true", StringComparison.InvariantCultureIgnoreCase)
                    || cellVal.ToString().Equals("yes", StringComparison.InvariantCultureIgnoreCase);
        }

        public bool CheckDb(StringBuilder message)
        {
            var connString = DbConnectionString;
            try
            {
                // Check DB Connection
                var repo = new WebPublisherRepository(connString);
                var records = repo.TestDBConnection(this.ReportName);
                message.AppendLine($"- DB Connected. {records} records found.");
                return true;
            }
            catch (Exception ex)
            {
                message.AppendLine($"- DB connection failed. Connection: {connString}, Info: {ex.Message}");
                return false;
            }
        }


        public bool Validate(StringBuilder message, Application app)
        {
            // Check Sheet Range configuration
            try
            {
                var range = app.Range[this.SheetRange];
                message.AppendLine($"- Range {this.SheetRange} found");
            }
            catch (Exception ex)
            {
                message.AppendLine($"- Cannot get Range {this.SheetRange}, Info: {ex.Message}");
                return false;
            }

            // Check Formatter
            if (!string.IsNullOrEmpty(this.FormatterString))
            {
                try
                {
                    JsonConvert.DeserializeObject<string[]>(this.FormatterString);
                    message.AppendLine($" - Formatter String OK.");
                }
                catch (Exception ex)
                {
                    message.AppendLine($"- Invalid Formatter String: {this.FormatterString}, Info: {ex.Message}");
                    return false;
                }
            }

            // Check Publish or not
            message.AppendLine($"- Pubish: {this.CanPublish()}");

            return true;
        }
    }
}
