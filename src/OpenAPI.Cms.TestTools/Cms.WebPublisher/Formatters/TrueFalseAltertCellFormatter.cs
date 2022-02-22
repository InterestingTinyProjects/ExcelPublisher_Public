using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cms.WebPublisher.Formatters
{
    public class TrueFalseAltertRowFormatter: IAlertRowFormatter
    {
        private string _indexColumnName;
        private Func<string> _styleFunc;

        public TrueFalseAltertRowFormatter(string indexColumnName, Func<string> styleFunc)
        {
            _indexColumnName = indexColumnName;
            _styleFunc = styleFunc;
        }

        public bool ShouldApply(object[] rowTexts,  string[] headers)
        {
            int cancelNColumnIndex = Array.IndexOf(headers, _indexColumnName);
            // No Such column
            if (cancelNColumnIndex < 0 || cancelNColumnIndex >= rowTexts.Length)
                return false;

            // Value is true
            return rowTexts[cancelNColumnIndex] != null
                && bool.TrueString.Equals(rowTexts[cancelNColumnIndex].ToString(), StringComparison.OrdinalIgnoreCase);
        }

        public object Apply(object[] rowTexts, string[] headers)
        {
            if (_styleFunc != null)
                return _styleFunc();

            return null;
        }
    }
}
