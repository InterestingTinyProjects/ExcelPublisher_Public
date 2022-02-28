using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Cms.WebPublisher.Formatters
{
    public class ConditionalCellFormatter : IAlertRowFormatter
    {
        private string _indexColumnName;
        private Func<object, bool> _condition;
        private string _style;

        public ConditionalCellFormatter(string indexColumnName, Func<object, bool> condition, string style)
        {
            _indexColumnName = indexColumnName;
            _condition = condition;
            _style = style;
        }

        public bool ShouldApply(object[] rowTexts, string[] headers)
        {
            int cancelNColumnIndex = Array.IndexOf(headers, _indexColumnName);
            // No Such column
            if (cancelNColumnIndex < 0 || cancelNColumnIndex >= rowTexts.Length)
                return false;

            // Value is true
            return _condition != null 
                    && _condition(rowTexts[cancelNColumnIndex]);
        }

        public object Apply(object[] rowTexts, string[] headers)
        {
            return _style;
        }
    }
}
