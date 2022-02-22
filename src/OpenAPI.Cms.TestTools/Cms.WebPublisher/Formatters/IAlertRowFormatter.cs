using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cms.WebPublisher.Formatters
{
    public interface IAlertRowFormatter
    {
        bool ShouldApply(object[] rowTexts, string[] headers);

        object Apply(object[] rowTexts, string[] headers);

    }
}
