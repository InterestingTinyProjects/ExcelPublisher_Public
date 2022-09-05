using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Publisher.ExcelAddin.Models
{
    public static class ExcelExtensions
    {
        public static ExcelReference ToExcelReference(this Range range)
        {
            return new ExcelReference(range.Row - 1,
                                        range.Row - 1 + range.Rows.Count - 1,
                                        range.Column - 1,
                                        range.Column - 1 + range.Columns.Count - 1,
                                        range.Worksheet.Name);

        }

    }
}
