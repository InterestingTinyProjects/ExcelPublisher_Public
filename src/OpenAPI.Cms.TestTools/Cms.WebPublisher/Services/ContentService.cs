using Cms.WebPublisher.Models;
using OpenApi.Cms.TestTools.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cms.WebPublisher.Services
{
    public class ContentService
    {
        /// <summary>
        /// See https://github.com/Excel-DNA/ExcelDna/blob/661aba734f08b2537866632f2295a77062640672/Source/ExcelDna.Integration/ExcelError.cs
        /// and https://groups.google.com/g/exceldna/c/Z6mmxJ4LSbM
        ///     ErrDiv0 = -2146826281
        ///     ErrNA = -2146826246
        ///     ErrName = -2146826259
        ///     ErrNull = -2146826288
        ///     ErrNum = -2146826252
        ///     ErrRef = -2146826265
        ///     ErrValue = -2146826273
        /// </summary>
        private static readonly Dictionary<long, string> ExcelErrorValues = new Dictionary<long, string>
        {
            {-2146826273, "#Value!" },
            {-2146826281, "#Div0!" },
            {-2146826246, "#NA!" },
            {-2146826259, "#Name!" },
            {-2146826288, "#Null!" },
            {-2146826252, "#Num!" },
            {-2146826265, "#Ref!" },
        };


        public SheetView GetSheetView(GenericCellData cellData)
        {
            return new SheetView
            {
                SheetName = cellData.SheetName,
                HtmlView = GetHtmlView(cellData),
                WarningView = GetNoPublishWarningView(cellData)
            };
        }


        private string GetHtmlView(GenericCellData genericData)
        {
            var builder = new StringBuilder();
            builder.Append($"<p class=\"text-muted text-center s-publishTime\" data-publishTime=\"{genericData.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")}\">Publish Time: {genericData.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")}</p>");
            builder.Append("<p><table class=\"table table-striped table-bordered\">");
            builder.Append("<tbody>");
            for (int i = 0; i < genericData.Rows; i++)
            {
                builder.Append("<tr>");
                for (int j = 0; j < genericData.Columns; j++)
                {
                    // Get Formatter
                    var formtatter = string.Empty;
                    if (genericData.Formatter != null && genericData.Formatter.Length > j)
                        formtatter = genericData.Formatter[j];

                    // Handle <td> tag by formatter
                    if (!HandleCellFormat(builder, formtatter))
                        continue;

                    // Handle values in <td>
                    if (genericData.Data[i, j] == null)
                        builder.Append(genericData.Data[i, j]); 
                    else 
                    {
                        // Show pre-set texts first
                        string presetCelValue = null;
                        if ((genericData.Data[i, j] is long || genericData.Data[i, j] is int)
                            && ExcelErrorValues.TryGetValue((long)genericData.Data[i, j], out presetCelValue))
                            builder.Append(presetCelValue);                     
                        else if ( genericData.Data[i, j] is double && !string.IsNullOrEmpty(formtatter))
                        {
                            // The value is a double and there is a formatter
                            double doubleVal = (double)genericData.Data[i, j];
                            var format = genericData.Formatter[j];
                            builder.Append(doubleVal.ToString(format));
                        }
                        else 
                            builder.Append(genericData.Data[i, j]);
                    }

                    builder.Append("</td>");
                }
                builder.Append("</tr>");
            }
            builder.Append("</tbody>");
            builder.Append("</table></p>");

            return builder.ToString();
        }

        /// <summary>
        /// Generate <td> according to formtter
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="formtatter"></param>
        /// <returns>true - add a td to the table; false - ignore the td</returns>
        private bool HandleCellFormat(StringBuilder builder, string formtatter)
        {
            // Handle Formatter
            if (formtatter == "P" || formtatter.StartsWith("f"))
                builder.Append("<td class=\"s-td-number\">");
            else if (formtatter.Equals("hidden", StringComparison.OrdinalIgnoreCase))
                return false;
            else
                builder.Append("<td>");

            return true;
        }



        private string GetNoPublishWarningView(GenericCellData data)
        {
            if (!data.InactiveTimeout.HasValue)
                return string.Empty;

            // It's been too long?
            if ((DateTime.Now - data.Timestamp).TotalMilliseconds <= data.InactiveTimeout)
                return string.Empty;

            // Return warning
            return @$"<div class='alert alert-warning text-center border border-warning' role='alert'>
                        <blockquote class='blockquote'>
                            <i class='bi bi-patch-exclamation'>
                                It has been too long since the last publish.
                            </i>
                        </blockquote>
                            Last Publish Time: {data.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")}
                    </div>";
        }
    }
}
