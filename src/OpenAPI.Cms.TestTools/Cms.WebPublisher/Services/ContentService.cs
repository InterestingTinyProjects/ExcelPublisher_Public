using Cms.WebPublisher.Formatters;
using Cms.WebPublisher.Models;
using OpenApi.Cms.TestTools.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
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

        /// <summary>
        /// Alert styles to table row
        /// </summary>
        private static readonly IAlertRowFormatter[] AltertFormatters = new IAlertRowFormatter[]
        {
            new ConditionalCellFormatter(  "Prem", 
                                            (cellValue) => {
                                                decimal d;
                                                if( cellValue == null || 
                                                    !decimal.TryParse(cellValue.ToString().TrimEnd('%'), out d))
                                                    return false;
                                                return d > 0;
                                            }, 
                                          "s-alert-prem"),
            new TrueFalseAltertRowFormatter("Cancel n", () => "s-alert-cancel-now"),
            new TrueFalseAltertRowFormatter("Bid n", () => "s-alert-bid-now"),
            new TrueFalseAltertRowFormatter("Bid w", () => "s-alert-bid-with"),
            new TrueFalseAltertRowFormatter("Ask n", () => "s-alert-ask-now"),
            new TrueFalseAltertRowFormatter("Ask w", () => "s-alert-ask-with"),
        };


        public SheetView GetSheetView(GenericCellData cellData, string filter = "")
        {
            return new SheetView
            {
                SheetName = cellData.SheetName,
                HtmlView = GetHtmlView(cellData, filter),
                WarningView = GetNoPublishWarningView(cellData)
            };
        }


        private string GetHtmlView(GenericCellData genericData, string filter)
        {
            var builder = new StringBuilder();
            builder.Append($"<p class=\"text-muted text-center s-publishTime\" data-publishTime=\"{genericData.Timestamp.ToString("yyyy-MM-dd HH:mm:ss.f")}\">Publish Time: {genericData.Timestamp.ToString("yyyy-MM-dd HH:mm:ss.f")}</p>");
            builder.Append("<p><table class=\"table table-striped table-bordered\">");
            builder.Append("<tbody>");
            string[] headers = null;
            for (int i = 0; i < genericData.Rows; i++)
            {
                // Get texts to be shown in a table row
                var texts = GetRowTexts(genericData, i);
                if (i == 0)
                    headers = texts.Select( h => h == null? string.Empty : h.ToString()).ToArray();

                // Apply filter
                if ( i > 0 
                    && !string.IsNullOrEmpty(filter)
                    && !texts.Any( t => t is string && ((string)t).Contains(filter)))
                    continue;

                // Generate html for a table row
                builder.Append("<tr");
                HandleRowFormatters(builder, texts, headers);
                builder.Append(">");
                for (int j = 0; j < genericData.Columns; j++)
                {
                    // Get Formatter
                    var formatter = string.Empty;
                    if (genericData.Formatter != null && genericData.Formatter.Length > j)
                        formatter = genericData.Formatter[j];

                    // Handle <td> tag by formatter
                    if (!HandleCellFormat(builder, formatter))
                        continue;

                    // Show TD
                    HandleCellText(builder, formatter, texts, j);
                    builder.Append("</td>");
                }
                builder.Append("</tr>");
            }
            builder.Append("</tbody>");
            builder.Append("</table></p>");

            return builder.ToString();
        }

        /// <summary>
        /// Add alert style to each row
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="cellTexts"></param>
        /// <param name="headers"></param>
        private void HandleRowFormatters(StringBuilder builder, object[] cellTexts, string[] headers)
        {
            var alterts = AltertFormatters.Where(f => f.ShouldApply(cellTexts, headers));
            if(alterts.Any())
            {
                builder.Append(" class='");
                foreach (var alert in alterts)
                    builder.Append($" {alert.Apply(cellTexts, headers)}");

                builder.Append("'");
            }
        }

        /// <summary>
        /// Return the array containing cell values
        /// </summary>
        /// <param name="genericData"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private object[] GetRowTexts(GenericCellData genericData, int rowIndex)
        {
            var rowTexts = new object[genericData.Columns];
            for (int j = 0; j < genericData.Columns; j++)
            {
                // Get formatter
                var formatter = string.Empty;
                if (genericData.Formatter != null && genericData.Formatter.Length > j)
                    formatter = genericData.Formatter[j];

                // Format Error Values
                string presetCelValue = null;
                if (genericData.Data[rowIndex, j] != null
                    && (genericData.Data[rowIndex, j] is long || genericData.Data[rowIndex, j] is int)
                    && ExcelErrorValues.TryGetValue((long)genericData.Data[rowIndex, j], out presetCelValue))
                    rowTexts[j] = presetCelValue;
                // Format Numbers
                else if (genericData.Data[rowIndex, j] != null
                         && genericData.Data[rowIndex, j] is double
                         && !string.IsNullOrEmpty(formatter))
                {
                    double doubleVal = (double)genericData.Data[rowIndex, j];
                    rowTexts[j] = doubleVal.ToString(formatter);
                }
                else
                    rowTexts[j] = genericData.Data[rowIndex, j];
            }

            return rowTexts;
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

        /// <summary>
        /// Generate content of a <td>
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="formatter"></param>
        /// <param name="texts"></param>
        /// <param name="index"></param>
        private void HandleCellText(StringBuilder builder, string formatter, object[] texts, int index)
        {
            if (texts[index] != null
                && formatter.Equals("alert", StringComparison.OrdinalIgnoreCase)
                && texts[index] != null
                && bool.TrueString.Equals(texts[index].ToString(), StringComparison.OrdinalIgnoreCase))
                builder.Append($"<b class=\"s-alert\">{texts[index]}</b>");
            else
                builder.Append(texts[index]);
        }


        /// <summary>
        /// Show warning
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
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
                            Last Publish Time: {data.Timestamp.ToString("yyyy-MM-dd HH:mm:ss.ff")}
                    </div>";
        }
    }
}
