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
            builder.Append($"<p><small class=\"text-muted text-center s-publishTime\" data-publishTime=\"{genericData.Timestamp.ToString("yyyy -MM-dd HH:mm:ss")}\">Publish Time: {genericData.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")}</p>");
            builder.Append("<p><table class=\"table\">");
            builder.Append("<tbody>");
            for (int i = 0; i < genericData.Rows; i++)
            {
                builder.Append("<tr>");
                for (int j = 0; j < genericData.Columns; j++)
                {
                    builder.Append("<td>");

                    // The value is a double and there is a formatter
                    if(genericData.Data[i, j] != null && genericData.Data[i, j].GetType() == typeof(double) &&
                        genericData.Formatter != null && genericData.Formatter.Length > j)
                    {
                        double doubleVal = (double)genericData.Data[i, j];
                        var format = genericData.Formatter[j];
                        builder.Append(doubleVal.ToString(format));
                    }
                    else
                        builder.Append(genericData.Data[i, j]);

                    builder.Append("</td>");
                }
                builder.Append("</tr>");
            }
            builder.Append("</tbody>");
            builder.Append("</table></p>");

            return builder.ToString();
        }

        private string GetNoPublishWarningView(GenericCellData data)
        {
            if (!data.InactiveTimeout.HasValue)
                return string.Empty;

            // It's been too long?
            if ((DateTime.Now - data.Timestamp).TotalMilliseconds <= data.InactiveTimeout)
                return string.Empty;

            // Return warning
            return @"<div class='alert alert-warning d - flex align - items - center' role='alert'>
                        <i class='bi bi-patch-exclamation'>
                            It has been too long since the last publish.
                        </i>
                    </div>";
        }
    }
}
