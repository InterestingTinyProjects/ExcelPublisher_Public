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
        private long _noPublishWarningThreshold;
        private Dictionary<string, string[]> _sheetFormatter;

        public ContentService(long noPublishWarningThreshold, Dictionary<string, string[]> sheetFormatter)
        {
            _noPublishWarningThreshold = noPublishWarningThreshold;
            _sheetFormatter = sheetFormatter;
        }


        public SheetView GetSheetView(GenericCellData cellData)
        {
            return new SheetView
            {
                SheetName = cellData.SheetName,
                HtmlView = GetHtmlView(cellData),
                WarningView = GetNoPublishWarningView(cellData)
            };
        }


        private string GetHtmlView(GenericCellData data)
        {
            string[] formatters = null;
            if (!_sheetFormatter.TryGetValue(data.SheetName, out formatters))
                formatters = null; 

            var builder = new StringBuilder();
            builder.Append($"<p><small class=\"text-muted text-center s-publishTime\" data-publishTime=\"{data.Timestamp.ToString("yyyy -MM-dd HH:mm:ss")}\">Publish Time: {data.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")}</p>");
            builder.Append("<p><table class=\"table\">");
            builder.Append("<tbody>");
            for (int i = 0; i < data.Rows; i++)
            {
                builder.Append("<tr>");
                for (int j = 0; j < data.Columns; j++)
                {
                    builder.Append("<td>");
                    if(data.Data[i, j] == null || data.Data[i, j].GetType() != typeof(double))                
                        builder.Append(data.Data[i, j]);
                    else if(formatters != null && formatters.Length > j)
                    {
                        double doubleVal = (double)data.Data[i, j];
                        var format = formatters[j];
                        builder.Append(doubleVal.ToString(format));
                    }

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
            // It's been too long?
            if ((DateTime.Now - data.Timestamp).TotalMilliseconds <= _noPublishWarningThreshold)
                return "";

            // Return warning
            return @"<div class='alert alert-warning d - flex align - items - center' role='alert'>
                        <i class='bi bi-patch-exclamation'>
                            It has been too long since the last publish.
                        </i>
                    </div>";
        }
    }
}
