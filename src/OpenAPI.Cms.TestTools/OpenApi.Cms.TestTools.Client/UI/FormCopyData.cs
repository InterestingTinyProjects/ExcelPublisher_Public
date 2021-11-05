using Cms.Confluences;
using Cms.Confluences.Models;
using Newtonsoft.Json;
using OpenApi.Cms.TestTools.Client.DB;
using OpenApi.Cms.TestTools.Client.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenApi.Cms.TestTools.Client.UI
{
    public partial class FormCopyData : Form
    {
        private ConfluenceService _conflenceService = new ConfluenceService("xxq@annageo.com", "n8OzBqdet3TAkotdhzgh9EB8", "annageo.atlassian.net/");
        private WebPublisherRepository _repo = new WebPublisherRepository("data source=(local);database=WebPublisher;Integrated Security=SSPI;persist security info=True;");


        public FormCopyData()
        {
            InitializeComponent();
        }

        public GenericCellData CellData { get; set; }

        public string Message 
        {
            get { return txtData.Text; }
            set { txtData.Text = value; }
        }

        public string Data
        {
            get
            {
                return txtData.Text;
            }

            set
            {
                txtData.Text = value;
            }
        }

        private void btnConfluence_Click(object sender, EventArgs e)
        {
            try
            {
                var html = GetHtmlView();
                var task = UpdateConfuence(html, "43712513")
                    .ContinueWith(t =>
                    {
                        try
                        {
                            var report = t.Result;
                            MessageBox.Show("Published to Confluence", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (AggregateException ex)
                        {
                            MessageBox.Show(ex.InnerExceptions[0].Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }, 
                    CancellationToken.None,
                    TaskContinuationOptions.OnlyOnFaulted,
                    TaskScheduler.FromCurrentSynchronizationContext());
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private async Task<Page> UpdateConfuence(string html, string pageId)
        {
            var pages = await _conflenceService.SearchPages("~255194745", "Test Root");
            if (!pages.IsSuccess)
                throw pages.Exception;

            var page = pages.Content[0];

            var newPage = await _conflenceService.UpdatePage(page, html, page.Title);

            return newPage.Content;
        }


        private string GetHtmlView()
        {
            var data = JsonConvert.DeserializeObject<GenericCellData>(txtData.Text);

            var builder = new StringBuilder();
            builder.Append($"<p> Publish Time: {data.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")}</p>");
            builder.Append("<p><table>");
            for (int i = 0; i < data.Rows; i++)
            {
                builder.Append("<tr>");
                for (int j = 0; j < data.Columns; j++)
                {
                    builder.Append("<td>");
                    builder.Append(data.Data[i,j]);
                    builder.Append("</td>");
                }
                builder.Append("</tr>");
            }
            builder.Append("</table></p>");

            var html = builder.ToString();

            return html;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                PublishPositions(this.CellData);
                MessageBox.Show("Published");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void PublishPositions(GenericCellData data)
        {
            //var publishFolder = Environment.GetEnvironmentVariable("WebPublisher_Folder", EnvironmentVariableTarget.Machine);
            //var target = Path.Combine(publishFolder, "Positions.json");

            //File.WriteAllText(target, content);
            _repo.PublishPositions(data);
        }

        private void FormCopyData_Load(object sender, EventArgs e)
        {

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtData.Text = "";
        }
    }
}
