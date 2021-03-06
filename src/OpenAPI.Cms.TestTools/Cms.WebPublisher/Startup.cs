using Cms.WebPublisher.DB;
using Cms.WebPublisher.Models;
using Cms.WebPublisher.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Diagnostics;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Newtonsoft.Json;
using OpenApi.Cms.TestTools.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Cms.WebPublisher
{
    public class Startup
    {
        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        private ContentService _service = new ContentService();

        public void ConfigureServices(IServiceCollection services)
        {
            services.AddSingleton<Config>();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            var config = app.ApplicationServices.GetService<Config>();

            app.UseStaticFiles();

            app.UseDeveloperExceptionPage();

            app.UseRouting();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapGet("/", async context =>
                {
                    var sheetName = context.Request.Query["SheetName"];
                    if (string.IsNullOrEmpty(sheetName))
                        sheetName = config.DefaultSheetName;

                    var templatePath = Path.Combine(env.ContentRootPath, "Pages", "Publish.html");
                    var data = LoadPositions(config.WebPublisherDB, sheetName);
                    var sheetView = _service.GetSheetView(data);
                    var page = new StringBuilder(File.ReadAllText(templatePath))
                                                .Replace("{{SheetName}}", sheetView.SheetName)
                                                .Replace("{{Warning}}", sheetView.WarningView)
                                                .Replace("{{Content}}", sheetView.HtmlView);

                    await context.Response.WriteAsync(page.ToString());
                });

                endpoints.MapGet("/sheets", async context =>
                {
                    try
                    {
                        // Get input parameter
                        var sheetName = context.Request.Query["SheetName"];
                        if (string.IsNullOrEmpty(sheetName))
                            sheetName = config.DefaultSheetName;

                        // Get filter parameter
                        var filter = context.Request.Query["filter"];
                        if (string.IsNullOrEmpty(filter))
                            filter = string.Empty;

                        // Get table view
                        var data = LoadPositions(config.WebPublisherDB, sheetName);
                        var sheetView = _service.GetSheetView(data, filter);

                        var json = JsonConvert.SerializeObject(sheetView);
                        context.Response.ContentType = "application/json";
                        await context.Response.WriteAsync(json);
                    }
                    catch(Exception ex)
                    {
                        var json = JsonConvert.SerializeObject(ex);
                        context.Response.ContentType = "application/json";
                        await context.Response.WriteAsync(json);
                    }
                });

               endpoints.MapGet("/sheets/formatter", async context =>
               {

                   var sheetName = context.Request.Query["SheetName"];
                   if (string.IsNullOrEmpty(sheetName))
                       sheetName = config.DefaultSheetName;

                   // Get table view
                   var data = LoadPositions(config.WebPublisherDB, sheetName);
                   var json = JsonConvert.SerializeObject(data.Formatter, Formatting.Indented);
                   context.Response.ContentType = "application/json";
                   await context.Response.WriteAsync(json);
               });

               endpoints.MapGet("/error", async context =>
                {
                    context.Response.ContentType = "text/html";
                    Exception ex = context.Features.Get<IExceptionHandlerPathFeature>().Error;

                    await context.Response.WriteAsync("<html><head><title>Error</title></head><body>");
                    await context.Response.WriteAsync($"<h3>{ex.Message}</h3>");
                    await context.Response.WriteAsync($"<p>Type: {ex.GetType().FullName}");
                    await context.Response.WriteAsync($"<p>StackTrace: {ex.StackTrace}");
                    await context.Response.WriteAsync("</body></html>");
                });
            });
        }


        private GenericCellData LoadPositions(string dbConnection, string sheetName)
        {
            var repo = new WebPublisherRepository(dbConnection);
            return repo.GetLatestPositions(sheetName);
        }
    }
}
