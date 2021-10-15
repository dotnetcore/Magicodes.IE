using Magicodes.ExporterAndImporter.Html;
using Magicodes.ExporterAndImporter.Pdf;
using Microsoft.Extensions.DependencyInjection;
using System;
using Volo.Abp.Modularity;

namespace Magicodes.IE.Excel.Abp
{
    public class MagicodesIEPdfModule: AbpModule
    {
        public override void ConfigureServices(ServiceConfigurationContext context)
        {
            context.Services.AddScoped<IPdfExporter, PdfExporter>();
        }
    }
}
