using Magicodes.ExporterAndImporter.Csv;
using Microsoft.Extensions.DependencyInjection;
using System;
using Volo.Abp.Modularity;

namespace Magicodes.IE.Excel.Abp
{
    public class MagicodesIECsvModule: AbpModule
    {
        public override void ConfigureServices(ServiceConfigurationContext context)
        {
            context.Services.AddScoped<ICsvExporter, CsvExporter>();

            context.Services.AddScoped<ICsvImporter, CsvImporter>();
        }
    }
}
