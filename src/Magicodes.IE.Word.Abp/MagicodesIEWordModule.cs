using Magicodes.ExporterAndImporter.Word;
using Microsoft.Extensions.DependencyInjection;
using System;
using Volo.Abp.Modularity;

namespace Magicodes.IE.Excel.Abp
{
    public class MagicodesIEWordModule: AbpModule
    {
        public override void ConfigureServices(ServiceConfigurationContext context)
        {
            context.Services.AddScoped<IWordExporter, WordExporter>();
        }
    }
}
