using Magicodes.ExporterAndImporter.Pdf;
using Microsoft.Extensions.DependencyInjection;
using Volo.Abp.Modularity;

namespace Magicodes.IE.Excel.Abp
{
    public class MagicodesIEPdfModule : AbpModule
    {
        public override void ConfigureServices(ServiceConfigurationContext context)
        {
            context.Services.AddMagicodesPdfExporter();
        }
    }
}
