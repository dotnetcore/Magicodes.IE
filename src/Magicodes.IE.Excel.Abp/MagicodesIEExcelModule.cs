using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Microsoft.Extensions.DependencyInjection;
using System;
using Volo.Abp.Modularity;

namespace Magicodes.IE.Excel.Abp
{
    public class MagicodesIEExcelModule: AbpModule
    {
        public override void ConfigureServices(ServiceConfigurationContext context)
        {
            context.Services.AddScoped<IExcelExporter, ExcelExporter>();

            context.Services.AddScoped<IExcelImporter, ExcelImporter>();

            context.Services.AddScoped<IExportFileByTemplate, ExcelExporter>();

            //TODO:处理筛选器
        }
    }
}
