#if NETCOREAPP
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using OfficeOpenXml;
using Shouldly;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    /// <summary>
    /// 依赖注入Filter测试
    /// </summary>
    public class DIFilter_Tests : TestBase, IDisposable
    {
        private readonly ITestOutputHelper _testOutputHelper;
        public IExcelImporter Importer = new ExcelImporter();
        ServiceCollection services;
        public DIFilter_Tests(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
            //初始化容器
            services = new ServiceCollection();
            //添加注入关系
            services.AddSingleton<IImportResultFilter, ImportResultFilterTest>();
            services.AddSingleton<IImportHeaderFilter, ImportHeaderFilterTest>();
            services.AddSingleton<IExporterHeaderFilter, TestExporterHeaderFilter1>();
            var serviceProvider = services.BuildServiceProvider();
            AppDependencyResolver.Init(serviceProvider);
            _testOutputHelper.WriteLine("DIFilter_Tests");
        }

        [Fact()]
        public void AppDependencyResolverGetService_Test()
        {
            AppDependencyResolver.Current.GetService<IImportResultFilter>().ShouldNotBeNull();
        }

        [Fact(DisplayName = "DI_结果筛选器测试")]
        public async Task DIImportResultFilter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Errors", "数据错误.xlsx");
            var labelingFilePath = Path.Combine(Directory.GetCurrentDirectory(), $"{nameof(DIImportResultFilter_Test)}.xlsx");
            //DIImportResultFilterDataDto1未设置ImportResultFilter属性
            var result = await Importer.Import<DIImportResultFilterDataDto1>(filePath, labelingFilePath);
            File.Exists(labelingFilePath).ShouldBeTrue();
            result.ShouldNotBeNull();
            result.HasError.ShouldBeTrue();
            result.Exception.ShouldBeNull();
            result.ImporterHeaderInfos.ShouldNotBeNull();
            result.ImporterHeaderInfos.Count.ShouldBeGreaterThan(0);

            //由于同时注册多个筛选器，筛选器之间会相互影响
            result.TemplateErrors.Count.ShouldBe(1);

            var errorRows = new List<int>()
            {
                5,6
            };
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex) && p.FieldErrors.ContainsKey("产品代码") &&
                p.FieldErrors.Values.Contains("Duplicate data exists, please check! Where:5，6。"));

            //TODO:检查标注

        }

        [Fact(DisplayName = "DI_导入列头筛选器测试")]
        public async Task ImportHeaderFilter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "导入列头筛选器测试.xlsx");
            var import = await Importer.Import<DIImportHeaderFilterDataDto1>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);

            //检查值映射
            for (int i = 0; i < import.Data.Count; i++)
            {
                if (i < 5)
                {
                    import.Data.ElementAt(i).Gender.ShouldBe(Genders.Man);
                }
                else
                {
                    import.Data.ElementAt(i).Gender.ShouldBe(Genders.Female);
                }
            }

        }

        [Fact(DisplayName = "DI_头部筛选器测试")]
        public async Task ExporterHeaderFilter_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"{nameof(ExporterHeaderFilter_Test)}.xlsx");

#region 通过筛选器修改列名

            if (File.Exists(filePath)) File.Delete(filePath);

            var data1 = GenFu.GenFu.ListOf<DIExporterHeaderFilterTestData1>();
            var result = await exporter.Export(filePath, data1);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells["D1"].Value.ShouldBe("name");
                sheet.Dimension.Columns.ShouldBe(4);
            }

            #endregion 通过筛选器修改列名
        }

        public void Dispose()
        {

            var descriptorToRemove1 = services.FirstOrDefault(d => d.ServiceType == typeof(IImportResultFilter));
            services.Remove(descriptorToRemove1);

            var descriptorToRemove2 = services.FirstOrDefault(d => d.ServiceType == typeof(IImportHeaderFilter));
            services.Remove(descriptorToRemove2);

            var descriptorToRemove3 = services.FirstOrDefault(d => d.ServiceType == typeof(IExporterHeaderFilter));
            services.Remove(descriptorToRemove3);

            AppDependencyResolver.Current.Dispose();
        }
    }
}
#endif