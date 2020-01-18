// ======================================================================
// 
//           filename : ExcelExporter_Tests.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using OfficeOpenXml;
using Shouldly;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelExporter_Tests : TestBase
    {
        /// <summary>
        ///     将entities直接转成DataTable
        /// </summary>
        /// <typeparam name="T">Entity type</typeparam>
        /// <param name="entities">entity集合</param>
        /// <returns>将Entity的值转为DataTable</returns>
        private static DataTable EntityToDataTable<T>(DataTable dt, IEnumerable<T> entities)
        {
            if (entities.Count() == 0) return dt;

            var properties = typeof(T).GetProperties();

            foreach (var entity in entities)
            {
                var dr = dt.NewRow();

                foreach (var property in properties)
                    if (dt.Columns.Contains(property.Name))
                        dr[property.Name] = property.GetValue(entity, null);

                dt.Rows.Add(dr);
            }

            return dt;
        }

        [Fact(DisplayName = "特性导出")]
        public async Task AttrsExport_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(AttrsExport_Test)}.xlsx");

            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            foreach (var item in data)
            {
                item.LongNo = long.MaxValue;
            }
            var result = await exporter.Export(filePath, data);

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "数据拆分多Sheet导出")]
        public async Task SplitData_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(SplitData_Test)}-1.xlsx");

            DeleteFile(filePath);

            var result = await exporter.Export(filePath,
                GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(300));

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //验证Sheet数是否为3
                pck.Workbook.Worksheets.Count.ShouldBe(3);
                //检查忽略列
                pck.Workbook.Worksheets.First().Cells["C1"].Value.ShouldBe("数值");
            }

            filePath = GetTestFilePath($"{nameof(SplitData_Test)}-2.xlsx");
            DeleteFile(filePath);

            result = await exporter.Export(filePath,
                GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(299));

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //验证Sheet数是否为3
                pck.Workbook.Worksheets.Count.ShouldBe(3);
            }

            filePath = GetTestFilePath($"{nameof(SplitData_Test)}-3.xlsx");
            DeleteFile(filePath);

            result = await exporter.Export(filePath,
                GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(302));

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //验证Sheet数是否为4
                pck.Workbook.Worksheets.Count.ShouldBe(4);
            }
        }

        //[Fact(DisplayName = "多语言特性导出")]
        //public async Task AttrsLocalizationExport_Test()
        //{
        //    IExporter exporter = new ExcelExporter();
        //    ExcelBuilder.Create().WithColumnHeaderStringFunc(key =>
        //    {
        //        if (key.Contains("文本")) return "Text";

        //        return "未知语言";
        //    }).Build();

        //    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "testAttrsLocalization.xlsx");
        //    if (File.Exists(filePath)) File.Delete(filePath);

        //    var data = GenFu.GenFu.ListOf<AttrsLocalizationTestData>();
        //    var result = await exporter.Export(filePath, data);
        //    result.ShouldNotBeNull();
        //    File.Exists(filePath).ShouldBeTrue();
        //}

        //[Fact(DisplayName = "动态列导出Excel")]
        //public async Task DynamicExport_Test()
        //{
        //    IExporter exporter = new ExcelExporter();
        //    var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExport_Test) + ".xlsx");
        //    if (File.Exists(filePath)) File.Delete(filePath);

        //    var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(1000);

        //    var dt = new DataTable();
        //    //2.创建带列名和类型名的列(两种方式任选其一)
        //    dt.Columns.Add("Text", Type.GetType("System.String"));
        //    dt.Columns.Add("Name", Type.GetType("System.String"));
        //    dt.Columns.Add("Number", Type.GetType("System.Decimal"));
        //    dt = EntityToDataTable(dt, exportDatas);

        //    var result = await exporter.Export<ExportTestDataWithAttrs>(filePath, dt);
        //    result.ShouldNotBeNull();
        //    File.Exists(filePath).ShouldBeTrue();

        //    var dt2 = dt.Copy();
        //    var arrResult = await exporter.ExportAsByteArray<ExportTestDataWithAttrs>(dt2);
        //    arrResult.ShouldNotBeNull();
        //    arrResult.Length.ShouldBeGreaterThan(0);
        //    filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExport_Test) + "_ByteArray.xlsx");
        //    if (File.Exists(filePath)) File.Delete(filePath);
        //    File.WriteAllBytes(filePath, arrResult);
        //    File.Exists(filePath).ShouldBeTrue();
        //}

        [Fact(DisplayName = "大量数据导出Excel")]
        public async Task Export_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(Export_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await exporter.Export(filePath, GenFu.GenFu.ListOf<ExportTestData>(100000));
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "ExportAsByteArray_Test")]
        public async Task ExportAsByteArray_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(ExportAsByteArray_Test)}.xlsx");

            DeleteFile(filePath);

            var result = await exporter.ExportAsByteArray(GenFu.GenFu.ListOf<ExportTestDataWithAttrs>());
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            File.WriteAllBytes(filePath, result);
            File.Exists(filePath).ShouldBeTrue();
        }

        //[Fact(DisplayName = "ExportHeaderAsByteArray_Test")]
        //public async Task ExportHeaderAsByteArray_Test()
        //{
        //    IExporter exporter = new ExcelExporter();

        //    var filePath = GetTestFilePath($"{nameof(ExportHeaderAsByteArray_Test)}.xlsx");

        //    DeleteFile(filePath);

        //    var result = await exporter.ExportHeaderAsByteArray(GenFu.GenFu.New<ExportTestDataWithAttrs>());
        //    result.ShouldNotBeNull();
        //    result.Length.ShouldBeGreaterThan(0);
        //    File.WriteAllBytes(filePath, result);
        //    File.Exists(filePath).ShouldBeTrue();
        //}

        //[Fact(DisplayName = "ExportHeaderAsByteArrayWithItems_Test")]
        //public async Task ExportHeaderAsByteArrayWithItems_Test()
        //{
        //    IExporter exporter = new ExcelExporter();

        //    var filePath = GetTestFilePath($"{nameof(ExportHeaderAsByteArrayWithItems_Test)}.xlsx");

        //    DeleteFile(filePath);

        //    var result =
        //        await exporter.ExportHeaderAsByteArray(new[] { "Name1", "Name2", "Name3", "Name4", "Name5", "Name6" },
        //            "Test");
        //    result.ShouldNotBeNull();
        //    result.Length.ShouldBeGreaterThan(0);
        //    File.WriteAllBytes(filePath, result);
        //    File.Exists(filePath).ShouldBeTrue();
        //    //TODO:Excel读取并验证
        //}

        //[Fact(DisplayName = "大数据动态列导出Excel", Skip = "太慢，默认跳过")]
        //public async Task LargeDataDynamicExport_Test()
        //{
        //    IExporter exporter = new ExcelExporter();
        //    var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(LargeDataDynamicExport_Test) + ".xlsx");
        //    if (File.Exists(filePath)) File.Delete(filePath);

        //    var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(1200000);

        //    var dt = new DataTable();
        //    //创建带列名和类型名的列
        //    dt.Columns.Add("Text", Type.GetType("System.String"));
        //    dt.Columns.Add("Name", Type.GetType("System.String"));
        //    dt.Columns.Add("Number", Type.GetType("System.Decimal"));
        //    dt = EntityToDataTable(dt, exportDatas);

        //    var result = await exporter.Export<ExportTestDataWithAttrs>(filePath, dt);
        //    result.ShouldNotBeNull();
        //    File.Exists(filePath).ShouldBeTrue();
        //}

        [Fact(DisplayName = "Excel模板导出教材订购明细样表")]
        public async Task ExportByTemplate_Test()
        {
            //模板路径
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "2020年春季教材订购明细样表.xlsx");
            //创建Excel导出对象
            IExportFileByTemplate exporter = new ExcelExporter();
            //导出路径
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportByTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);
            //根据模板导出
            await exporter.ExportByTemplate(filePath,
                new TextbookOrderInfo("湖南心莱信息科技有限公司", "湖南长沙岳麓区", "雪雁", "1367197xxxx", null, DateTime.Now.ToLongDateString(),
                    new List<BookInfo>()
                    {
                        new BookInfo(1, "0000000001", "《XX从入门到放弃》", null, "机械工业出版社", "3.14", 100, "备注"),
                        new BookInfo(2, "0000000002", "《XX从入门到放弃》", "张三", "机械工业出版社", "3.14", 100, null),
                        new BookInfo(3, null, "《XX从入门到放弃》", "张三", "机械工业出版社", "3.14", 100, "备注")
                    }),
                tplPath);
        }

        [Fact(DisplayName = "Excel模板大量导出")]
        public async Task ExportByTemplate_Large_Test()
        {
            //导出5000条数据不超过1秒
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "2020年春季教材订购明细样表.xlsx");
            IExportFileByTemplate exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportByTemplate_Large_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var books = new List<BookInfo>();
            for (int i = 0; i < 5000; i++)
            {
                books.Add(new BookInfo(i + 1, "000000000" + i, "《XX从入门到放弃》", "张三", "机械工业出版社", "3.14", 100 + i, "备注"));
            }

            var stopwatch = new Stopwatch();
            stopwatch.Start();
            await exporter.ExportByTemplate(filePath, new TextbookOrderInfo("湖南心莱信息科技有限公司", "湖南长沙岳麓区", "雪雁", "1367197xxxx", "雪雁", DateTime.Now.ToLongDateString(), books), tplPath);
            stopwatch.Stop();
            //执行时间不得超过1秒（受实际执行机器性能影响）,在测试管理器中运行普遍小于400ms
            stopwatch.ElapsedMilliseconds.ShouldBeLessThanOrEqualTo(1000);

        }

        [Fact(DisplayName = "无特性定义导出测试")]
        public async Task ExportTestDataWithoutExcelExporter_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportTestDataWithoutExcelExporter_Test)}.xlsx");
            DeleteFile(filePath);

            var result = await exporter.Export(filePath,
                GenFu.GenFu.ListOf<ExportTestDataWithoutExcelExporter>());
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }
    }
}