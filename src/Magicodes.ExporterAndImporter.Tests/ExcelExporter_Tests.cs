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

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Extensions;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using Magicodes.ExporterAndImporter.Tests.Models.Export.ExportByTemplate_Test1;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using Shouldly;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
            if (!entities.Any()) return dt;

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

        [Fact(DisplayName = "DTO特性导出（测试格式化以及列头索引）")]
        public async Task AttrsExport_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(AttrsExport_Test)}.xlsx");

            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            foreach (var item in data)
            {
                item.LongNo = 458752665;
            }

            var result = await exporter.Export(filePath, data);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells[sheet.Dimension.Address].Rows.ShouldBe(101);
                sheet.Cells["A2"].Text.ShouldBe(data[0].Text2);

                //[ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
                sheet.Cells["E2"].Text.Equals(DateTime.Parse(sheet.Cells["E2"].Text).ToString("yyyy-MM-dd"));

                //[ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]
                sheet.Cells["F2"].Text.Equals(DateTime.Parse(sheet.Cells["F2"].Text).ToString("yyyy-MM-dd HH:mm:ss"));

                //默认DateTime
                sheet.Cells["G2"].Text.Equals(DateTime.Parse(sheet.Cells["G2"].Text).ToString("yyyy-MM-dd"));

                //单元格宽度测试
                sheet.Column(7).Width.ShouldBe(100);

                sheet.Tables.Count.ShouldBe(1);

                var tb = sheet.Tables.First();
                tb.Columns.Count.ShouldBe(9);
                tb.Columns.First().Name.ShouldBe("普通文本");

                sheet.Tables.First();
                tb.Columns.Count.ShouldBe(9);
                tb.Columns[2].Name.ShouldBe("加粗文本");
            }
        }

        [Fact(DisplayName = "导出字段顺序测试")]
        public async Task ExportByColumnIndex_Test()
        {
            var exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportByColumnIndex_Test)}.xlsx");
            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<Issue179>(100);
            var result = await exporter.Export(filePath, data);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Tables.Count.ShouldBe(1);

                var tb = sheet.Tables.First();
                tb.Columns.Count.ShouldBe(typeof(Issue179).GetProperties().Where(p => !p.GetAttribute<ExporterHeaderAttribute>().IsIgnore).Count());
                tb.Columns.First().Name.ShouldBe("员工姓名");
                tb.Columns[1].Name.ShouldBe("料号");
            }
        }

        [Fact(DisplayName = "空数据导出")]
        public async Task AttrsExportWithNoData_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(AttrsExportWithNoData_Test)}.xlsx");

            DeleteFile(filePath);

            var data = new List<ExportTestDataWithAttrs>();
            var result = await exporter.Export(filePath, data);

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                pck.Workbook.Worksheets.First().Cells[pck.Workbook.Worksheets.First().Dimension.Address].Rows
                    .ShouldBe(1);
            }
        }

        [Fact(DisplayName = "全局居中数据导出测试")]
        public async Task AttrExportWithAutoCenterData_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(AttrExportWithAutoCenterData_Test)}.xlsx");

            DeleteFile(filePath);

            var result = await exporter.Export(filePath, GenFu.GenFu.ListOf<ExportTestDataWithAutoCenter>());

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                pck.Workbook.Worksheets.First().Cells[pck.Workbook.Worksheets.First().Dimension.Address].Rows
                    .ShouldBe(26);
                pck.Workbook.Worksheets.First()
                    .Cells[1, 1, 10, 2].Style.HorizontalAlignment.ShouldBe(ExcelHorizontalAlignment.Center);
                pck.Workbook.Worksheets.First()
                    .Cells[2, 2, 10, 2].Style.HorizontalAlignment.ShouldBe(ExcelHorizontalAlignment.Center);
            }
        }

        [Fact(DisplayName = "居中数据导出测试")]
        public async Task AttrExportWithColAutoCenterData_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(AttrExportWithAutoCenterData_Test)}.xlsx");

            DeleteFile(filePath);

            var result = await exporter.Export(filePath, GenFu.GenFu.ListOf<ExportTestDataWithColAutoCenter>());

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                pck.Workbook.Worksheets.First().Cells[pck.Workbook.Worksheets.First().Dimension.Address].Rows
                    .ShouldBe(26);
                pck.Workbook.Worksheets.First()
                    .Cells[1, 1, 10, 2].Style.HorizontalAlignment.ShouldBe(ExcelHorizontalAlignment.Center);
            }
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
                pck.Workbook.Worksheets.First().Cells[pck.Workbook.Worksheets.First().Dimension.Address].Rows
                    .ShouldBe(101);
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
                //请不要使用索引（NET461和.NET Core的Sheet索引值不一致）
                var lastSheet = pck.Workbook.Worksheets.Last();
                lastSheet.Cells[lastSheet.Dimension.Address].Rows.ShouldBe(100);
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
                //请不要使用索引（NET461和.NET Core的Sheet索引值不一致）
                var lastSheet = pck.Workbook.Worksheets.Last();
                lastSheet.Cells[lastSheet.Dimension.Address].Rows.ShouldBe(3);
            }
        }

        [Fact(DisplayName = "头部筛选器测试")]
        public async Task ExporterHeaderFilter_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"{nameof(ExporterHeaderFilter_Test)}.xlsx");

            #region 通过筛选器修改列名

            if (File.Exists(filePath)) File.Delete(filePath);

            var data1 = GenFu.GenFu.ListOf<ExporterHeaderFilterTestData1>();
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

            #region 通过筛选器修改忽略列

            if (File.Exists(filePath)) File.Delete(filePath);
            var data2 = GenFu.GenFu.ListOf<ExporterHeaderFilterTestData2>();
            result = await exporter.Export(filePath, data2);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Dimension.Columns.ShouldBe(5);
            }

            #endregion 通过筛选器修改忽略列
        }

        [Fact(DisplayName = "DataTable结合DTO导出Excel")]
        public async Task DynamicExport_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExport_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(1000);
            var dt = exportDatas.ToDataTable();
            var result = await exporter.Export<ExportTestDataWithAttrs>(filePath, dt);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Dimension.Columns.ShouldBe(9);
            }
        }

        [Fact(DisplayName = "DataTable结合DTO自定义行开始位置导出Excel")]
        public async Task DynamicExportCustomRowStartIndex_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExportCustomRowStartIndex_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrsCustomRowStartIndex>(1000);
            var dt = exportDatas.ToDataTable();
            var result = await exporter.Export<ExportTestDataWithAttrsCustomRowStartIndex>(filePath, dt);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Dimension.Columns.ShouldBe(9);
                sheet.Dimension.Rows.ShouldBe(1005);
            }
        }

        [Fact(DisplayName = "DataTable结合DTO类型导出ByteArray Excel")]
        public async Task DynamicExport_ByteArray_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExport_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(1000);
            var dt = exportDatas.ToDataTable();
            var result = await exporter.ExportAsByteArray<ExportTestDataWithAttrs>(dt);
            result.ShouldNotBeNull();
            using (var file = File.OpenWrite(filePath))
            {
                file.Write(result, 0, result.Length);
            }

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Dimension.Columns.ShouldBe(9);
            }
        }

        [Fact(DisplayName = "DataTable结合Type类型导出ByteArray Excel")]
        public async Task DynamicExportByType_ByteArray_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExport_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(1000);
            var dt = exportDatas.ToDataTable();
            var result = await exporter.ExportAsByteArray(dt, typeof(ExportTestDataWithAttrs));
            result.ShouldNotBeNull();
            using (var file = File.OpenWrite(filePath))
            {
                file.Write(result, 0, result.Length);
            }

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Dimension.Columns.ShouldBe(9);
            }
        }

        [Fact(DisplayName = "DataTable导出Excel（无需定义类，支持列筛选器和表拆分）")]
        public async Task DynamicDataTableExport_Test()
        {
            IExcelExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicDataTableExport_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(50);
            var dt = new DataTable();
            //创建带列名和类型名的列
            dt.Columns.Add("Text", Type.GetType("System.String"));
            dt.Columns.Add("Name", Type.GetType("System.String"));
            dt.Columns.Add("Number", Type.GetType("System.Decimal"));
            dt = EntityToDataTable(dt, exportDatas);
            //加个筛选器导出
            var result = await exporter.Export(filePath, dt, new DataTableTestExporterHeaderFilter(), 10);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //判断Sheet拆分
                pck.Workbook.Worksheets.Count.ShouldBe(5);
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells["C1"].Value.ShouldBe("数值");
                sheet.Dimension.Columns.ShouldBe(3);
            }
        }

#if DEBUG

        [Fact(DisplayName = "大量数据导出Excel", Skip = "本地Debug模式下跳过，太费时")]
#else
        [Fact(DisplayName = "大量数据导出Excel")]
#endif
        public async Task Export100000Data_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(Export100000Data_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await exporter.Export(filePath, GenFu.GenFu.ListOf<ExportTestData>(100000));
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "DTO导出")]
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

        [Fact(DisplayName = "DTO导出支持动态类型")]
        public async Task ExportAsByteArraySupportDynamicType_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportAsByteArraySupportDynamicType_Test)}.xlsx");

            DeleteFile(filePath);

            var source = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>();
            string fields = "text,number,name";
            var shapedData = source.ShapeData(fields) as ICollection<ExpandoObject>;

            var result = await exporter.ExportAsByteArray<ExpandoObject>(shapedData);
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            File.WriteAllBytes(filePath, result);
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "导出分割当前Sheet追加Column")]
        public async Task ExprotSeparateByColumn_Test()
        {
            var exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExprotSeparateByColumn_Test)}.xlsx");

            DeleteFile(filePath);

            var list1 = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>();

            var list2 = GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(30);

            var result = await exporter.Append(list1).SeparateByColumn().Append(list2)
                .SeparateByColumn()
                .Append(list2).ExportAppendData(filePath);

            result.ShouldNotBeNull();

            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
            }
        }

        [Fact(DisplayName = "多个sheet导出")]
        public async Task ExportMutiCollection_Test()
        {
            var exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(ExportMutiCollection_Test)}.xlsx");

            DeleteFile(filePath);

            var list1 = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>();

            var list2 = GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(30);

            var result = exporter.Append(list1, "sheet1").SeparateBySheet().Append(list2).ExportAppendData(filePath);

            await result.ShouldNotBeNull();

            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(2);
                pck.Workbook.Worksheets.First().Name
                    .ShouldBe("sheet1");
                pck.Workbook.Worksheets.Last().Name
                    .ShouldBe(typeof(ExportTestDataWithSplitSheet).GetAttribute<ExcelExporterAttribute>().Name);
            }
        }

        [Fact(DisplayName = "多个sheet导出（空数据）")]
        public async Task ExportMutiCollectionWithEmpty_Test()
        {
            var exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(ExportMutiCollectionWithEmpty_Test)}.xlsx");

            DeleteFile(filePath);

            var list1 = new List<ExportTestDataWithAttrs>();

            var list2 = new List<ExportTestDataWithSplitSheet>();

            var result = exporter.Append(list1).SeparateBySheet().Append(list2).ExportAppendData(filePath);
            await result.ShouldNotBeNull();

            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(2);
            }
        }

        [Fact(DisplayName = "通过Dto导出表头")]
        public async Task ExportHeaderAsByteArray_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(ExportHeaderAsByteArray_Test)}.xlsx");

            DeleteFile(filePath);

            var result = await exporter.ExportHeaderAsByteArray(GenFu.GenFu.New<ExportTestDataWithAttrs>());
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            result.ToExcelExportFileInfo(filePath);
            File.Exists(filePath).ShouldBeTrue();

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Name.ShouldBe("测试");
                sheet.Dimension.Columns.ShouldBe(9);
            }
        }

        [Fact(DisplayName = "通过动态传值导出表头")]
        public async Task ExportHeaderAsByteArrayWithItems_Test()
        {
            IExcelExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(ExportHeaderAsByteArrayWithItems_Test)}.xlsx");

            DeleteFile(filePath);
            var arr = new[] { "Name1", "Name2", "Name3", "Name4", "Name5", "Name6" };
            var sheetName = "Test";
            var result = await exporter.ExportHeaderAsByteArray(arr, sheetName);
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            result.ToExcelExportFileInfo(filePath);
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Name.ShouldBe(sheetName);
                sheet.Dimension.Columns.ShouldBe(arr.Length);
            }
        }

#if DEBUG

        [Fact(DisplayName = "大数据动态列导出Excel", Skip = "本地Debug模式下跳过，太费时")]
#else
        [Fact(DisplayName = "大数据动态列导出Excel")]
#endif
        public async Task LargeDataDynamicExport_Test()
        {
            IExcelExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(LargeDataDynamicExport_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(1200000);

            var dt = new DataTable();
            //创建带列名和类型名的列
            dt.Columns.Add("Text", Type.GetType("System.String"));
            dt.Columns.Add("Name", Type.GetType("System.String"));
            dt.Columns.Add("Number", Type.GetType("System.Decimal"));
            dt = EntityToDataTable(dt, exportDatas);

            var result = await exporter.Export(filePath, dt, maxRowNumberOnASheet: 100000);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //判断Sheet拆分
                pck.Workbook.Worksheets.Count.ShouldBe(12);
            }
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

        #region 图片导出

        [Fact(DisplayName = "Excel导出图片测试")]
        public async Task ExportPicture_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportPicture_Test)}.xlsx");
            DeleteFile(filePath);
            var data = GenFu.GenFu.ListOf<ExportTestDataWithPicture>(5);
            var url = Path.Combine("TestFiles", "ExporterTest.png");
            for (var i = 0; i < data.Count; i++)
            {
                var item = data[i];
                item.Img1 = url;
                if (i == 4)
                    item.Img = null;
                else
                    item.Img = "https://docs.microsoft.com/en-us/media/microsoft-logo-dark.png";
            }

            var result = await exporter.Export(filePath, data);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                //验证Alt
                sheet.Cells["G6"].Value.ShouldBe("404");
                //验证图片
                sheet.Drawings.Count.ShouldBe(9);
                foreach (ExcelPicture item in sheet.Drawings)
                {
                    //检查图片位置
                    new[] { 2, 6 }.ShouldContain(item.From.Column);
                    item.ShouldNotBeNull();
                }

                sheet.Tables.Count.ShouldBe(1);
            }
        }

        [Fact(DisplayName = "Excel导出图片测试自定义开始行位置")]
        public async Task ExportPictureCustomRowStartIndex_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportPictureCustomRowStartIndex_Test)}.xlsx");
            DeleteFile(filePath);
            var data = GenFu.GenFu.ListOf<ExportTestDataWithPictureCustomRowStatIndex>(5);
            var url = Path.Combine("TestFiles", "ExporterTest.png");
            for (var i = 0; i < data.Count; i++)
            {
                var item = data[i];
                item.Img1 = url;
                if (i == 4)
                    item.Img = null;
                else
                    item.Img = "https://docs.microsoft.com/en-us/media/microsoft-logo-dark.png";
            }

            var result = await exporter.Export(filePath, data);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                //验证Alt
                sheet.Cells["G9"].Value.ShouldBe("404");
                //验证图片
                sheet.Drawings.Count.ShouldBe(9);
                foreach (ExcelPicture item in sheet.Drawings)
                {
                    //检查图片位置
                    new[] { 2, 6 }.ShouldContain(item.From.Column);
                    item.ShouldNotBeNull();
                }
                sheet.Dimension.Start.Row.ShouldBe(4);
                sheet.Dimension.Rows.ShouldBe(6);
                sheet.Tables.Count.ShouldBe(1);
            }
        }

        #endregion 图片导出

        [Fact(DisplayName = "数据注解导出测试")]
        public async Task ExportTestDataAnnotations_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExportTestDataAnnotations_Test)}.xlsx");
            DeleteFile(filePath);
            var data = GenFu.GenFu.ListOf<ExportTestDataAnnotations>();

            data[0].Number = null;
            var result = await exporter.Export(filePath,
                data);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                var sheet = pck.Workbook.Worksheets.First();

                sheet.Cells["C2"].Text.Equals(DateTime.Parse(sheet.Cells["C2"].Text).ToString("yyyy-MM-dd"));

                sheet.Cells["D2"].Text.Equals(DateTime.Parse(sheet.Cells["D2"].Text).ToString("yyyy-MM-dd"));
                new List<string> { "是", "否" }.ShouldContain(sheet.Cells["G2"].Text);
                sheet.Tables.Count.ShouldBe(1);
                var tb = sheet.Tables.First();

                var enums = typeof(MyEmum).GetEnumDefinitionList();
                var list = new List<string>();
                foreach (var (item1, item2, item3, item4) in enums)
                {
                    list.Add(item1);
                    list.Add(item2.ToString());
                    list.Add(item3);
                    list.Add(item4);
                }
                list.Add("A Test");
                list.Add("B Test");
                list.ShouldContain(sheet.Cells["E2"].Text);

                tb.Columns[0].Name.ShouldBe("Custom列1");
                tb.Columns[1].Name.ShouldBe("列2");
                tb.Columns.Count.ShouldBe(9);
            }
        }

        [Fact(DisplayName = "样式错误测试")]
        public async Task TenExport_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(TenExport_Test)}.xlsx");

            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<GalleryLineExportModel>(100);

            var result = await exporter.Export(filePath, data);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "导出分割当前Sheet追加Rows")]
        public async Task ExprotSeparateByRows_Test()
        {
            var exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExprotSeparateByRows_Test)}.xlsx");

            DeleteFile(filePath);

            var list1 = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>();

            var list2 = GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(30);

            var result = await exporter.Append(list1).SeparateByRow().Append(list2)
                .ExportAppendData(filePath);

            result.ShouldNotBeNull();

            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                var ec = pck.Workbook.Worksheets.First();
                ec.Dimension.Rows.ShouldBe(57);
            }
        }

        [Fact(DisplayName = "导出分割当前Sheet追加Rows和headers")]
        public async Task ExprotSeparateByRowsAndHeaders_Test()
        {
            var exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ExprotSeparateByRowsAndHeaders_Test)}.xlsx");

            DeleteFile(filePath);

            var list1 = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>();

            var list2 = GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(30);

            var result = await exporter.Append(list1).SeparateByRow().AppendHeaders().Append(list2)
                .ExportAppendData(filePath);

            result.ShouldNotBeNull();

            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                var ec = pck.Workbook.Worksheets.First();
                ec.Dimension.Rows.ShouldBe(58);
            }
        }

        /// <summary>
        /// #140 https://github.com/dotnetcore/Magicodes.IE/issues/140
        /// 身份证导出文本格式测试
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "身份证导出文本格式测试")]
        public async Task Issue140_IdCardExport_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(Issue140_IdCardExport_Test)}.xlsx");

            DeleteFile(filePath);

            var list = new List<Issue140_IdCardExportDto>()
            {
                new Issue140_IdCardExportDto()
                {
                    IdCard = "430626111111111111"
                },
                new Issue140_IdCardExportDto()
                {
                    IdCard = "430626111111111111"
                },
                new Issue140_IdCardExportDto()
                {
                    IdCard = "430626111111111111"
                },
            };
            var result = await exporter.Export(filePath, list);

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                pck.Workbook.Worksheets.First().Cells[pck.Workbook.Worksheets.First().Dimension.Address].Rows
                    .ShouldBe(list.Count + 1);
                pck.Workbook.Worksheets.First().Cells["A2"].Text.ShouldBe(list[0].IdCard);
            }
        }

        [Fact(DisplayName = "忽略所有列进行导出测试")]
        public async Task ItThrowsIfIgnoresAllColumnsExport_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = GetTestFilePath($"{nameof(ItThrowsIfIgnoresAllColumnsExport_Test)}.xlsx");
            DeleteFile(filePath);

            Func<Task> f = async () => await exporter.ExportAsByteArray(GenFu.GenFu.ListOf<ExportTestIgnoreAllColumns>());
            var exception = await Assert.ThrowsAsync<ArgumentException>(f);
            exception.Message.ShouldBe("请勿忽略全部表头！");
        }
    }
}