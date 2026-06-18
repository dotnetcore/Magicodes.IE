using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using OfficeOpenXml;
using Shouldly;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    /// <summary>
    /// Issue #617 回归测试：多列表模板导出数据错乱
    /// 覆盖场景：单表格、多表格不同行、多表格同行、各种数据量、Cell Writer 偏移等
    /// </summary>
    public class Issue617_MultiTableTemplateExport_Tests : TestBase
    {
        #region 测试数据模型

        public class OrderItem
        {
            public string ProductName { get; set; }
            public int Quantity { get; set; }
            public decimal Price { get; set; }
        }

        public class OrderLog
        {
            public string Action { get; set; }
            public string Operator { get; set; }
            public string Time { get; set; }
        }

        public class MultiTableDto
        {
            public string Title { get; set; }
            public string ReportDate { get; set; }
            public List<OrderItem> OrderItems { get; set; }
            public List<OrderLog> OrderLogs { get; set; }
            public string Summary { get; set; }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 创建模板 Excel 文件，包含两个表格区域和静态单元格。
        /// 模板布局：
        ///   Row 1: A1={{Title}}, C1={{ReportDate}}
        ///   Row 2: A2={{Table>>OrderItems|ProductName}} B2={{Quantity}} C2={{Price|>>Table}}
        ///   Row 3: A3={{Table>>OrderLogs|Action}} B3={{Operator}} C3={{Time|>>Table}}
        ///   Row 4: A4={{Summary}}
        /// </summary>
        private string CreateTwoTableTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_Template_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = "{{Title}}";
                sheet.Cells[1, 3].Value = "{{ReportDate}}";
                sheet.Cells[2, 1].Value = "{{Table>>OrderItems|ProductName}}";
                sheet.Cells[2, 2].Value = "{{Quantity}}";
                sheet.Cells[2, 3].Value = "{{Price|>>Table}}";
                sheet.Cells[3, 1].Value = "{{Table>>OrderLogs|Action}}";
                sheet.Cells[3, 2].Value = "{{Operator}}";
                sheet.Cells[3, 3].Value = "{{Time|>>Table}}";
                sheet.Cells[4, 1].Value = "{{Summary}}";
                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        /// <summary>
        /// 创建模板，两个表格在同一行（#296 场景）
        ///   Row 1: A1={{Title}}
        ///   Row 2: A2={{Table>>OrderItems|ProductName}} B2={{Quantity}} C2={{Price|>>Table}}
        ///          D2={{Table>>OrderLogs|Action}} E2={{Operator}} F2={{Time|>>Table}}
        ///   Row 3: A3={{Summary}}
        /// </summary>
        private string CreateSameRowTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_SameRow_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = "{{Title}}";
                sheet.Cells[2, 1].Value = "{{Table>>OrderItems|ProductName}}";
                sheet.Cells[2, 2].Value = "{{Quantity}}";
                sheet.Cells[2, 3].Value = "{{Price|>>Table}}";
                sheet.Cells[2, 4].Value = "{{Table>>OrderLogs|Action}}";
                sheet.Cells[2, 5].Value = "{{Operator}}";
                sheet.Cells[2, 6].Value = "{{Time|>>Table}}";
                sheet.Cells[3, 1].Value = "{{Summary}}";
                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        /// <summary>
        /// 创建模板，两个表格之间有 Cell Writer
        ///   Row 1: A1={{Title}}
        ///   Row 2: A2={{Table>>OrderItems|ProductName}} B2={{Quantity}} C2={{Price|>>Table}}
        ///   Row 3: A3={{Summary}}
        ///   Row 4: A4={{Table>>OrderLogs|Action}} B4={{Operator}} C4={{Time|>>Table}}
        ///   Row 5: A5={{ReportDate}}
        /// </summary>
        private string CreateCellBetweenTablesTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_CellBetween_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = "{{Title}}";
                sheet.Cells[2, 1].Value = "{{Table>>OrderItems|ProductName}}";
                sheet.Cells[2, 2].Value = "{{Quantity}}";
                sheet.Cells[2, 3].Value = "{{Price|>>Table}}";
                sheet.Cells[3, 1].Value = "{{Summary}}";
                sheet.Cells[4, 1].Value = "{{Table>>OrderLogs|Action}}";
                sheet.Cells[4, 2].Value = "{{Operator}}";
                sheet.Cells[4, 3].Value = "{{Time|>>Table}}";
                sheet.Cells[5, 1].Value = "{{ReportDate}}";
                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        /// <summary>
        /// 创建单表格模板
        ///   Row 1: A1={{Title}}
        ///   Row 2: A2={{Table>>OrderItems|ProductName}} B2={{Quantity}} C2={{Price|>>Table}}
        ///   Row 3: A3={{Summary}}
        /// </summary>
        private string CreateSingleTableTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_Single_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = "{{Title}}";
                sheet.Cells[2, 1].Value = "{{Table>>OrderItems|ProductName}}";
                sheet.Cells[2, 2].Value = "{{Quantity}}";
                sheet.Cells[2, 3].Value = "{{Price|>>Table}}";
                sheet.Cells[3, 1].Value = "{{Summary}}";
                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        private static List<OrderItem> CreateOrderItems(int count)
        {
            var items = new List<OrderItem>();
            for (int i = 1; i <= count; i++)
            {
                items.Add(new OrderItem
                {
                    ProductName = $"Product{i}",
                    Quantity = i * 10,
                    Price = i * 99.5m
                });
            }
            return items;
        }

        private static List<OrderLog> CreateOrderLogs(int count)
        {
            var logs = new List<OrderLog>();
            for (int i = 1; i <= count; i++)
            {
                logs.Add(new OrderLog
                {
                    Action = $"Action{i}",
                    Operator = $"User{i}",
                    Time = $"2024-01-{i:D2}"
                });
            }
            return logs;
        }

        private void CleanupFile(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        #endregion

        #region 1. 单表格基础场景

        [Theory(DisplayName = "单表格-不同数据量导出正确")]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(5)]
        [InlineData(10)]
        public async Task SingleTable_VariousCounts_ShouldExportCorrectly(int itemCount)
        {
            var tplPath = CreateSingleTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_Single_{itemCount}.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Test Report",
                    OrderItems = CreateOrderItems(itemCount),
                    Summary = "End"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Test Report");
                    // 第一条数据在 Row 2
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[2, 2].Text.ShouldBe("10");
                    // 最后一条数据
                    sheet.Cells[itemCount + 1, 1].Text.ShouldBe($"Product{itemCount}");
                    // Summary 在表格之后
                    sheet.Cells[itemCount + 2, 1].Text.ShouldBe("End");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 2. 多表格不同行 - Issue #617 核心场景

        [Theory(DisplayName = "#617 多表格不同行-各种数据量导出正确")]
        [InlineData(1, 1)]
        [InlineData(2, 1)]
        [InlineData(3, 1)]
        [InlineData(3, 2)]
        [InlineData(4, 1)]
        [InlineData(4, 3)]
        [InlineData(5, 1)]
        [InlineData(5, 3)]
        [InlineData(5, 5)]
        [InlineData(10, 5)]
        [InlineData(1, 10)]
        public async Task TwoTables_DifferentRows_VariousCounts_ShouldExportCorrectly(int orderCount, int logCount)
        {
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_Two_{orderCount}_{logCount}.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    ReportDate = "2024-06-12",
                    OrderItems = CreateOrderItems(orderCount),
                    OrderLogs = CreateOrderLogs(logCount),
                    Summary = "Done"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    // Row 1: 静态字段
                    sheet.Cells[1, 1].Text.ShouldBe("Report");
                    sheet.Cells[1, 3].Text.ShouldBe("2024-06-12");

                    // OrderItems 从 Row 2 开始
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[2, 2].Text.ShouldBe("10");
                    sheet.Cells[2, 3].Text.ShouldBe("99.5");
                    if (orderCount > 1)
                    {
                        sheet.Cells[orderCount + 1, 1].Text.ShouldBe($"Product{orderCount}");
                    }

                    // OrderLogs 紧跟 OrderItems 之后
                    var logStartRow = orderCount + 2;
                    sheet.Cells[logStartRow, 1].Text.ShouldBe("Action1");
                    sheet.Cells[logStartRow, 2].Text.ShouldBe("User1");
                    sheet.Cells[logStartRow, 3].Text.ShouldBe("2024-01-01");
                    if (logCount > 1)
                    {
                        sheet.Cells[logStartRow + logCount - 1, 1].Text.ShouldBe($"Action{logCount}");
                    }

                    // Summary 在所有表格之后
                    var summaryRow = logStartRow + logCount;
                    sheet.Cells[summaryRow, 1].Text.ShouldBe("Done");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 3. 多表格不同行 - 两个表格之间有 Cell Writer

        [Theory(DisplayName = "#617 表格间有CellWriter-各种数据量导出正确")]
        [InlineData(1, 1)]
        [InlineData(3, 1)]
        [InlineData(4, 2)]
        [InlineData(5, 3)]
        [InlineData(10, 5)]
        public async Task CellWriterBetweenTables_VariousCounts_ShouldExportCorrectly(int orderCount, int logCount)
        {
            var tplPath = CreateCellBetweenTablesTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_CellBetween_{orderCount}_{logCount}.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    ReportDate = "2024-06-12",
                    OrderItems = CreateOrderItems(orderCount),
                    OrderLogs = CreateOrderLogs(logCount),
                    Summary = "MiddleText"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    // Row 1: Title
                    sheet.Cells[1, 1].Text.ShouldBe("Report");

                    // OrderItems 从 Row 2 开始
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[orderCount + 1, 1].Text.ShouldBe($"Product{orderCount}");

                    // Summary (Cell Writer) 在 OrderItems 之后
                    var summaryRow = orderCount + 2;
                    sheet.Cells[summaryRow, 1].Text.ShouldBe("MiddleText");

                    // OrderLogs 在 Summary 之后
                    var logStartRow = summaryRow + 1;
                    sheet.Cells[logStartRow, 1].Text.ShouldBe("Action1");
                    if (logCount > 1)
                    {
                        sheet.Cells[logStartRow + logCount - 1, 1].Text.ShouldBe($"Action{logCount}");
                    }

                    // ReportDate 在 OrderLogs 之后
                    var reportDateRow = logStartRow + logCount;
                    sheet.Cells[reportDateRow, 1].Text.ShouldBe("2024-06-12");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 4. 多表格同行 (#296 场景)

        [Theory(DisplayName = "#296 多表格同行-各种数据量导出正确")]
        [InlineData(1, 1)]
        [InlineData(5, 5)]
        public async Task TwoTables_SameRow_VariousCounts_ShouldExportCorrectly(int orderCount, int logCount)
        {
            var tplPath = CreateSameRowTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_SameRow_{orderCount}_{logCount}.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    OrderItems = CreateOrderItems(orderCount),
                    OrderLogs = CreateOrderLogs(logCount),
                    Summary = "Done"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Report");

                    // 两个表格在同一行，以较大的为准插入行
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[2, 4].Text.ShouldBe("Action1");

                    // Summary 在 max(orderCount, logCount) 之后
                    var maxCount = Math.Max(orderCount, logCount);
                    sheet.Cells[maxCount + 2, 1].Text.ShouldBe("Done");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 5. Issue #617 精确复现场景

        [Fact(DisplayName = "#617 精确复现：DataList 4条 + FlowHistoryList 2条")]
        public async Task Issue617_ExactRepro_4Plus2_ShouldNotGarble()
        {
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_Repro_4_2.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "设备巡检报告",
                    ReportDate = "2024年06月12日",
                    OrderItems = CreateOrderItems(4),
                    OrderLogs = CreateOrderLogs(2),
                    Summary = "巡检完毕，一切正常"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("设备巡检报告");
                    sheet.Cells[1, 3].Text.ShouldBe("2024年06月12日");

                    // OrderItems: Row 2-5
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[3, 1].Text.ShouldBe("Product2");
                    sheet.Cells[4, 1].Text.ShouldBe("Product3");
                    sheet.Cells[5, 1].Text.ShouldBe("Product4");

                    // OrderLogs: Row 6-7
                    sheet.Cells[6, 1].Text.ShouldBe("Action1");
                    sheet.Cells[6, 2].Text.ShouldBe("User1");
                    sheet.Cells[7, 1].Text.ShouldBe("Action2");
                    sheet.Cells[7, 2].Text.ShouldBe("User2");

                    // Summary: Row 8
                    sheet.Cells[8, 1].Text.ShouldBe("巡检完毕，一切正常");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        [Fact(DisplayName = "#617 精确复现：DataList 5条 + FlowHistoryList 3条")]
        public async Task Issue617_ExactRepro_5Plus3_ShouldNotGarble()
        {
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_Repro_5_3.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    ReportDate = "2024-06-12",
                    OrderItems = CreateOrderItems(5),
                    OrderLogs = CreateOrderLogs(3),
                    Summary = "End"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Report");
                    // OrderItems: Row 2-6
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[6, 1].Text.ShouldBe("Product5");
                    // OrderLogs: Row 7-9
                    sheet.Cells[7, 1].Text.ShouldBe("Action1");
                    sheet.Cells[9, 1].Text.ShouldBe("Action3");
                    // Summary: Row 10
                    sheet.Cells[10, 1].Text.ShouldBe("End");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        [Fact(DisplayName = "#617 精确复现：DataList 10条 + FlowHistoryList 5条")]
        public async Task Issue617_ExactRepro_10Plus5_ShouldNotGarble()
        {
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_Repro_10_5.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Big Report",
                    ReportDate = "2024-06-12",
                    OrderItems = CreateOrderItems(10),
                    OrderLogs = CreateOrderLogs(5),
                    Summary = "Finished"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Big Report");
                    // OrderItems: Row 2-11
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[11, 1].Text.ShouldBe("Product10");
                    // OrderLogs: Row 12-16
                    sheet.Cells[12, 1].Text.ShouldBe("Action1");
                    sheet.Cells[16, 1].Text.ShouldBe("Action5");
                    // Summary: Row 17
                    sheet.Cells[17, 1].Text.ShouldBe("Finished");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 6. 边界场景

        [Fact(DisplayName = "单表格-1条数据不需要插入行")]
        public async Task SingleTable_SingleItem_NoInsertRow()
        {
            var tplPath = CreateSingleTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_Single1.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "T",
                    OrderItems = CreateOrderItems(1),
                    Summary = "S"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("T");
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[2, 2].Text.ShouldBe("10");
                    sheet.Cells[2, 3].Text.ShouldBe("99.5");
                    sheet.Cells[3, 1].Text.ShouldBe("S");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        [Fact(DisplayName = "多表格-第一个有数据第二个空")]
        public async Task TwoTables_FirstHasDataSecondEmpty()
        {
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_FirstHasSecondEmpty.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    ReportDate = "2024-06-12",
                    OrderItems = CreateOrderItems(5),
                    OrderLogs = new List<OrderLog>(),
                    Summary = "End"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Report");
                    // OrderItems: Row 2-6
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[6, 1].Text.ShouldBe("Product5");
                    // OrderLogs empty, Summary should be in the right place
                    // Note: when second table is empty, the empty row remains but Summary should follow
                    var summaryRow = 8; // 2(OrderItems end) + 1(OrderLogs empty row) + 1(Summary)
                    sheet.Cells[summaryRow, 1].Text.ShouldBe("End");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        [Fact(DisplayName = "多表格同行-第一个多第二个多应正确")]
        public async Task SameRow_FirstLargerThanSecond_ShouldExportCorrectly()
        {
            var tplPath = CreateSameRowTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_SameRow_LargeSmall.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    OrderItems = CreateOrderItems(5),
                    OrderLogs = CreateOrderLogs(2),
                    Summary = "End"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Report");
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[2, 4].Text.ShouldBe("Action1");
                    sheet.Cells[3, 4].Text.ShouldBe("Action2");
                    // Summary after max(5,2)=5 rows
                    sheet.Cells[7, 1].Text.ShouldBe("End");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        [Fact(DisplayName = "多表格同行-第二个多第一个少应正确")]
        public async Task SameRow_SecondLargerThanFirst_ShouldExportCorrectly()
        {
            var tplPath = CreateSameRowTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_SameRow_SmallLarge.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    OrderItems = CreateOrderItems(2),
                    OrderLogs = CreateOrderLogs(5),
                    Summary = "End"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Report");
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[2, 4].Text.ShouldBe("Action1");
                    sheet.Cells[6, 4].Text.ShouldBe("Action5");
                    // Summary after max(2,5)=5 rows
                    sheet.Cells[7, 1].Text.ShouldBe("End");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 7. 三个表格不同行

        /// <summary>
        /// 创建三个表格不同行的模板
        ///   Row 1: A1={{Title}}
        ///   Row 2: A2={{Table>>OrderItems|ProductName}} B2={{Quantity}} C2={{Price|>>Table}}
        ///   Row 3: A3={{Table>>OrderLogs|Action}} B3={{Operator}} C3={{Time|>>Table}}
        ///   Row 4: A4={{Table>>OrderItems|ProductName}} B4={{Quantity}} C4={{Price|>>Table}}
        ///   Row 5: A5={{Summary}}
        /// 注意：复用 OrderItems 作为第三个表格（不同 key 名需不同的 List 属性，
        ///       这里用同一个 List 测试重复 key 的场景，实际使用中可用不同属性名）
        /// </summary>
        private string CreateThreeTableTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_ThreeTbl_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = "{{Title}}";
                sheet.Cells[2, 1].Value = "{{Table>>OrderItems|ProductName}}";
                sheet.Cells[2, 2].Value = "{{Quantity}}";
                sheet.Cells[2, 3].Value = "{{Price|>>Table}}";
                sheet.Cells[3, 1].Value = "{{Table>>OrderLogs|Action}}";
                sheet.Cells[3, 2].Value = "{{Operator}}";
                sheet.Cells[3, 3].Value = "{{Time|>>Table}}";
                sheet.Cells[4, 1].Value = "{{ReportDate}}";
                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        [Fact(DisplayName = "#617 三个表格不同行-中间有CellWriter")]
        public async Task TwoTablesPlusCellWriter_DifferentRowCounts_ShouldExportCorrectly()
        {
            // 模板: Title → OrderItems(4) → OrderLogs(2) → ReportDate
            // 这个场景中 ReportDate 是 CellWriter，在第二个表格之后
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_ThreeLayers.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    ReportDate = "2024-12-25",
                    OrderItems = CreateOrderItems(4),
                    OrderLogs = CreateOrderLogs(2),
                    Summary = "End"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    // Title + ReportDate 在 Row 1
                    sheet.Cells[1, 1].Text.ShouldBe("Report");
                    sheet.Cells[1, 3].Text.ShouldBe("2024-12-25");
                    // OrderItems: Row 2-5
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[5, 1].Text.ShouldBe("Product4");
                    sheet.Cells[2, 3].Text.ShouldBe("99.5"); // decimal 格式验证
                    // OrderLogs: Row 6-7
                    sheet.Cells[6, 1].Text.ShouldBe("Action1");
                    sheet.Cells[7, 1].Text.ShouldBe("Action2");
                    // Summary: Row 8
                    sheet.Cells[8, 1].Text.ShouldBe("End");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 8. CellWriter 在表格之前

        /// <summary>
        /// 创建 CellWriter 在表格之前 + 表格之间 + 表格之后的模板
        ///   Row 1: A1={{ReportDate}}  ← CellWriter，在所有表格之前
        ///   Row 2: A2={{Table>>OrderItems|ProductName}} B2={{Quantity}} C2={{Price|>>Table}}
        ///   Row 3: A3={{Title}}  ← CellWriter，在两个表格之间
        ///   Row 4: A4={{Table>>OrderLogs|Action}} B4={{Operator}} C4={{Time|>>Table}}
        ///   Row 5: A5={{Summary}}  ← CellWriter，在所有表格之后
        /// </summary>
        private string CreateCellWriterSurroundTablesTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_Surround_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = "{{ReportDate}}";
                sheet.Cells[2, 1].Value = "{{Table>>OrderItems|ProductName}}";
                sheet.Cells[2, 2].Value = "{{Quantity}}";
                sheet.Cells[2, 3].Value = "{{Price|>>Table}}";
                sheet.Cells[3, 1].Value = "{{Title}}";
                sheet.Cells[4, 1].Value = "{{Table>>OrderLogs|Action}}";
                sheet.Cells[4, 2].Value = "{{Operator}}";
                sheet.Cells[4, 3].Value = "{{Time|>>Table}}";
                sheet.Cells[5, 1].Value = "{{Summary}}";
                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        [Theory(DisplayName = "#617 CellWriter包围表格-各种数据量导出正确")]
        [InlineData(1, 1)]
        [InlineData(3, 2)]
        [InlineData(5, 5)]
        [InlineData(10, 3)]
        public async Task CellWriterSurroundTables_VariousCounts_ShouldExportCorrectly(int orderCount, int logCount)
        {
            var tplPath = CreateCellWriterSurroundTablesTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_Surround_{orderCount}_{logCount}.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "中间标题",
                    ReportDate = "2024-01-01",
                    OrderItems = CreateOrderItems(orderCount),
                    OrderLogs = CreateOrderLogs(logCount),
                    Summary = "结尾"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    // Row 1: ReportDate（在所有表格之前，不应被偏移）
                    sheet.Cells[1, 1].Text.ShouldBe("2024-01-01");
                    // OrderItems: Row 2 开始
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[orderCount + 1, 1].Text.ShouldBe($"Product{orderCount}");
                    // Title（CellWriter 在两个表格之间）
                    var titleRow = orderCount + 2;
                    sheet.Cells[titleRow, 1].Text.ShouldBe("中间标题");
                    // OrderLogs
                    var logStartRow = titleRow + 1;
                    sheet.Cells[logStartRow, 1].Text.ShouldBe("Action1");
                    if (logCount > 1)
                    {
                        sheet.Cells[logStartRow + logCount - 1, 1].Text.ShouldBe($"Action{logCount}");
                    }
                    // Summary（在所有表格之后）
                    var summaryRow = logStartRow + logCount;
                    sheet.Cells[summaryRow, 1].Text.ShouldBe("结尾");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 9. 空列表场景

        [Fact(DisplayName = "#617 两个表格都为空-不应崩溃")]
        public async Task TwoTables_BothEmpty_ShouldNotCrash()
        {
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_BothEmpty.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Empty Report",
                    ReportDate = "2024-06-12",
                    OrderItems = new List<OrderItem>(),
                    OrderLogs = new List<OrderLog>(),
                    Summary = "Nothing"
                };
                // 不应抛异常
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Empty Report");
                    sheet.Cells[1, 3].Text.ShouldBe("2024-06-12");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        [Fact(DisplayName = "#617 第一个空第二个有数据")]
        public async Task TwoTables_FirstEmptySecondHasData_ShouldExportCorrectly()
        {
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_FirstEmpty.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Report",
                    ReportDate = "2024-06-12",
                    OrderItems = new List<OrderItem>(),
                    OrderLogs = CreateOrderLogs(3),
                    Summary = "End"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Report");
                    // OrderLogs 应该正确渲染（第一个表格为空时不影响第二个）
                    sheet.Cells[3, 1].Text.ShouldBe("Action1");
                    sheet.Cells[5, 1].Text.ShouldBe("Action3");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 10. 大数据量

        [Fact(DisplayName = "#617 大数据量(50+20)导出正确")]
        public async Task TwoTables_LargeDataCount_ShouldExportCorrectly()
        {
            var tplPath = CreateTwoTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_Large.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Large Report",
                    ReportDate = "2024-06-12",
                    OrderItems = CreateOrderItems(50),
                    OrderLogs = CreateOrderLogs(20),
                    Summary = "Done"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Large Report");
                    // OrderItems: Row 2-51
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[2, 3].Text.ShouldBe("99.5"); // decimal 精度
                    sheet.Cells[51, 1].Text.ShouldBe("Product50");
                    // OrderLogs: Row 52-71
                    sheet.Cells[52, 1].Text.ShouldBe("Action1");
                    sheet.Cells[71, 1].Text.ShouldBe("Action20");
                    // Summary: Row 72
                    sheet.Cells[72, 1].Text.ShouldBe("Done");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 11. 数据类型和格式验证

        [Fact(DisplayName = "#617 decimal/int 数据类型正确渲染")]
        public async Task DataTypes_ShouldRenderCorrectly()
        {
            var tplPath = CreateSingleTableTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_DataTypes.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiTableDto
                {
                    Title = "Types",
                    OrderItems = new List<OrderItem>
                    {
                        new OrderItem { ProductName = "Widget", Quantity = 999, Price = 1234.56m },
                        new OrderItem { ProductName = "Gadget", Quantity = 0, Price = 0.01m },
                    },
                    Summary = "End"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[2, 1].Text.ShouldBe("Widget");
                    sheet.Cells[2, 2].Text.ShouldBe("999");
                    sheet.Cells[2, 3].Text.ShouldBe("1234.56");
                    sheet.Cells[3, 1].Text.ShouldBe("Gadget");
                    sheet.Cells[3, 2].Text.ShouldBe("0");
                    sheet.Cells[3, 3].Text.ShouldBe("0.01");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 12. 三个及以上表格

        public class ThreeTableItem
        {
            public string Name { get; set; }
            public int Value { get; set; }
        }

        public class ThreeTableDto
        {
            public string Title { get; set; }
            public List<OrderItem> TableA { get; set; }
            public List<OrderLog> TableB { get; set; }
            public List<ThreeTableItem> TableC { get; set; }
            public string Footer { get; set; }
        }

        private string CreateThreeTableMultiTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_3Tbl_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = "{{Title}}";
                sheet.Cells[2, 1].Value = "{{Table>>TableA|ProductName}}";
                sheet.Cells[2, 2].Value = "{{Quantity}}";
                sheet.Cells[2, 3].Value = "{{Price|>>Table}}";
                sheet.Cells[3, 1].Value = "{{Table>>TableB|Action}}";
                sheet.Cells[3, 2].Value = "{{Operator}}";
                sheet.Cells[3, 3].Value = "{{Time|>>Table}}";
                sheet.Cells[4, 1].Value = "{{Table>>TableC|Name}}";
                sheet.Cells[4, 2].Value = "{{Value|>>Table}}";
                sheet.Cells[5, 1].Value = "{{Footer}}";
                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        private static List<ThreeTableItem> CreateThreeTableItems(int count)
        {
            var items = new List<ThreeTableItem>();
            for (int i = 1; i <= count; i++)
                items.Add(new ThreeTableItem { Name = $"Item{i}", Value = i * 100 });
            return items;
        }

        [Theory(DisplayName = "#617 三个表格不同行-各种数据量导出正确")]
        [InlineData(1, 1, 1)]
        [InlineData(2, 3, 1)]
        [InlineData(5, 2, 4)]
        [InlineData(3, 5, 2)]
        public async Task ThreeTables_DifferentRows_VariousCounts_ShouldExportCorrectly(int a, int b, int c)
        {
            var tplPath = CreateThreeTableMultiTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_3Tbl_{a}_{b}_{c}.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new ThreeTableDto
                {
                    Title = "3-Table Report",
                    TableA = CreateOrderItems(a),
                    TableB = CreateOrderLogs(b),
                    TableC = CreateThreeTableItems(c),
                    Footer = "EOF"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("3-Table Report");

                    // TableA starts Row 2
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    if (a > 1) sheet.Cells[a + 1, 1].Text.ShouldBe($"Product{a}");

                    // TableB after TableA
                    var bStart = a + 2;
                    sheet.Cells[bStart, 1].Text.ShouldBe("Action1");
                    if (b > 1) sheet.Cells[bStart + b - 1, 1].Text.ShouldBe($"Action{b}");

                    // TableC after TableB
                    var cStart = bStart + b;
                    sheet.Cells[cStart, 1].Text.ShouldBe("Item1");
                    sheet.Cells[cStart, 2].Text.ShouldBe("100");
                    if (c > 1) sheet.Cells[cStart + c - 1, 1].Text.ShouldBe($"Item{c}");

                    // Footer
                    sheet.Cells[cStart + c, 1].Text.ShouldBe("EOF");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        [Fact(DisplayName = "#617 三个表格-中间为空不崩溃")]
        public async Task ThreeTables_MiddleEmpty_ShouldNotCrash()
        {
            var tplPath = CreateThreeTableMultiTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_3Tbl_MidEmpty.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new ThreeTableDto
                {
                    Title = "Middle Empty",
                    TableA = CreateOrderItems(3),
                    TableB = new List<OrderLog>(),
                    TableC = CreateThreeTableItems(2),
                    Footer = "EOF"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Middle Empty");
                    // TableA: Row 2-4
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[4, 1].Text.ShouldBe("Product3");
                    // TableB empty → TableC follows at Row 6
                    sheet.Cells[6, 1].Text.ShouldBe("Item1");
                    sheet.Cells[7, 1].Text.ShouldBe("Item2");
                    sheet.Cells[8, 1].Text.ShouldBe("EOF");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 13. 不同列数表格混合

        public class MixedColDto
        {
            public string Title { get; set; }
            public List<OrderItem> WideTable { get; set; }   // 3 columns
            public List<ThreeTableItem> NarrowTable { get; set; } // 2 columns
            public string Summary { get; set; }
        }

        private string CreateMixedColTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_MixCol_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells[1, 1].Value = "{{Title}}";
                // 3-column table
                sheet.Cells[2, 1].Value = "{{Table>>WideTable|ProductName}}";
                sheet.Cells[2, 2].Value = "{{Quantity}}";
                sheet.Cells[2, 3].Value = "{{Price|>>Table}}";
                // 2-column table starting at col D (offset)
                sheet.Cells[3, 1].Value = "{{Table>>NarrowTable|Name}}";
                sheet.Cells[3, 2].Value = "{{Value|>>Table}}";
                sheet.Cells[4, 1].Value = "{{Summary}}";
                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        [Theory(DisplayName = "#617 不同列数表格混合-导出正确")]
        [InlineData(3, 2)]
        [InlineData(5, 5)]
        [InlineData(1, 5)]
        [InlineData(5, 1)]
        public async Task MixedColumnCounts_DifferentSizes_ShouldExportCorrectly(int wideCount, int narrowCount)
        {
            var tplPath = CreateMixedColTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_MixCol_{wideCount}_{narrowCount}.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MixedColDto
                {
                    Title = "Mixed Columns",
                    WideTable = CreateOrderItems(wideCount),
                    NarrowTable = CreateThreeTableItems(narrowCount),
                    Summary = "Done"
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = pck.Workbook.Worksheets.First();
                    sheet.Cells[1, 1].Text.ShouldBe("Mixed Columns");

                    // WideTable (3 cols): Row 2+
                    sheet.Cells[2, 1].Text.ShouldBe("Product1");
                    sheet.Cells[2, 2].Text.ShouldBe("10");
                    sheet.Cells[2, 3].Text.ShouldBe("99.5");
                    if (wideCount > 1)
                        sheet.Cells[wideCount + 1, 3].Text.ShouldBe($"{wideCount * 99.5m}");

                    // NarrowTable (2 cols): after WideTable
                    var nStart = wideCount + 2;
                    sheet.Cells[nStart, 1].Text.ShouldBe("Item1");
                    sheet.Cells[nStart, 2].Text.ShouldBe("100");
                    if (narrowCount > 1)
                        sheet.Cells[nStart + narrowCount - 1, 1].Text.ShouldBe($"Item{narrowCount}");

                    // Summary
                    sheet.Cells[nStart + narrowCount, 1].Text.ShouldBe("Done");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion

        #region 14. 多 Sheet 表格导出

        public class MultiSheetDto
        {
            public string Title { get; set; }
            public List<OrderItem> Sheet1Items { get; set; }
            public List<OrderLog> Sheet2Items { get; set; }
        }

        private string CreateMultiSheetTemplate()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_MultiSht_{Guid.NewGuid():N}.xlsx");
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Orders");
                sheet1.Cells[1, 1].Value = "{{Title}}";
                sheet1.Cells[2, 1].Value = "{{Table>>Sheet1Items|ProductName}}";
                sheet1.Cells[2, 2].Value = "{{Quantity}}";
                sheet1.Cells[2, 3].Value = "{{Price|>>Table}}";

                var sheet2 = package.Workbook.Worksheets.Add("Logs");
                sheet2.Cells[1, 1].Value = "{{Table>>Sheet2Items|Action}}";
                sheet2.Cells[1, 2].Value = "{{Operator}}";
                sheet2.Cells[1, 3].Value = "{{Time|>>Table}}";

                package.SaveAs(new FileInfo(path));
            }
            return path;
        }

        [Theory(DisplayName = "#617 多Sheet表格-各Sheet独立导出正确")]
        [InlineData(3, 2)]
        [InlineData(5, 5)]
        [InlineData(1, 5)]
        public async Task MultiSheet_DifferentData_ShouldExportCorrectly(int sheet1Count, int sheet2Count)
        {
            var tplPath = CreateMultiSheetTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Issue617_MultiSht_{sheet1Count}_{sheet2Count}.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiSheetDto
                {
                    Title = "Multi-Sheet",
                    Sheet1Items = CreateOrderItems(sheet1Count),
                    Sheet2Items = CreateOrderLogs(sheet2Count)
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Sheet 1
                    var s1 = pck.Workbook.Worksheets["Orders"];
                    s1.ShouldNotBeNull();
                    s1.Cells[1, 1].Text.ShouldBe("Multi-Sheet");
                    s1.Cells[2, 1].Text.ShouldBe("Product1");
                    if (sheet1Count > 1)
                        s1.Cells[sheet1Count + 1, 1].Text.ShouldBe($"Product{sheet1Count}");

                    // Sheet 2
                    var s2 = pck.Workbook.Worksheets["Logs"];
                    s2.ShouldNotBeNull();
                    s2.Cells[1, 1].Text.ShouldBe("Action1");
                    if (sheet2Count > 1)
                        s2.Cells[sheet2Count, 1].Text.ShouldBe($"Action{sheet2Count}");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        [Fact(DisplayName = "#617 多Sheet-一个Sheet为空不崩溃")]
        public async Task MultiSheet_OneSheetEmpty_ShouldNotCrash()
        {
            var tplPath = CreateMultiSheetTemplate();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Issue617_MultiSht_Empty.xlsx");
            try
            {
                IExportFileByTemplate exporter = new ExcelExporter();
                var data = new MultiSheetDto
                {
                    Title = "One Empty",
                    Sheet1Items = new List<OrderItem>(),
                    Sheet2Items = CreateOrderLogs(3)
                };
                await exporter.ExportByTemplate(filePath, data, tplPath);

                using (var pck = new ExcelPackage(new FileInfo(filePath)))
                {
                    var s1 = pck.Workbook.Worksheets["Orders"];
                    s1.Cells[1, 1].Text.ShouldBe("One Empty");

                    var s2 = pck.Workbook.Worksheets["Logs"];
                    s2.Cells[1, 1].Text.ShouldBe("Action1");
                    s2.Cells[3, 1].Text.ShouldBe("Action3");
                }
            }
            finally
            {
                CleanupFile(tplPath);
                CleanupFile(filePath);
            }
        }

        #endregion
    }
}
