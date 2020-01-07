// ======================================================================
// 
//           filename : ExcelExporter.cs
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
using System.Linq;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel.Utility;
using Magicodes.ExporterAndImporter.Excel.Utility.TemplateExport;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     Excel导出程序
    /// </summary>
    public class ExcelExporter : IExporter, IExportFileByTemplate
    {
        /// <summary>
        ///     表头处理函数
        /// </summary>
        public static Func<string, string> ColumnHeaderStringFunc { get; set; } = str => str;

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="dataItems">数据列</param>
        /// <returns>文件</returns>
        public Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("文件名必须填写!", nameof(fileName));
            var fileInfo = ExcelHelper.CreateExcelPackage(fileName, excelPackage =>
            {
                //导出定义
                var exporter = GetExporterAttribute<T>();

                if (exporter?.Author != null)
                    excelPackage.Workbook.Properties.Author = exporter?.Author;

                var sheet = excelPackage.Workbook.Worksheets.Add(exporter?.Name ?? "导出结果");
                sheet.OutLineApplyStyle = true;
                if (GetExporterHeaderInfoList<T>(out var exporterHeaderList)) return;
                AddHeader(exporterHeaderList, sheet, exporter);
                AddDataItems(sheet, exporterHeaderList, dataItems, exporter);
                AddStyle(exporter, exporterHeaderList, sheet);
            });
            return Task.FromResult(fileInfo);
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <returns>文件二进制数组</returns>
        public Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems) where T : class
        {
            using (var excelPackage = new ExcelPackage())
            {
                //导出定义
                var exporter = GetExporterAttribute<T>();

                if (exporter?.Author != null)
                    excelPackage.Workbook.Properties.Author = exporter?.Author;

                var sheet = excelPackage.Workbook.Worksheets.Add(exporter?.Name ?? "导出结果");
                sheet.OutLineApplyStyle = true;
                if (GetExporterHeaderInfoList<T>(out var exporterHeaderList)) return null;
                AddHeader(exporterHeaderList, sheet, exporter);
                AddDataItems(sheet, exporterHeaderList, dataItems, exporter);
                AddStyle(exporter, exporterHeaderList, sheet);
                return Task.FromResult(excelPackage.GetAsByteArray());
            }
        }

        public Task<ExportFileInfo> Export<T>(string fileName, DataTable dataItems) where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("文件名必须填写!", nameof(fileName));
            var fileInfo = ExcelHelper.CreateExcelPackage(fileName, excelPackage =>
            {
                //导出定义
                var exporter = GetExporterAttribute<T>();

                if (exporter?.Author != null)
                    excelPackage.Workbook.Properties.Author = exporter?.Author;

                if (GetExporterHeaderInfoList<T>(out var exporterHeaderList, dataItems.Columns))
                    return;

                var data = dataItems.SplitDataTable();

                var count = 0;
                foreach (DataTable table in data.Tables)
                {
                    var sheet = GetWorksheet(excelPackage, exporter, count);
                    AddWorksheet<T>(table, exporter, exporterHeaderList, sheet);
                    count++;
                }
            });
            return Task.FromResult(fileInfo);
        }

        /// <summary>
        /// 导出字节
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public Task<byte[]> ExportAsByteArray<T>(DataTable dataItems) where T : class
        {
            using (var excelPackage = new ExcelPackage())
            {
                //导出定义
                var exporter = GetExporterAttribute<T>();

                if (exporter?.Author != null)
                    excelPackage.Workbook.Properties.Author = exporter?.Author;

                if (GetExporterHeaderInfoList<T>(out var exporterHeaderList, dataItems.Columns))
                    return null;

                var data = dataItems.SplitDataTable();
                var count = 0;
                foreach (DataTable table in data.Tables)
                {
                    var sheet = GetWorksheet(excelPackage, exporter, count);
                    AddWorksheet<T>(table, exporter, exporterHeaderList, sheet);
                    count++;
                }

                return Task.FromResult(excelPackage.GetAsByteArray());
            }
        }

        /// <summary>
        ///     导出excel表头
        /// </summary>
        /// <param name="items">表头数组</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <param name="globalStyle">全局样式</param>
        /// <param name="styles">样式</param>
        /// <returns></returns>
        public Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName, ExcelHeadStyle globalStyle = null,
            List<ExcelHeadStyle> styles = null)
        {
            using (var excelPackage = new ExcelPackage())
            {
                var sheet = excelPackage.Workbook.Worksheets.Add(sheetName ?? "导出结果");
                sheet.OutLineApplyStyle = true;
                AddHeader(items, sheet);
                AddStyle(sheet, items.Length, globalStyle, styles);
                return Task.FromResult(excelPackage.GetAsByteArray());
            }
        }

        /// <summary>
        ///     导出Excel表头
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>文件二进制数组</returns>
        public Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class
        {
            using (var excelPackage = new ExcelPackage())
            {
                //导出定义
                var exporter = GetExporterAttribute<T>();

                if (exporter?.Author != null)
                    excelPackage.Workbook.Properties.Author = exporter?.Author;

                var sheet = excelPackage.Workbook.Worksheets.Add(exporter?.Name ?? "导出结果");
                sheet.OutLineApplyStyle = true;
                if (GetExporterHeaderInfoList<T>(out var exporterHeaderList)) return null;
                AddHeader(exporterHeaderList, sheet, exporter);
                AddStyle(exporter, exporterHeaderList, sheet);
                return Task.FromResult(excelPackage.GetAsByteArray());
            }

            throw new NotImplementedException();
        }

        private void AddWorksheet<T>(DataTable dataItems, ExcelExporterAttribute exporter,
            List<ExporterHeaderInfo> exporterHeaderList, ExcelWorksheet sheet) where T : class
        {
            AddHeader(exporterHeaderList, sheet, exporter);
            AddDataItems<T>(sheet, dataItems, exporter);
            AddStyle(exporter, exporterHeaderList, sheet);
        }


        private static ExcelWorksheet GetWorksheet(ExcelPackage excelPackage, ExcelExporterAttribute exporter,
            int count = 0)
        {
            var name = exporter?.Name ?? "导出结果";
            var sheet = excelPackage.Workbook.Worksheets.Add($"{name}-{count}");
            sheet.OutLineApplyStyle = true;
            return sheet;
        }

        /// <summary>
        ///     创建表头
        /// </summary>
        /// <param name="exporterHeaderDtoList"></param>
        /// <param name="sheet"></param>
        /// <param name="exporter"></param>
        protected void AddHeader(List<ExporterHeaderInfo> exporterHeaderDtoList, ExcelWorksheet sheet,
            ExcelExporterAttribute exporter)
        {
            foreach (var exporterHeaderDto in exporterHeaderDtoList)
                if (exporterHeaderDto != null)
                {
                    if (exporterHeaderDto.ExporterHeader != null)
                    {
                        var exporterHeaderAttribute = exporterHeaderDto.ExporterHeader;
                        if (exporterHeaderAttribute != null && !exporterHeaderAttribute.IsIgnore)
                        {
                            var name = exporterHeaderAttribute.DisplayName.IsNullOrWhiteSpace()
                                ? exporterHeaderDto.PropertyName
                                : exporterHeaderAttribute.DisplayName;

                            sheet.Cells[1, exporterHeaderDto.Index].Value = ColumnHeaderStringFunc(name);
                            sheet.Cells[1, exporterHeaderDto.Index].Style.Font.Bold = exporterHeaderAttribute.IsBold;

                            var size = exporter?.HeaderFontSize ?? exporterHeaderAttribute.FontSize;
                            if (size.HasValue)
                                sheet.Cells[1, exporterHeaderDto.Index].Style.Font.Size = size.Value;
                        }
                    }
                    else
                    {
                        sheet.Cells[1, exporterHeaderDto.Index].Value =
                            ColumnHeaderStringFunc(exporterHeaderDto.PropertyName);
                    }
                }
        }


        /// <summary>
        ///     创建表头
        /// </summary>
        /// <param name="exporterHeaders">表头数组</param>
        /// <param name="sheet">工作簿</param>
        protected void AddHeader(string[] exporterHeaders, ExcelWorksheet sheet)
        {
            var columnIndex = 0;
            foreach (var exporterHeader in exporterHeaders)
                if (exporterHeader != null)
                {
                    columnIndex++;
                    sheet.Cells[1, columnIndex].Value = exporterHeader;
                }
        }

        /// <summary>
        ///     添加导出数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="exporterHeaders"></param>
        /// <param name="items"></param>
        /// <param name="exporter"></param>
        protected void AddDataItems<T>(ExcelWorksheet sheet, List<ExporterHeaderInfo> exporterHeaders,
            ICollection<T> items,
            ExcelExporterAttribute exporter)
        {
            if (items == null || items.Count == 0)
                return;
            var tbStyle = TableStyles.Medium10;
            if (exporter != null && !exporter.TableStyle.IsNullOrWhiteSpace())
                tbStyle = (TableStyles)Enum.Parse(typeof(TableStyles), exporter.TableStyle);
            sheet.Cells["A2"].LoadFromCollection(items, false, tbStyle);
        }

        /// <summary>
        ///     添加导出数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="items"></param>
        /// <param name="exporter"></param>
        protected void AddDataItems<T>(ExcelWorksheet sheet, DataTable items, ExcelExporterAttribute exporter)
        {
            if (items == null || items.Rows.Count == 0)
                return;
            var tbStyle = TableStyles.Medium10;
            if (exporter != null && !exporter.TableStyle.IsNullOrWhiteSpace())
                tbStyle = (TableStyles)Enum.Parse(typeof(TableStyles), exporter.TableStyle);
            sheet.Cells["A2"].LoadFromDataTable(items, false, tbStyle);
        }


        /// <summary>
        ///     添加样式
        /// </summary>
        /// <param name="exporterHeaders"></param>
        /// <param name="sheet"></param>
        protected void AddStyle(ExcelExporterAttribute exporter, List<ExporterHeaderInfo> exporterHeaders,
            ExcelWorksheet sheet)
        {
            foreach (var exporterHeader in exporterHeaders)
                if (exporterHeader.ExporterHeader != null)
                {
                    if (exporterHeader.ExporterHeader.IsIgnore)
                    {
                        //TODO:后续直接修改数据导出逻辑（不写忽略列数据）
                        sheet.DeleteColumn(exporterHeader.Index);
                        //删除之后，序号依次-1
                        foreach (var item in exporterHeaders.Where(p => p.Index > exporterHeader.Index)) item.Index--;
                        continue;
                    }

                    var col = sheet.Column(exporterHeader.Index);
                    col.Style.Numberformat.Format = exporterHeader.ExporterHeader.Format;

                    if (exporter.AutoFitAllColumn || exporterHeader.ExporterHeader.IsAutoFit)
                        col.AutoFit();
                }
        }

        /// <summary>
        ///     添加样式
        /// </summary>
        /// <param name="sheet">excel工作簿</param>
        /// <param name="columns">总列数</param>
        /// <param name="globalStyle">全局样式</param>
        /// <param name="styles">样式</param>
        protected void AddStyle(ExcelWorksheet sheet, int columns, ExcelHeadStyle globalStyle = null,
            List<ExcelHeadStyle> styles = null)
        {
            var col = 0;
            if (styles != null)
            {
                foreach (var style in styles)
                {
                    col++;
                    if (col <= columns)
                    {
                        if (style.IsIgnore)
                        {
                            sheet.DeleteColumn(col);
                            continue;
                        }

                        var excelCol = sheet.Column(col);
                        if (!style.Format.IsNullOrWhiteSpace()) excelCol.Style.Numberformat.Format = style.Format;
                        excelCol.Style.Font.Bold = style.IsBold;
                        excelCol.Style.Font.Size = style.FontSize;

                        if (style.IsAutoFit) excelCol.AutoFit();
                    }
                }
            }
            else
            {
                if (globalStyle != null)
                    for (var i = 1; i <= columns; i++)
                    {
                        if (globalStyle.IsIgnore)
                        {
                            sheet.DeleteColumn(i);
                            continue;
                        }

                        var excelCol = sheet.Column(i);
                        if (!globalStyle.Format.IsNullOrWhiteSpace())
                            excelCol.Style.Numberformat.Format = globalStyle.Format;
                        excelCol.Style.Font.Bold = globalStyle.IsBold;
                        excelCol.Style.Font.Size = globalStyle.FontSize;

                        if (globalStyle.IsAutoFit) excelCol.AutoFit();
                    }
            }
        }

        /// <summary>
        ///     获取头部定义
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="exporterHeaderList"></param>
        /// <returns></returns>
        private static bool GetExporterHeaderInfoList<T>(out List<ExporterHeaderInfo> exporterHeaderList)
        {
            exporterHeaderList = new List<ExporterHeaderInfo>();
            var objProperties = typeof(T).GetProperties();
            if (objProperties == null || objProperties.Length == 0)
                return true;
            for (var i = 0; i < objProperties.Length; i++)
                exporterHeaderList.Add(new ExporterHeaderInfo
                {
                    Index = i + 1,
                    PropertyName = objProperties[i].Name,
                    ExporterHeader =
                        (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as
                            ExporterHeaderAttribute[])?.FirstOrDefault()
                });
            return false;
        }

        /// <summary>
        ///     获取头部定义
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="exporterHeaderList"></param>
        /// <param name="dataColumns"></param>
        /// <returns></returns>
        private static bool GetExporterHeaderInfoList<T>(out List<ExporterHeaderInfo> exporterHeaderList,
            DataColumnCollection dataColumns)
        {
            exporterHeaderList = new List<ExporterHeaderInfo>();
            var objProperties = typeof(T).GetProperties();
            if (objProperties == null || objProperties.Length == 0)
                return true;

            var index = 0;
            for (var k = 0; k < dataColumns.Count; k++)
                for (var i = 0; i < objProperties.Length; i++)
                    if (dataColumns[k].ColumnName.Equals(objProperties[i].Name))
                    {
                        index += 1;
                        exporterHeaderList.Add(new ExporterHeaderInfo
                        {
                            Index = index,
                            PropertyName = objProperties[i].Name,
                            ExporterHeader =
                                (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as
                                    ExporterHeaderAttribute[])?.FirstOrDefault()
                        });
                    }

            return false;
        }

        /// <summary>
        ///     获取表全局定义
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private static ExcelExporterAttribute GetExporterAttribute<T>() where T : class
        {
            var exporterTableAttributes =
                typeof(T).GetCustomAttributes(typeof(ExcelExporterAttribute), true) as ExcelExporterAttribute[];
            if (exporterTableAttributes != null && exporterTableAttributes.Length > 0)
                return exporterTableAttributes[0];

            var exporterAttributes =
                typeof(T).GetCustomAttributes(typeof(ExporterAttribute), true) as ExporterAttribute[];

            if (exporterAttributes != null && exporterAttributes.Length > 0)
            {
                var export = exporterAttributes[0];
                return new ExcelExporterAttribute
                {
                    FontSize = export.FontSize,
                    HeaderFontSize = export.HeaderFontSize
                };
            }

            return null;
        }


        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        /// <param name="template">HTML模板或模板路径</param>
        /// <returns></returns>
        public Task<ExportFileInfo> ExportByTemplate<T>(string fileName, T data, string template) where T : class
        {
            using (var helper = new TemplateExportHelper<T>())
            {
                helper.Export(fileName, template, data);
                return Task.FromResult(new ExportFileInfo());
            }
        }
    }
}