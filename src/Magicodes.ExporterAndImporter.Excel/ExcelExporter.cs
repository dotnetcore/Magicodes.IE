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
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
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
        ///     导出Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="dataItems">数据列</param>
        /// <returns>文件</returns>
        public async Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class
        {
            CheckFileName(fileName);

            var bytes = await ExportAsByteArray(dataItems);
            File.WriteAllBytes(fileName, bytes);

            var file = new ExportFileInfo(fileName,
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            return file;
        }

        private static void CheckFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("文件名必须填写!", nameof(fileName));
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <returns>文件二进制数组</returns>
        public Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems) where T : class
        {
            var helper = new ExportHelper<T>();
            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet != 0)
            {
                //TODO:数据为空判断
                using (helper.CurrentExcelPackage)
                {
                    var sheetCount = (int)(dataItems.Count / helper.ExcelExporterSettings.MaxRowNumberOnASheet) + ((dataItems.Count % helper.ExcelExporterSettings.MaxRowNumberOnASheet) > 0 ? 1 : 0);
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = dataItems.Skip(i * helper.ExcelExporterSettings.MaxRowNumberOnASheet).Take(helper.ExcelExporterSettings.MaxRowNumberOnASheet).ToList();
                        helper.AddExcelWorksheet();
                        helper.Export(sheetDataItems);
                    }
                    return Task.FromResult(helper.CurrentExcelPackage.GetAsByteArray());
                }
            }
            else
            {
                using (var ep = helper.Export(dataItems))
                {
                    return Task.FromResult(ep.GetAsByteArray());
                }
            }

        }

        //public Task<ExportFileInfo> Export<T>(string fileName, DataTable dataItems) where T : class
        //{
        //    CheckFileName(fileName);

        //    var fileInfo = ExcelHelper.CreateExcelPackage(fileName, excelPackage =>
        //    {
        //        //导出定义
        //        var exporter = GetExporterAttribute<T>();

        //        if (exporter?.Author != null)
        //            excelPackage.Workbook.Properties.Author = exporter?.Author;

        //        if (GetExporterHeaderInfoList<T>(out var exporterHeaderList, dataItems.Columns))
        //            return;

        //        var data = dataItems.SplitDataTable();

        //        var count = 0;
        //        foreach (DataTable table in data.Tables)
        //        {
        //            var sheet = GetWorksheet(excelPackage, exporter, count);
        //            AddWorksheet<T>(table, exporter, exporterHeaderList, sheet);
        //            count++;
        //        }
        //    });
        //    return Task.FromResult(fileInfo);
        //}

        /// <summary>
        /// 导出字节
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        //public Task<byte[]> ExportAsByteArray<T>(DataTable dataItems) where T : class
        //{
        //    using (var excelPackage = new ExcelPackage())
        //    {
        //        //导出定义
        //        var exporter = GetExporterAttribute<T>();

        //        if (exporter?.Author != null)
        //            excelPackage.Workbook.Properties.Author = exporter?.Author;

        //        if (GetExporterHeaderInfoList<T>(out var exporterHeaderList, dataItems.Columns))
        //            return null;

        //        var data = dataItems.SplitDataTable();
        //        var count = 0;
        //        foreach (DataTable table in data.Tables)
        //        {
        //            var sheet = GetWorksheet(excelPackage, exporter, count);
        //            AddWorksheet<T>(table, exporter, exporterHeaderList, sheet);
        //            count++;
        //        }

        //        return Task.FromResult(excelPackage.GetAsByteArray());
        //    }
        //}

        /// <summary>
        ///     导出excel表头
        /// </summary>
        /// <param name="items">表头数组</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <param name="globalStyle">全局样式</param>
        /// <param name="styles">样式</param>
        /// <returns></returns>
        //public Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName, ExcelHeadStyle globalStyle = null,
        //    List<ExcelHeadStyle> styles = null)
        //{
        //    using (var excelPackage = new ExcelPackage())
        //    {
        //        var sheet = excelPackage.Workbook.Worksheets.Add(sheetName ?? "导出结果");
        //        sheet.OutLineApplyStyle = true;
        //        AddHeader(items, sheet);
        //        AddStyle(sheet, items.Length, globalStyle, styles);
        //        return Task.FromResult(excelPackage.GetAsByteArray());
        //    }
        //}

        /// <summary>
        ///     导出Excel表头
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>文件二进制数组</returns>
        //public Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class
        //{
        //    using (var excelPackage = new ExcelPackage())
        //    {
        //        //导出定义
        //        var exporter = GetExporterAttribute<T>();

        //        if (exporter?.Author != null)
        //            excelPackage.Workbook.Properties.Author = exporter?.Author;

        //        var sheet = excelPackage.Workbook.Worksheets.Add(exporter?.Name ?? "导出结果");
        //        sheet.OutLineApplyStyle = true;
        //        if (GetExporterHeaderInfoList<T>(out var exporterHeaderList)) return null;
        //        AddHeader(exporterHeaderList, sheet, exporter);
        //        AddStyle(exporter, exporterHeaderList, sheet);
        //        return Task.FromResult(excelPackage.GetAsByteArray());
        //    }
        //}




        //private static ExcelWorksheet GetWorksheet(ExcelPackage excelPackage, ExcelExporterAttribute exporter,
        //    int count = 0)
        //{
        //    var name = exporter?.Name ?? "导出结果";
        //    var sheet = excelPackage.Workbook.Worksheets.Add($"{name}-{count}");
        //    sheet.OutLineApplyStyle = true;
        //    return sheet;
        //}



        /// <summary>
        ///     创建表头
        /// </summary>
        /// <param name="exporterHeaderDtoList"></param>
        /// <param name="sheet"></param>
        /// <param name="exporter"></param>
        //protected void AddHeader(List<ExporterHeaderInfo> exporterHeaderDtoList, ExcelWorksheet sheet,
        //    ExcelExporterAttribute exporter)
        //{
        //    foreach (var exporterHeaderDto in exporterHeaderDtoList)
        //        if (exporterHeaderDto != null)
        //        {
        //            if (exporterHeaderDto.ExporterHeaderAttribute != null)
        //            {
        //                var exporterHeaderAttribute = exporterHeaderDto.ExporterHeaderAttribute;
        //                if (exporterHeaderAttribute != null && !exporterHeaderAttribute.IsIgnore)
        //                {
        //                    var name = exporterHeaderAttribute.DisplayName.IsNullOrWhiteSpace()
        //                        ? exporterHeaderDto.PropertyName
        //                        : exporterHeaderAttribute.DisplayName;

        //                    sheet.Cells[1, exporterHeaderDto.Index].Value = ColumnHeaderStringFunc(name);
        //                    sheet.Cells[1, exporterHeaderDto.Index].Style.Font.Bold = exporterHeaderAttribute.IsBold;

        //                    var size = exporter?.HeaderFontSize ?? exporterHeaderAttribute.FontSize;
        //                    if (size.HasValue)
        //                        sheet.Cells[1, exporterHeaderDto.Index].Style.Font.Size = size.Value;
        //                }
        //            }
        //            else
        //            {
        //                sheet.Cells[1, exporterHeaderDto.Index].Value =
        //                    ColumnHeaderStringFunc(exporterHeaderDto.PropertyName);
        //            }
        //        }
        //}

        /// <summary>
        ///     创建表头
        /// </summary>
        /// <param name="exporterHeaders">表头数组</param>
        /// <param name="sheet">工作簿</param>
        //protected void AddHeader(string[] exporterHeaders, ExcelWorksheet sheet)
        //{
        //    var columnIndex = 0;
        //    foreach (var exporterHeader in exporterHeaders)
        //        if (exporterHeader != null)
        //        {
        //            columnIndex++;
        //            sheet.Cells[1, columnIndex].Value = exporterHeader;
        //        }
        //}






        /// <summary>
        ///     添加样式
        /// </summary>
        /// <param name="sheet">excel工作簿</param>
        /// <param name="columns">总列数</param>
        /// <param name="globalStyle">全局样式</param>
        /// <param name="styles">样式</param>
        //protected void AddStyle(ExcelWorksheet sheet, int columns, ExcelHeadStyle globalStyle = null,
        //    List<ExcelHeadStyle> styles = null)
        //{
        //    var col = 0;
        //    if (styles != null)
        //    {
        //        foreach (var style in styles)
        //        {
        //            col++;
        //            if (col <= columns)
        //            {
        //                if (style.IsIgnore)
        //                {
        //                    sheet.DeleteColumn(col);
        //                    continue;
        //                }

        //                var excelCol = sheet.Column(col);
        //                if (!style.Format.IsNullOrWhiteSpace()) excelCol.Style.Numberformat.Format = style.Format;
        //                excelCol.Style.Font.Bold = style.IsBold;
        //                excelCol.Style.Font.Size = style.FontSize;

        //                if (style.IsAutoFit) excelCol.AutoFit();
        //            }
        //        }
        //    }
        //    else
        //    {
        //        if (globalStyle != null)
        //            for (var i = 1; i <= columns; i++)
        //            {
        //                if (globalStyle.IsIgnore)
        //                {
        //                    sheet.DeleteColumn(i);
        //                    continue;
        //                }

        //                var excelCol = sheet.Column(i);
        //                if (!globalStyle.Format.IsNullOrWhiteSpace())
        //                    excelCol.Style.Numberformat.Format = globalStyle.Format;
        //                excelCol.Style.Font.Bold = globalStyle.IsBold;
        //                excelCol.Style.Font.Size = globalStyle.FontSize;

        //                if (globalStyle.IsAutoFit) excelCol.AutoFit();
        //            }
        //    }
        //}





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