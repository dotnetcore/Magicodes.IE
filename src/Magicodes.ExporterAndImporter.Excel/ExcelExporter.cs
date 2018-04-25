using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Linq;
using System.Reflection;
using System.ComponentModel;

namespace Magicodes.ExporterAndImporter.Excel
{
    public class ExcelExporter : IExporter
    {
        /// <summary>
        /// 语言处理函数
        /// </summary>
        public static Func<string, string> LocalStringFunc { get; set; } = (str) => str;

        public ExcelExporter()
        {
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="dataItems">数据列</param>
        /// <returns>文件</returns>
        public Task<ExportFileInfo> Export<T>(string fileName, IList<T> dataItems) where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("文件名必须填写!", nameof(fileName));
            }
            var extension = Path.GetExtension(fileName);
            if (string.IsNullOrWhiteSpace(extension))
            {
                fileName = fileName + ".xlsx";
            }
            var fileInfo = CreateExcelPackage(fileName, excelPackage =>
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
        ///     创建Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="creator"></param>
        /// <returns></returns>
        protected ExportFileInfo CreateExcelPackage(string fileName, Action<ExcelPackage> creator)
        {
            var file = new ExportFileInfo(fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            using (var excelPackage = new ExcelPackage())
            {
                creator(excelPackage);
                Save(excelPackage, file);
            }
            return file;
        }

        /// <summary>
        ///     创建表头
        /// </summary>
        /// <param name="exporterHeaderDtoList"></param>
        /// <param name="sheet"></param>
        protected void AddHeader(List<ExporterHeaderInfo> exporterHeaderDtoList, ExcelWorksheet sheet, ExcelExporterAttribute exporter)
        {
            foreach (var exporterHeaderDto in exporterHeaderDtoList)
            {
                if (exporterHeaderDto != null)
                {
                    if (exporterHeaderDto.ExporterHeader != null)
                    {
                        var exporterHeaderAttribute = exporterHeaderDto.ExporterHeader;
                        if (exporterHeaderAttribute != null && !exporterHeaderAttribute.IsIgnore)
                        {
                            var name = exporterHeaderAttribute.DisplayName.IsNullOrWhiteSpace() ? exporterHeaderDto.PropertyName
                                        : exporterHeaderAttribute.DisplayName;

                            sheet.Cells[1, exporterHeaderDto.Index].Value = LocalStringFunc(name);
                            sheet.Cells[1, exporterHeaderDto.Index].Style.Font.Bold = exporterHeaderAttribute.IsBold;

                            var size = exporter?.HeaderFontSize ?? exporterHeaderAttribute.FontSize;
                            if (size.HasValue)
                                sheet.Cells[1, exporterHeaderDto.Index].Style.Font.Size = size.Value;
                        }
                    }
                    else
                    {
                        sheet.Cells[1, exporterHeaderDto.Index].Value = LocalStringFunc(exporterHeaderDto.PropertyName);
                    }

                }

            }
        }

        /// <summary>
        ///     添加导出数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="items"></param>
        protected void AddDataItems<T>(ExcelWorksheet sheet, List<ExporterHeaderInfo> exporterHeaders, IList<T> items, ExcelExporterAttribute exporter)
        {
            if (items == null || items.Count == 0)
                return;
            var tbStyle = TableStyles.Medium10;
            if (exporter != null && !exporter.TableStyle.IsNullOrWhiteSpace())
                tbStyle = (TableStyles)Enum.Parse(typeof(TableStyles), exporter.TableStyle);
            sheet.Cells["A2"].LoadFromCollection(items, false, tbStyle);
        }


        /// <summary>
        ///     添加样式
        /// </summary>
        /// <param name="exporterHeaders"></param>
        /// <param name="sheet"></param>
        protected void AddStyle(ExcelExporterAttribute exporter, List<ExporterHeaderInfo> exporterHeaders, ExcelWorksheet sheet)
        {
            foreach (var exporterHeader in exporterHeaders)
            {
                if (exporterHeader.ExporterHeader != null)
                {
                    if (exporterHeader.ExporterHeader.IsIgnore)
                    {
                        //TODO:后续直接修改数据导出逻辑（不写忽略列数据）
                        sheet.DeleteColumn(exporterHeader.Index);
                        //删除之后，序号依次-1
                        foreach (var item in exporterHeaders.Where(p => p.Index > exporterHeader.Index))
                        {
                            item.Index--;
                        }
                        continue;
                    }

                    var col = sheet.Column(exporterHeader.Index);
                    col.Style.Numberformat.Format = exporterHeader.ExporterHeader.Format;

                    if (exporter.AutoFitAllColumn || exporterHeader.ExporterHeader.IsAutoFit)
                        col.AutoFit();
                }
            }
        }

        /// <summary>
        ///     保存
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="file"></param>
        protected void Save(ExcelPackage excelPackage, ExportFileInfo file)
        {
            excelPackage.SaveAs(new FileInfo(file.FileName));
        }

        /// <summary>
        /// 获取头部定义
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
            {
                exporterHeaderList.Add(new ExporterHeaderInfo
                {
                    Index = i + 1,
                    PropertyName = objProperties[i].Name,
                    ExporterHeader = (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as ExporterHeaderAttribute[])?.FirstOrDefault()
                });
            }
            return false;
        }

        /// <summary>
        /// 获取表全局定义
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private static ExcelExporterAttribute GetExporterAttribute<T>() where T : class
        {
            var exporterTableAttributes = (typeof(T).GetCustomAttributes(typeof(ExcelExporterAttribute), true) as ExcelExporterAttribute[]);
            if (exporterTableAttributes != null && exporterTableAttributes.Length > 0)
                return exporterTableAttributes[0];

            var exporterAttributes = (typeof(T).GetCustomAttributes(typeof(ExporterAttribute), true) as ExporterAttribute[]);

            if (exporterAttributes != null && exporterAttributes.Length > 0)
            {
                var export = exporterAttributes[0];
                return new ExcelExporterAttribute()
                {
                    FontSize = export.FontSize,
                    HeaderFontSize = export.HeaderFontSize
                };
            }
            return null;
        }
    }
}
