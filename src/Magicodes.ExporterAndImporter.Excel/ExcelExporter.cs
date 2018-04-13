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
                 var exporter = GetExporterAttribute<T>();

                 var sheet = excelPackage.Workbook.Worksheets.Add(fileName);
                 sheet.OutLineApplyStyle = true;
                 if (GetExporterHeaderDtoList<T>(out var exporterHeaderDtoList)) return;
                 AddHeader(exporterHeaderDtoList, sheet, exporter);
                 AddDataItems(sheet, 1, dataItems, exporter);
                 AddStyle(exporterHeaderDtoList, sheet);
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
                if (exporterHeaderDto != null && exporterHeaderDto.Attribute.Length > 0)
                {
                    var exporterHeaderAttribute = exporterHeaderDto.Attribute[0];
                    if (exporterHeaderAttribute != null)
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
            }
        }

        /// <summary>
        ///     添加导出数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="items"></param>
        protected void AddDataItems<T>(ExcelWorksheet sheet, int startRowIndex, IList<T> items, ExcelExporterAttribute exporter)
        {
            if (items == null || items.Count == 0)
                return;
            var tbStyle = TableStyles.Medium10;
            if (exporter != null && exporter.TableStyle.HasValue) tbStyle = exporter.TableStyle.Value;

            sheet.Cells["A" + startRowIndex].LoadFromCollection(items, true, tbStyle);
        }

        /// <summary>
        ///     添加样式
        /// </summary>
        /// <param name="exporterHeaderDtos"></param>
        /// <param name="sheet"></param>
        protected void AddStyle(List<ExporterHeaderInfo> exporterHeaderDtos, ExcelWorksheet sheet)
        {
            foreach (var exporterHeaderDto in exporterHeaderDtos)
            {
                if (exporterHeaderDto.Attribute.Length > 0)
                {
                    var exporterHeaderAttribute = exporterHeaderDto.Attribute[0];
                    if (exporterHeaderAttribute != null)
                    {
                        sheet.Column(exporterHeaderDto.Index).Style.Numberformat.Format = exporterHeaderAttribute.Format;
                        if (exporterHeaderAttribute.IsAutoFit)
                            sheet.Column(exporterHeaderDto.Index).AutoFit();
                    }
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

        private static bool GetExporterHeaderDtoList<T>(out List<ExporterHeaderInfo> exporterHeaderDtoList)
        {
            exporterHeaderDtoList = new List<ExporterHeaderInfo>();
            var objProperties = typeof(T).GetProperties();
            if (objProperties == null || objProperties.Length == 0)
                return true;
            for (var i = 0; i < objProperties.Length; i++)
            {
                exporterHeaderDtoList.Add(new ExporterHeaderInfo
                {
                    Index = i + 1,
                    PropertyName = objProperties[i].Name,
                    Attribute = objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as ExporterHeaderAttribute[]
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
                    TableStyle = TableStyles.Medium10,
                    HeaderFontSize = export.HeaderFontSize
                };
            }
            return null;
        }
    }
}
