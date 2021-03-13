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

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel.Utility;
using Magicodes.ExporterAndImporter.Excel.Utility.TemplateExport;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     Excel导出程序
    /// </summary>
    public class ExcelExporter : IExcelExporter
    {
        private ExcelPackage _excelPackage;
        private bool _isSeparateColumn;
        private bool _isSeparateBySheet;
        private bool _isSeparateByRow;
        private bool _isAppendHeaders;

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="dataItems">数据列</param>
        /// <returns>文件</returns>
        public async Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class, new()
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAsByteArray(dataItems);
            return bytes.ToExcelExportFileInfo(fileName);
        }

        /// <summary>
        /// append collectioin to context
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelExporter Append<T>(ICollection<T> dataItems, string sheetName = null) where T : class, new()
        {
            var helper = this._excelPackage == null ? new ExportHelper<T>(sheetName) : new ExportHelper<T>(_excelPackage, sheetName);
            if (_isSeparateColumn || _isSeparateBySheet || _isSeparateByRow)
            {
                var name = helper.ExcelExporterSettings?.Name ?? "导出结果";

                if (this._excelPackage?.Workbook.Worksheets.Any(x => x.Name == name) ?? false)
                {
                    throw new ArgumentNullException($"已经存在名字为{name}的sheet");
                }
            }

            this._excelPackage = helper.Export(dataItems);

            if (_isSeparateColumn)
            {
#if NET461
                helper.CopySheet(1,
                    2);
#else
                helper.CopySheet(0,
                      1);
#endif

                _isSeparateColumn = false;
            }

            if (_isSeparateByRow)
            {
#if NET461
                helper.CopyRows(1,
                    2, _isAppendHeaders);
#else
                helper.CopyRows(0,
                      1, _isAppendHeaders);
#endif
            }

            _isSeparateBySheet = false;
            _isSeparateByRow = false;
            _isAppendHeaders = false;
            return this;
        }


        /// <summary>
        ///		分割集合到当前Sheet追加Column
        /// </summary>
        /// <returns></returns>
        public ExcelExporter SeparateByColumn()
        {
            if (_excelPackage == null)
            {
                throw new ArgumentNullException("调用当前方法之前，必须先调用Append方法！");
            }

            _isSeparateColumn = true;
            return this;
        }

        /// <summary>
        ///     分割多出多个sheet
        /// </summary>
        /// <returns></returns>
        public ExcelExporter SeparateBySheet()
        {
            if (_excelPackage == null)
            {
                throw new ArgumentNullException("调用当前方法之前，必须先调用Append方法！");
            }

            _isSeparateBySheet = true;
            return this;
        }

        /// <summary>
        ///     追加rows到当前sheet
        /// </summary>
        /// <returns></returns>
        public ExcelExporter SeparateByRow()
        {
            if (_excelPackage == null)
            {
                throw new ArgumentNullException("调用当前方法之前，必须先调用Append方法！");
            }

            _isSeparateByRow = true;
            return this;
        }

        /// <summary>
        ///     追加表头
        /// </summary>
        /// <returns></returns>
        public ExcelExporter AppendHeaders()
        {
            if (_excelPackage == null)
            {
                throw new ArgumentNullException("调用当前方法之前，必须先调用Append方法！");
            }

            if (!_isSeparateByRow)
            {
                throw new ArgumentNullException("调用当前方法之前，必须先调用SeparateByRow方法！");
            }

            _isAppendHeaders = true;
            return this;
        }

        /// <summary>
        /// 导出所有的追加数据
        /// </summary>
        /// <returns></returns>
        public Task<byte[]> ExportAppendDataAsByteArray()
        {
            if (this._excelPackage == null)
            {
                throw new ArgumentNullException("调用当前方法之前，必须先调用Append方法！");
            }

            return Task.FromResult(_excelPackage.GetAsByteArray());
        }

        /// <summary>
        /// export excel after append all collectioins
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportAppendData(string fileName)
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAppendDataAsByteArray();
            return bytes.ToExcelExportFileInfo(fileName);
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <returns>文件二进制数组</returns>
        public Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems) where T : class, new()
        {
            var helper = new ExportHelper<T>();
            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var sheetCount = (int)(dataItems.Count / helper.ExcelExporterSettings.MaxRowNumberOnASheet) +
                                     ((dataItems.Count % helper.ExcelExporterSettings.MaxRowNumberOnASheet) > 0
                                         ? 1
                                         : 0);
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = dataItems.Skip(i * helper.ExcelExporterSettings.MaxRowNumberOnASheet)
                            .Take(helper.ExcelExporterSettings.MaxRowNumberOnASheet).ToList();
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

        /// <summary>
        /// 导出DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> Export<T>(string fileName, DataTable dataItems) where T : class, new()
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAsByteArray<T>(dataItems);
            return bytes.ToExcelExportFileInfo(fileName);
        }

        /// <summary>
        /// 导出字节
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public Task<byte[]> ExportAsByteArray<T>(DataTable dataItems) where T : class, new()
        {
            var helper = new ExportHelper<T>();
            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Rows.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var ds = dataItems.SplitDataTable(helper.ExcelExporterSettings.MaxRowNumberOnASheet);
                    var sheetCount = ds.Tables.Count;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = ds.Tables[i];
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

        /// <summary>
        /// 导出字节
        /// </summary>
        /// <param name="type"></param>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public Task<byte[]> ExportAsByteArray(DataTable dataItems, Type type)
        {
            var helper = new ExportHelper<DataTable>(type);
            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Rows.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var ds = dataItems.SplitDataTable(helper.ExcelExporterSettings.MaxRowNumberOnASheet);
                    var sheetCount = ds.Tables.Count;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = ds.Tables[i];
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

        /// <summary>
        ///     导出excel表头
        /// </summary>
        /// <param name="items">表头数组</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns></returns>
        public Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName = "导出结果")
        {
            var helper = new ExportHelper<DataTable>();
            var headerList = new List<ExporterHeaderInfo>();
            for (var i = 1; i <= items.Length; i++)
            {
                var item = items[i - 1];
                var exporterHeaderInfo =
                    new ExporterHeaderInfo()
                    {
                        Index = i,
                        DisplayName = item,
                        CsTypeName = "string",
                        PropertyName = item,
                        ExporterHeaderAttribute = new ExporterHeaderAttribute(item) { },
                    };
                headerList.Add(exporterHeaderInfo);
            }

            helper.AddExcelWorksheet(sheetName);
            helper.AddExporterHeaderInfoList(headerList);
            using (var ep = helper.ExportHeaders())
            {
                return Task.FromResult(ep.GetAsByteArray());
            }
        }

        /// <summary>
        ///     导出Excel表头
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>文件二进制数组</returns>
        public Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class, new()
        {
            var helper = new ExportHelper<T>();
            using (var ep = helper.ExportHeaders())
            {
                return Task.FromResult(ep.GetAsByteArray());
            }
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
                var file = new FileInfo(fileName);

                helper.Export(template, data, (package) => { package.SaveAs(file); });
                return Task.FromResult(new ExportFileInfo(file.Name, file.Extension));
            }
        }

        /// <summary>
        ///     根据模板导出到载荷
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="template">HTML模板或模板路径</param>
        public Task<byte[]> ExportBytesByTemplate<T>(T data, string template) where T : class
        {
            using (var helper = new TemplateExportHelper<T>())
            {
                using (var sr = new MemoryStream())
                {
                    helper.Export(template, data, (package) => { package.SaveAs(sr); });
                    return Task.FromResult(sr.ToArray());
                }
            }
        }

        /// <summary>
        ///		根据模板导出
        /// </summary>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public Task<byte[]> ExportBytesByTemplate(object data, string template, Type type)
        {
            throw new NotImplementedException();
        }

        public async Task<ExportFileInfo> Export(string fileName, DataTable dataItems,
            IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000)
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAsByteArray(dataItems, exporterHeaderFilter, maxRowNumberOnASheet);
            return bytes.ToExcelExportFileInfo(fileName);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="exporterHeaderFilter"></param>
        /// <param name="maxRowNumberOnASheet"></param>
        /// <returns></returns>
        public Task<byte[]> ExportAsByteArray(DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null,
            int maxRowNumberOnASheet = 1000000)
        {
            var helper = new ExportHelper<DataTable>();
            helper.ExcelExporterSettings.MaxRowNumberOnASheet = maxRowNumberOnASheet;
            helper.SetExporterHeaderFilter(exporterHeaderFilter);

            if (helper.ExcelExporterSettings.MaxRowNumberOnASheet > 0 &&
                dataItems.Rows.Count > helper.ExcelExporterSettings.MaxRowNumberOnASheet)
            {
                using (helper.CurrentExcelPackage)
                {
                    var ds = dataItems.SplitDataTable(helper.ExcelExporterSettings.MaxRowNumberOnASheet);
                    var sheetCount = ds.Tables.Count;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        var sheetDataItems = ds.Tables[i];
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


    }
}