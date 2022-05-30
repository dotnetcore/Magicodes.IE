using NPOI.XSSF.UserModel;
using System.IO;

namespace Magicodes.ExporterAndImporter.Excel.NPOI
{
    /// <summary>
    ///     扩展类
    /// </summary>
    public static class Extension
    {
        public static byte[] SaveToExcelWithXSSFWorkbook(this byte[] data)
        {
            //for excel compability
            using (var stream = new MemoryStream(data))
            {
                XSSFWorkbook wb = new XSSFWorkbook(stream);
                using (MemoryStream ms = new MemoryStream())
                {
                    wb.Write(ms);
                    return ms.ToArray();
                }
            }
        }

        /// <summary>
        /// 导出DataTable With XSSFWorkbook
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportWithXSSFWorkbook<T>(string fileName, DataTable dataItems) where T : class, new()
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAsByteArray<T>(dataItems);
            bytes =bytes.SaveToExcelWithXSSFWorkbook();
            return bytes.ToExcelExportFileInfo(fileName);
        }

        /// <summary>
        ///     导出Excel with XSSFWorkbook
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="dataItems">数据列</param>
        /// <returns>文件</returns>
        public async Task<ExportFileInfo> ExportWithXSSFWorkbook<T>(string fileName, ICollection<T> dataItems) where T : class, new()
        {
            var bytes = await ExportWithXSSFWorkbookAsByteArray(dataItems);
            return bytes.ToExcelExportFileInfo(fileName);
        }

        /// <summary>
        /// 导出字节 With XSSFWorkbook
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public Task<byte[]> ExportWithXSSFWorkbookAsByteArray<T>(DataTable dataItems) where T : class, new()
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
                    return Task.FromResult(NPOI.Extension.SaveToExcelWithXSSFWorkbook(helper.CurrentExcelPackage.GetAsByteArray()));
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
        public Task<byte[]> ExportWithXSSFWorkbookAsByteArray(DataTable dataItems, Type type)
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

                    return Task.FromResult(NPOI.Extension.SaveToExcelWithXSSFWorkbook(helper.CurrentExcelPackage.GetAsByteArray()));
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
        ///     导出Excel
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <returns>文件二进制数组</returns>
        public Task<byte[]> ExportWithXSSFWorkbookAsByteArray<T>(ICollection<T> dataItems) where T : class, new()
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

                    return Task.FromResult(NPOI.Extension.SaveToExcelWithXSSFWorkbook(helper.CurrentExcelPackage.GetAsByteArray()));
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
        ///     导出excel表头 With XSSFWorkbook
        /// </summary>
        /// <param name="items">表头数组</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns></returns>
        public Task<byte[]> ExportWithXSSFWorkbookHeaderAsByteArray(string[] items, string sheetName = "导出结果")
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
                return Task.FromResult(NPOI.Extension.SaveToExcelWithXSSFWorkbook(ep.GetAsByteArray()));
            }
        }

        /// <summary>
        ///     导出Excel表头
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>文件二进制数组</returns>
        public Task<byte[]> ExportHeaderWithXSSFWorkbookAsByteArray<T>(T type) where T : class, new()
        {
            var helper = new ExportHelper<T>();
            using (var ep = helper.ExportHeaders())
            {
                return Task.FromResult(NPOI.Extension.SaveToExcelWithXSSFWorkbook(ep.GetAsByteArray()));
            }
        }

        public async Task<ExportFileInfo> ExportWithXSSFWorkbook(string fileName, DataTable dataItems,
          IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000)
        {
            fileName.CheckExcelFileName();
            var bytes = await ExportAsByteArray(dataItems, exporterHeaderFilter, maxRowNumberOnASheet);
            bytes = NPOI.Extension.SaveToExcelWithXSSFWorkbook(bytes);

            return bytes.ToExcelExportFileInfo(fileName);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="exporterHeaderFilter"></param>
        /// <param name="maxRowNumberOnASheet"></param>
        /// <returns></returns>
        public Task<byte[]> ExportWithXSSFWorkbook(DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null,
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