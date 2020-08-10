using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Dynamic.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using System.Drawing;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Globalization;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    /// <summary>
    /// 导出辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExportHelper<T> where T : class
    {
        private ExcelExporterAttribute _excelExporterAttribute;
        private ExcelWorksheet _excelWorksheet;
        private ExcelPackage _excelPackage;
        private List<ExporterHeaderInfo> _exporterHeaderList;
        private Type _type;
        private string _sheetName;
        /// <summary>
        /// 
        /// </summary>
        public ExportHelper(string sheetName = null)
        {
            if (typeof(DataTable).Equals(typeof(T)))
            {
                IsDynamicDatableExport = true;
            }

            _sheetName = sheetName;
        }
        /// <summary>
        /// </summary>
        /// <param name="type"></param>
        public ExportHelper(Type type)
        {
            this._type = type;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="existExcelPackage"></param>

        public ExportHelper(ExcelPackage existExcelPackage, string sheetName = null)
        {
            if (typeof(DataTable).Equals(typeof(T)))
            {
                IsDynamicDatableExport = true;
            }
            if (existExcelPackage != null)
            {
                this._excelPackage = existExcelPackage;
            }
            _sheetName = sheetName;
        }


        /// <summary>
        ///     导出设置
        /// </summary>
        public ExcelExporterAttribute ExcelExporterSettings
        {
            get
            {
                if (_excelExporterAttribute == null)
                {
                    var type = _type ?? typeof(T);
                    if (typeof(DataTable) == type)
                    {
                        _excelExporterAttribute = new ExcelExporterAttribute();
                    }
                    else
                        _excelExporterAttribute = type.GetAttribute<ExcelExporterAttribute>(true);

                    if (_excelExporterAttribute == null)
                    {
                        var exporterAttribute = type.GetAttribute<ExporterAttribute>(true);
                        if (exporterAttribute != null)
                        {
                            _excelExporterAttribute = new ExcelExporterAttribute()
                            {
                                Author = exporterAttribute.Author,
                                AutoFitAllColumn = exporterAttribute.AutoFitAllColumn,
                                //ExcelOutputType = 
                                ExporterHeaderFilter = exporterAttribute.ExporterHeaderFilter,
                                FontSize = exporterAttribute.FontSize,
                                HeaderFontSize = exporterAttribute.HeaderFontSize,
                                MaxRowNumberOnASheet = exporterAttribute.MaxRowNumberOnASheet,
                                Name = exporterAttribute.Name,
                                TableStyle = exporterAttribute.TableStyle,
                                AutoCenter = _excelExporterAttribute != null && _excelExporterAttribute.AutoCenter
                            };
                        }
                        else
                            _excelExporterAttribute = new ExcelExporterAttribute();
                    }

                    //加载表头筛选器
                    if (_excelExporterAttribute.ExporterHeaderFilter != null && typeof(IExporterHeaderFilter).IsAssignableFrom(_excelExporterAttribute.ExporterHeaderFilter))
                    {
                        ExporterHeaderFilter = (IExporterHeaderFilter)_excelExporterAttribute.ExporterHeaderFilter.Assembly.CreateInstance(_excelExporterAttribute.ExporterHeaderFilter.FullName, true, System.Reflection.BindingFlags.Default, null, _excelExporterAttribute.ExporterHeaderFilter.CreateType(), null, null);
                    }
                }
                return _excelExporterAttribute;
            }
            set => _excelExporterAttribute = value;
        }

        /// <summary>
        /// 当前工作Sheet
        /// </summary>
        protected ExcelWorksheet CurrentExcelWorksheet
        {
            get
            {
                if (_excelWorksheet == null)
                {
                    AddExcelWorksheet(_sheetName);
                }


                return _excelWorksheet;
            }
            set => _excelWorksheet = value;
        }

        /// <summary>
        /// 当前Sheet索引
        /// </summary>
        protected int SheetIndex = 0;
        private string _exporterHeadersString;

        /// <summary>
        /// 当前工作
        /// </summary>
        protected List<ExcelWorksheet> ExcelWorksheets { get; set; } = new List<ExcelWorksheet>();

        /// <summary>
        /// 当前Excel包
        /// </summary>
        public ExcelPackage CurrentExcelPackage
        {
            get
            {
                if (_excelPackage == null)
                {
                    _excelPackage = new ExcelPackage();

                    if (ExcelExporterSettings?.Author != null)
                        _excelPackage.Workbook.Properties.Author = ExcelExporterSettings?.Author;
                }
                return _excelPackage;
            }
            set => _excelPackage = value;
        }

        /// <summary>
        /// 表头列表
        /// </summary>
        protected List<ExporterHeaderInfo> ExporterHeaderList
        {
            get
            {
                if (_exporterHeaderList == null) GetExporterHeaderInfoList();
                if ((_exporterHeaderList == null || _exporterHeaderList.Count == 0) && !IsDynamicDatableExport) throw new Exception("请定义表头！");
                return _exporterHeaderList;
            }
            set => _exporterHeaderList = value;
        }

        /// <summary>
        /// 列头表达式
        /// </summary>
        protected string ExporterHeadersString
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_exporterHeadersString))
                {
                    _exporterHeadersString = string.Join(",", ExporterHeaderList.Select(p => p.PropertyName));
                }
                return _exporterHeadersString;
            }
            set => _exporterHeadersString = value;
        }

        /// <summary>
        /// Excel数据表
        /// </summary>
        protected ExcelTable CurrentExcelTable { get; set; }

        /// <summary>
        /// 表头筛选器
        /// </summary>
        protected IExporterHeaderFilter ExporterHeaderFilter { get; set; }

        /// <summary>
        /// 是否为动态DataTable导出
        /// </summary>
        protected bool IsDynamicDatableExport { get; set; }

        /// <summary>
        /// 添加导出表头
        /// </summary>
        /// <param name="exporterHeaderInfos"></param>
        public virtual void AddExporterHeaderInfoList(List<ExporterHeaderInfo> exporterHeaderInfos)
        {
            _exporterHeaderList = exporterHeaderInfos;
        }

        /// <summary>
        ///     获取头部定义
        /// </summary>
        /// <returns></returns>
        protected virtual void GetExporterHeaderInfoList(DataTable dt = null)
        {
            _exporterHeaderList = new List<ExporterHeaderInfo>();
            if (dt != null)
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    var item = new ExporterHeaderInfo
                    {
                        Index = i + 1,
                        PropertyName = dt.Columns[i].ColumnName,
                        ExporterHeaderAttribute = new ExporterHeaderAttribute(dt.Columns[i].ColumnName),
                        CsTypeName = dt.Columns[i].DataType.GetCSharpTypeName(),
                        DisplayName = dt.Columns[i].ColumnName
                    };
                    AddExportHeaderInfo(item);
                }
            }
            else if (!IsDynamicDatableExport)
            {
                var type = _type ?? typeof(T);
                var objProperties = type.GetProperties();
                if (objProperties.Length == 0)
                    return;
                for (var i = 0; i < objProperties.Length; i++)
                {

                    var item = new ExporterHeaderInfo
                    {
                        Index = i + 1,
                        PropertyName = objProperties[i].Name,
                        ExporterHeaderAttribute =
                            (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as
                                ExporterHeaderAttribute[])?.FirstOrDefault() ?? new ExporterHeaderAttribute(objProperties[i].GetDisplayName() ?? objProperties[i].Name),
                        CsTypeName = objProperties[i].PropertyType.GetCSharpTypeName(),
                        ExportImageFieldAttribute = objProperties[i].GetAttribute<ExportImageFieldAttribute>(true)

                    };

                    //设置列显示名
                    item.DisplayName = item.ExporterHeaderAttribute == null ||
                                       item.ExporterHeaderAttribute.DisplayName == null ||
                                       item.ExporterHeaderAttribute.DisplayName.IsNullOrWhiteSpace()
                        ? item.PropertyName
                        : item.ExporterHeaderAttribute.DisplayName;
                    //设置Format
                    item.ExporterHeaderAttribute.Format = item.ExporterHeaderAttribute.Format.IsNullOrWhiteSpace()
                        ? objProperties[i].GetDisplayFormat()
                        : item.ExporterHeaderAttribute.Format;
                    //设置Ignore
                    item.ExporterHeaderAttribute.IsIgnore =
                        (objProperties[i].GetAttribute<IEIgnoreAttribute>(true) == null) ?
                        item.ExporterHeaderAttribute.IsIgnore : objProperties[i].GetAttribute<IEIgnoreAttribute>(true).IsExportIgnore;

                    AddExportHeaderInfo(item);
                }
            }
        }

        /// <summary>
        /// 添加列头并执行列头筛选器
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        protected virtual void AddExportHeaderInfo(ExporterHeaderInfo item)
        {
            //执行列头筛选器
            if (ExporterHeaderFilter != null)
            {
                item = ExporterHeaderFilter.Filter(item);
            }
            _exporterHeaderList.Add(item);
        }

        /// <summary>
        /// 设置列头筛选器
        /// </summary>
        public virtual void SetExporterHeaderFilter(IExporterHeaderFilter exporterHeaderFilter)
        {
            ExporterHeaderFilter = exporterHeaderFilter;
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <returns>文件</returns>
        public virtual ExcelPackage Export(ICollection<T> dataItems)
        {
            AddDataItems(dataItems);
            //仅当存在图片表头才渲染图片
            if (ExporterHeaderList.Any(p => p.ExportImageFieldAttribute != null))
            {
                AddPictures(dataItems.Count);
            }
            return AddHeaderAndStyles();
        }

        /// <summary>
        ///     复制Sheet
        /// </summary>
        /// <param name="currentws"></param>
        /// <param name="tempws"></param>
        public void CopySheet(int currentws, int tempws)
        {
            var tempWorksheet = _excelPackage.Workbook.Worksheets[tempws];
            var ws = _excelPackage.Workbook.Worksheets[currentws];

            tempWorksheet.Cells[1, 1, tempWorksheet.Dimension.Rows, tempWorksheet.Dimension.Columns]
                .Copy(ws.Cells[1, ws.Dimension.End.Column + 2,
                    tempWorksheet.Dimension.Rows, ws.Dimension.End.Column + tempWorksheet.Dimension.End.Column]);

            _excelPackage.Workbook.Worksheets.Delete(tempWorksheet);
        }

        /// <summary>
        ///     复制Rows
        /// </summary>
        /// <param name="currentws"></param>
        /// <param name="tempws"></param>
        /// <param name="isAppendHeaders"></param>
        public void CopyRows(int currentws, int tempws, bool isAppendHeaders)
        {
            var tempWorksheet = _excelPackage.Workbook.Worksheets[tempws];
            var ws = _excelPackage.Workbook.Worksheets[currentws];

            int beginRows = 2;
            if (isAppendHeaders)
            {
                beginRows = 1;
            }

            tempWorksheet.Cells[beginRows, 1, tempWorksheet.Dimension.Rows, tempWorksheet.Dimension.Columns]
                .Copy(ws.Cells[ws.Dimension.Rows + 2, 1,
                    tempWorksheet.Dimension.Rows + ws.Dimension.Rows, tempWorksheet.Dimension.End.Column]);

            _excelPackage.Workbook.Worksheets.Delete(tempWorksheet);
        }


        /// <summary>
        ///     导出Excel空表头
        /// </summary>
        /// <returns>文件</returns>
        public virtual ExcelPackage ExportHeaders()
        {
            return AddHeaderAndStyles();
        }

        /// <summary>
        /// 添加表头、样式以及忽略列、格式处理
        /// </summary>
        /// <returns></returns>
        private ExcelPackage AddHeaderAndStyles()
        {
            AddHeader();

            if (ExcelExporterSettings.AutoFitAllColumn)
            {
                CurrentExcelWorksheet.Cells[CurrentExcelWorksheet.Dimension.Address].AutoFitColumns();
            }

            AddStyle();
            DeleteIgnoreColumns();
            //以便支持导出多Sheet
            SheetIndex++;
            return CurrentExcelPackage;
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        public ExcelPackage Export(DataTable dataItems)
        {
            if ((ExporterHeaderList == null || ExporterHeaderList.Count == 0) && IsDynamicDatableExport) GetExporterHeaderInfoList(dataItems);
            AddDataItems(dataItems);
            //TODO:动态导出暂不考虑支持图片导出，后续可以考虑通过约定实现
            //AddPictures(dataItems.Rows.Count);
            return AddHeaderAndStyles();
        }

        /// <summary>
        /// 添加Sheet
        /// 支持同一个数据拆成多个Sheet
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public ExcelWorksheet AddExcelWorksheet(string name = null)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                name = ExcelExporterSettings?.Name ?? "导出结果";
            }
            if (SheetIndex != 0)
            {
                name += "-" + SheetIndex;
            }
            _excelWorksheet = CurrentExcelPackage.Workbook.Worksheets.Add(name);
            _excelWorksheet.OutLineApplyStyle = true;
            ExcelWorksheets.Add(_excelWorksheet);
            return _excelWorksheet;
        }
        /// <summary>
        ///     添加导出数据
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="excelRange"></param>
        protected void AddDataItems(ICollection<T> dataItems, ExcelRangeBase excelRange = null)
        {
            if (excelRange == null)
                excelRange = CurrentExcelWorksheet.Cells["A1"];

            if (dataItems == null || dataItems.Count == 0)
                return;

            if (ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable)
            {
                var tbStyle = TableStyles.Medium10;
                if (!ExcelExporterSettings.TableStyle.IsNullOrWhiteSpace())
                    tbStyle = (TableStyles)Enum.Parse(typeof(TableStyles), ExcelExporterSettings.TableStyle);
                var er = excelRange.LoadFromCollection(dataItems, true, TableStyles.None);
                CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
            }
            else
            {
                excelRange.LoadFromCollection(dataItems, true, TableStyles.None);
            }
        }

        /// <summary>
        ///     添加图片
        /// </summary>
        protected void AddPictures(int rowCount)
        {
            for (var colIndex = 0; colIndex < ExporterHeaderList.Count; colIndex++)
            {
                if (ExporterHeaderList[colIndex].ExportImageFieldAttribute != null)
                {
                    for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                    {
                        var cell = CurrentExcelWorksheet.Cells[rowIndex + 1, colIndex + 1];
                        var url = cell.Text;
                        if (File.Exists(url) || url.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                cell.Value = string.Empty;
                                var bitmap = Extension.GetBitmapByUrl(url);
                                if (bitmap == null)
                                {
                                    cell.Value = ExporterHeaderList[colIndex].ExportImageFieldAttribute.Alt;
                                }
                                else
                                {
                                    var pic = CurrentExcelWorksheet.Drawings.AddPicture(Guid.NewGuid().ToString(), bitmap);
                                    pic.SetPosition(rowIndex, ExporterHeaderList[colIndex].ExportImageFieldAttribute.Height / 5, colIndex - 1, 0);
                                    CurrentExcelWorksheet.Row(rowIndex + 1).Height = ExporterHeaderList[colIndex].ExportImageFieldAttribute.Height;
                                    //pic.SetSize(ExporterHeaderList[colIndex].ExportImageFieldAttribute.Width * 7, ExporterHeaderList[colIndex].ExportImageFieldAttribute.Height);
                                    pic.SetSize(ExporterHeaderList[colIndex].ExportImageFieldAttribute.Width * 7, ExporterHeaderList[colIndex].ExportImageFieldAttribute.Height);
                                }

                            }
                            catch (Exception)
                            {
                                cell.Value = ExporterHeaderList[colIndex].ExportImageFieldAttribute.Alt;
                            }
                        }
                        else
                        {
                            cell.Value = ExporterHeaderList[colIndex].ExportImageFieldAttribute.Alt;
                        }
                    }
                }
                //else
                //{
                //    for (var j = 1; j <= rowCount; j++)
                //    {
                //        CurrentExcelWorksheet.Cells[j + 1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                //        CurrentExcelWorksheet.Cells[j + 1, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                //    }
                //}
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="excelRange"></param>
        protected void AddDataItems(DataTable dataTable, ExcelRangeBase excelRange = null)
        {
            if (excelRange == null)
                excelRange = CurrentExcelWorksheet.Cells["A1"];

            if (dataTable == null || dataTable.Rows.Count == 0)
                return;

            var tbStyle = TableStyles.Medium10;
            if (!ExcelExporterSettings.TableStyle.IsNullOrWhiteSpace())
                tbStyle = (TableStyles)Enum.Parse(typeof(TableStyles), ExcelExporterSettings.TableStyle);

            var er = excelRange.LoadFromDataTable(dataTable, true, tbStyle);
            CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
        }

        /// <summary>
        /// 删除忽略列
        /// </summary>  
        protected void DeleteIgnoreColumns()
        {
            var deletedCount = 0;
            foreach (var exporterHeaderDto in ExporterHeaderList.Where(p => p.ExporterHeaderAttribute.IsIgnore))
            {
                //TODO:后续重写底层逻辑，直接从数据层面拦截
                CurrentExcelWorksheet.DeleteColumn(exporterHeaderDto.Index - deletedCount);
                deletedCount++;
            }
        }

        /// <summary>
        ///     创建表头
        /// </summary>
        protected void AddHeader()
        {
            //NoneStyle的时候没创建Table
            //https://github.com/JanKallman/EPPlus/blob/4dacf27661b24d92e8ba3d03d51dd5468845e6c1/EPPlus/ExcelRangeBase.cs#L2013
            var isNoneStyle = ExcelExporterSettings.TableStyle == TableStyles.None.ToString();

            if (CurrentExcelTable == null && ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable && !isNoneStyle)
            {
                var cols = ExporterHeaderList.Count;
                var range = CurrentExcelWorksheet.Cells[1, 1, CurrentExcelWorksheet.Dimension?.End.Row ?? 10, cols];
                //https://github.com/dotnetcore/Magicodes.IE/issues/66
                CurrentExcelTable = CurrentExcelWorksheet.Tables.Add(range, $"Table{CurrentExcelWorksheet.Index}");
                CurrentExcelTable.ShowHeader = true;
                Enum.TryParse(ExcelExporterSettings.TableStyle, out TableStyles outStyle);
                CurrentExcelTable.TableStyle = outStyle;
            }

            if (ExcelExporterSettings.AutoCenter)
            {
                CurrentExcelWorksheet.Cells[1, 1, CurrentExcelWorksheet.Dimension?.End.Row ?? 10, ExporterHeaderList.Count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            foreach (var exporterHeaderDto in ExporterHeaderList)
            {
                var exporterHeaderAttribute = exporterHeaderDto.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null && !exporterHeaderAttribute.IsIgnore)
                {
                    var colCell = CurrentExcelWorksheet.Cells[1, exporterHeaderDto.Index];
                    colCell.Style.Font.Bold = exporterHeaderAttribute.IsBold;

                    if (CurrentExcelTable != null)
                    {
                        var col = CurrentExcelTable.Columns[exporterHeaderDto.Index - 1];
                        col.Name = exporterHeaderDto.DisplayName;
                    }
                    else
                    {
                        colCell.Value = exporterHeaderDto.DisplayName;
                    }


                    var size = ExcelExporterSettings?.HeaderFontSize ?? exporterHeaderAttribute.FontSize;
                    if (size.HasValue)
                        colCell.Style.Font.Size = size.Value;
                }
            }
        }

        /// <summary>
        ///     添加样式
        /// </summary>
        protected virtual void AddStyle()
        {
            foreach (var exporterHeader in ExporterHeaderList)
            {
                var col = CurrentExcelWorksheet.Column(exporterHeader.Index);
                if (!string.IsNullOrWhiteSpace(exporterHeader.ExporterHeaderAttribute.Format))
                {
                    col.Style.Numberformat.Format = exporterHeader.ExporterHeaderAttribute.Format;

                }
                else
                {
                    //处理日期格式
                    switch (exporterHeader.CsTypeName)
                    {
                        case "DateTime":
                        case "DateTimeOffset":
                        //case "DateTime?":
                        case "Nullable<DateTime>":
                        case "Nullable<DateTimeOffset>":
                            //设置本地化时间格式
                            col.Style.Numberformat.Format = CultureInfo.CurrentUICulture.DateTimeFormat.FullDateTimePattern;
                            break;
                        default:
                            break;
                    }
                }

                if (!ExcelExporterSettings.AutoFitAllColumn && exporterHeader.ExporterHeaderAttribute.IsAutoFit)
                    col.AutoFit();

                if (exporterHeader.ExportImageFieldAttribute != null)
                {
                    col.Width = exporterHeader.ExportImageFieldAttribute.Width;
                }

                if (exporterHeader.ExporterHeaderAttribute.AutoCenterColumn)
                {
                    col.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

            }
        }
    }
}
