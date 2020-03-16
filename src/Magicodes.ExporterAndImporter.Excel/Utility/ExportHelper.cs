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

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    /// <summary>
    /// 导出辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExportHelper<T> where T : class
    {
        private ExporterAttribute _excelExporterAttribute;
        private ExcelWorksheet _excelWorksheet;
        private ExcelPackage _excelPackage;
        private List<ExporterHeaderInfo> _exporterHeaderList;
        /// <summary>
        /// 
        /// </summary>
        public ExportHelper()
        {
            if (typeof(DataTable).Equals(typeof(T)))
            {
                IsDynamicDatableExport = true;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="existExcelPackage"></param>

        public ExportHelper(ExcelPackage existExcelPackage)
        {
            if (typeof(DataTable).Equals(typeof(T)))
            {
                IsDynamicDatableExport = true;
            }
            if (existExcelPackage != null)
            {
                this._excelPackage = existExcelPackage;
            }
        }


        /// <summary>
        ///     导出设置
        /// </summary>
        public ExporterAttribute ExcelExporterSettings
        {
            get
            {
                if (_excelExporterAttribute == null)
                {
                    var type = typeof(T);
                    if (typeof(DataTable).Equals(type))
                    {
                        _excelExporterAttribute = new ExporterAttribute();
                    }
                    else
                        _excelExporterAttribute = type.GetAttribute<ExporterAttribute>(true) ?? new ExporterAttribute();

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
                    AddExcelWorksheet();
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
                var type = typeof(T);
                var objProperties = type.GetProperties();
                if (objProperties == null || objProperties.Length == 0)
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

            var tbStyle = TableStyles.Medium10;
            if (!ExcelExporterSettings.TableStyle.IsNullOrWhiteSpace())
                tbStyle = (TableStyles)Enum.Parse(typeof(TableStyles), ExcelExporterSettings.TableStyle);

            var er = excelRange.LoadFromCollection(dataItems, true, tbStyle);
            CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
        }

        /// <summary>
        ///     添加图片
        /// </summary>
        protected void AddPictures(int rowCount)
        {
            for (var i = 0; i < ExporterHeaderList.Count; i++)
            {
                if (ExporterHeaderList[i].ExportImageFieldAttribute != null)
                {
                    for (var j = 1; j <= rowCount; j++)
                    {
                        var cell = CurrentExcelWorksheet.Cells[j + 1, i + 1];
                        var url = cell.Text;
                        if (File.Exists(url) || url.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                cell.Value = string.Empty;
                                var bitmap = Extension.GetBitmapByUrl(url);
                                if (bitmap == null)
                                {
                                    cell.Value = ExporterHeaderList[i].ExportImageFieldAttribute.Alt;
                                }
                                else
                                {
                                    var pic = CurrentExcelWorksheet.Drawings.AddPicture(Guid.NewGuid().ToString(), bitmap);
                                    pic.SetPosition(j, ExporterHeaderList[i].ExportImageFieldAttribute.Height / 5, i - 1, 0);
                                    CurrentExcelWorksheet.Row(j + 1).Height = ExporterHeaderList[i].ExportImageFieldAttribute.Height;
                                    pic.SetSize(ExporterHeaderList[i].ExportImageFieldAttribute.Width * 7, ExporterHeaderList[i].ExportImageFieldAttribute.Height);
                                }
                               
                            }
                            catch (Exception)
                            {
                                cell.Value = ExporterHeaderList[i].ExportImageFieldAttribute.Alt;
                            }
                        }
                        else
                        {
                            cell.Value = ExporterHeaderList[i].ExportImageFieldAttribute.Alt;
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
            if (CurrentExcelTable == null)
            {
                var cols = ExporterHeaderList.Count;
                var range = CurrentExcelWorksheet.Cells[1, 1, 10, cols];
                CurrentExcelTable = CurrentExcelWorksheet.Tables.Add(range, "Table");
            }
            CurrentExcelTable.ShowHeader = true;
            foreach (var exporterHeaderDto in ExporterHeaderList)
            {
                var exporterHeaderAttribute = exporterHeaderDto.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null && !exporterHeaderAttribute.IsIgnore)
                {
                    var col = CurrentExcelTable.Columns[exporterHeaderDto.Index - 1];
                    col.Name = exporterHeaderDto.DisplayName;

                    var colCell = CurrentExcelWorksheet.Cells[1, exporterHeaderDto.Index];
                    colCell.Style.Font.Bold = exporterHeaderAttribute.IsBold;

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
                    col.Style.Numberformat.Format = exporterHeader.ExporterHeaderAttribute.Format;

                if (!ExcelExporterSettings.AutoFitAllColumn && exporterHeader.ExporterHeaderAttribute.IsAutoFit)
                    col.AutoFit();
                if (exporterHeader.ExportImageFieldAttribute != null)
                {
                    col.Width = exporterHeader.ExportImageFieldAttribute.Width;
                }
                //处理日期格式
                switch (exporterHeader.CsTypeName)
                {
                    case "DateTime":
                    case "DateTime?":
                        col.Style.Numberformat.Format = "yyyy-MM-dd";
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
