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
                    _excelExporterAttribute = type.GetAttribute<ExporterAttribute>(true) ?? new ExporterAttribute();
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
                if (_exporterHeaderList == null || _exporterHeaderList.Count == 0) throw new Exception("请定义表头！");
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
        ///     获取头部定义
        /// </summary>
        /// <returns></returns>
        protected virtual void GetExporterHeaderInfoList()
        {
            _exporterHeaderList = new List<ExporterHeaderInfo>();
            var objProperties = typeof(T).GetProperties();
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
                            ExporterHeaderAttribute[])?.FirstOrDefault() ?? new ExporterHeaderAttribute(objProperties[i].Name),
                    CsTypeName = objProperties[i].PropertyType.GetCSharpTypeName()
                };
                ////过滤忽略列
                //if (item.ExporterHeaderAttribute.IsIgnore)
                //{
                //    continue;
                //}
                //设置列显示名
                item.DisplayName = item.ExporterHeaderAttribute == null || item.ExporterHeaderAttribute.DisplayName == null || item.ExporterHeaderAttribute.DisplayName.IsNullOrWhiteSpace()
                                ? item.PropertyName
                                : item.ExporterHeaderAttribute.DisplayName;
                //TODO:执行列头筛选器
                _exporterHeaderList.Add(item);
            }
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <returns>文件</returns>
        public ExcelPackage Export(ICollection<T> dataItems)
        {
            AddDataItems(dataItems);

            //TODO:数据为空时仅导出表头
            AddHeader();

            AddStyle();

            DeleteIgnoreColumns();
            if (ExcelExporterSettings.AutoFitAllColumn)
            {
                CurrentExcelWorksheet.Cells[CurrentExcelWorksheet.Dimension.Address].AutoFitColumns();
            }

            //以便支持导出多Sheet
            SheetIndex++;
            return CurrentExcelPackage;
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

            var er = excelRange.LoadFromCollection(dataItems, false, tbStyle);
            CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
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

            var er = excelRange.LoadFromDataTable(dataTable, false, tbStyle);
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
            CurrentExcelTable.ShowHeader = true;
            foreach (var exporterHeaderDto in ExporterHeaderList)
            {
                var exporterHeaderAttribute = exporterHeaderDto.ExporterHeaderAttribute;
                if (exporterHeaderAttribute != null && !exporterHeaderAttribute.IsIgnore)
                {
                    var col = CurrentExcelTable.Columns[exporterHeaderDto.Index - 1];
                    col.Name = exporterHeaderDto.DisplayName;

                    CurrentExcelWorksheet.Cells[1, exporterHeaderDto.Index].Style.Font.Bold = exporterHeaderAttribute.IsBold;

                    var size = ExcelExporterSettings?.HeaderFontSize ?? exporterHeaderAttribute.FontSize;
                    if (size.HasValue)
                        CurrentExcelWorksheet.Cells[1, exporterHeaderDto.Index].Style.Font.Size = size.Value;
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
