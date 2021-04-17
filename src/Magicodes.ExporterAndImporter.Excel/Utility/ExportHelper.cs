using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

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

            if (typeof(ExpandoObject).Equals(typeof(T)))
            {
                IsExpandoObjectType = true;
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
        /// <param name="sheetName"></param>

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
                                AutoFitMaxRows = exporterAttribute.AutoFitMaxRows,
                                ExporterHeaderFilter = exporterAttribute.ExporterHeaderFilter,
                                FontSize = exporterAttribute.FontSize,
                                HeaderFontSize = exporterAttribute.HeaderFontSize,
                                MaxRowNumberOnASheet = exporterAttribute.MaxRowNumberOnASheet,
                                Name = exporterAttribute.Name,
                                TableStyle = _excelExporterAttribute?.TableStyle ?? TableStyles.None,
                                AutoCenter = _excelExporterAttribute != null && _excelExporterAttribute.AutoCenter,
                                IsDisableAllFilter = exporterAttribute.IsDisableAllFilter
                            };
                        }
                        else
                            _excelExporterAttribute = new ExcelExporterAttribute();
                    }

                    #region 加载表头筛选器
                    ExporterHeaderFilter = GetFilter<IExporterHeaderFilter>(_excelExporterAttribute.ExporterHeaderFilter);
                    #endregion
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
                if ((_exporterHeaderList == null || _exporterHeaderList.Count == 0) && !IsDynamicDatableExport && !IsExpandoObjectType) throw new ArgumentException("请定义表头！");
                if (_exporterHeaderList.Count(t => t.ExporterHeaderAttribute.IsIgnore == false) == 0 && _exporterHeaderList.Count != 0) throw new ArgumentException("请勿忽略全部表头！");
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
        /// 是否为ExpandoObject类型，调用LoadFromDictionaries以支持动态导出
        /// </summary>
        protected bool IsExpandoObjectType { get; set; }

        /// <summary>
        /// 添加导出表头
        /// </summary>
        /// <param name="exporterHeaderInfos"></param>
        public virtual void AddExporterHeaderInfoList(List<ExporterHeaderInfo> exporterHeaderInfos)
        {
            _exporterHeaderList = exporterHeaderInfos;
        }
        /// <summary>
        /// 获得经过排序的属性
        /// </summary>
        protected virtual List<PropertyInfo> SortedProperties
        {
            get
            {
                var type = _type ?? typeof(T);
                var objProperties = type.GetProperties().OrderBy(p => p.GetAttribute<ExporterHeaderAttribute>()?.ColumnIndex ?? 10000).ToList();
                return objProperties;
            }
        }

        /// <summary>
        ///     获取头部定义
        /// </summary>
        /// <returns></returns>
        protected virtual void GetExporterHeaderInfoList(DataTable dt = null, ICollection<T> dataItems = null)
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
            else if (IsExpandoObjectType)
            {
                var items = dataItems as IEnumerable<IDictionary<string, object>>;
                var keys = new List<string>(items.First().Keys);
                for (int i = 0; i < keys.Count; i++)
                {
                    var item = new ExporterHeaderInfo
                    {
                        Index = i + 1,
                        PropertyName = keys[i],
                        ExporterHeaderAttribute = new ExporterHeaderAttribute(keys[i]),
                        CsTypeName = keys[i].GetType().GetCSharpTypeName(),
                        DisplayName = keys[i]
                    };
                    AddExportHeaderInfo(item);
                }
            }
            else if (!IsDynamicDatableExport)
            {
                //var type = _type ?? typeof(T);
                //#179 GetProperties方法不按特定顺序（如字母顺序或声明顺序）返回属性，因此此处支持按ColumnIndex排序返回
                //var objProperties = type.GetProperties().OrderBy(p => p.GetAttribute<ExporterHeaderAttribute>()?.ColumnIndex ?? 10000).ToArray();
                var objProperties = SortedProperties;
                if (objProperties.Count == 0)
                    return;
                for (var i = 0; i < objProperties.Count; i++)
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

                    var mappings = objProperties[i].GetAttributes<ValueMappingAttribute>().ToList();
                    foreach (var mappingAttribute in mappings.Where(mappingAttribute =>
                        !item.MappingValues.ContainsKey(mappingAttribute.Value)))
                        item.MappingValues.Add(mappingAttribute.Value, mappingAttribute.Text);

                    //如果存在自定义映射，则不会生成默认映射
                    if (!mappings.Any())
                    {
                        if (objProperties[i].PropertyType.IsEnum)
                        {
                            var propType = objProperties[i].PropertyType;
                            var isNullable = propType.IsNullable();
                            if (isNullable) propType = propType.GetNullableUnderlyingType();
                            var values = propType.GetEnumTextAndValues();

                            foreach (var value in values.Where(value => !item.MappingValues.ContainsKey(value.Key)))
                                item.MappingValues.Add(value.Value, value.Key);

                            if (isNullable)
                                if (!item.MappingValues.ContainsKey(string.Empty))
                                    item.MappingValues.Add(string.Empty, null);

                        }
                    }

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
            if (!IsExpandoObjectType)
            {
                var list = ParseData(dataItems);
                AddDataItems(list);
            }
            else
            {
                AddDataItems(dataItems);
            }
            // 为了传入dataItems，在这里提前调用一下
            if (_exporterHeaderList == null) GetExporterHeaderInfoList(null, dataItems);
            //仅当存在图片表头才渲染图片
            if (ExporterHeaderList.Any(p => p.ExportImageFieldAttribute != null))
            {
                AddPictures(dataItems.Count);
            }
            DisableAutoFitWhenDataRowsIsLarge(dataItems.Count);
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
        ///  <param name="beginRows"></param>
        public void CopyRows(int currentws, int tempws, bool isAppendHeaders, int beginRows = 2)
        {
            var tempWorksheet = _excelPackage.Workbook.Worksheets[tempws];
            var ws = _excelPackage.Workbook.Worksheets[currentws];

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
            SetSkipRows();
            return CurrentExcelPackage;
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        public ExcelPackage Export(DataTable dataItems)
        {
            if ((ExporterHeaderList == null || ExporterHeaderList.Count == 0) && IsDynamicDatableExport) GetExporterHeaderInfoList(dataItems);
            AddDataItems(dataItems);
            SetSkipRows();
            //TODO:动态导出暂不考虑支持图片导出，后续可以考虑通过约定实现
            //AddPictures(dataItems.Rows.Count);

            DisableAutoFitWhenDataRowsIsLarge(dataItems.Rows.Count);
            return AddHeaderAndStyles();
        }

        /// <summary>
        /// 在数据达到设置值时禁用自适应列
        /// </summary>
        /// <param name="count"></param>
        private void DisableAutoFitWhenDataRowsIsLarge(int count)
        {
            //如果已经设置了AutoFitMaxRows并且当前数据超过此设置，则关闭自适应列的配置
            if (ExcelExporterSettings.AutoFitMaxRows != 0 && count > ExcelExporterSettings.AutoFitMaxRows)
            {
                ExcelExporterSettings.AutoFitAllColumn = false;
                foreach (var item in ExporterHeaderList)
                {
                    if (item.ExporterHeaderAttribute != null)
                        item.ExporterHeaderAttribute.IsAutoFit = false;
                }
            }
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
        protected void AddDataItems(dynamic dataItems, ExcelRangeBase excelRange = null)
        {
            if (excelRange == null)
                excelRange = CurrentExcelWorksheet.Cells["A1"];

            if (dataItems == null || dataItems.Count == 0)
                return;

            if (ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable)
            {
                //如果TableStyle=None则Table不为null
                var er = excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
            }
            else
            {
                //if (IsExpandoObjectType)
                //    excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                //else
                excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
            }
        }

        /// <summary>
        ///     数据解析
        /// </summary>
        /// <param name="dataItems"></param>
        protected virtual DataTable ParseData(ICollection<T> dataItems)
        {
            var x = ExporterHeaderList.Count;
            var type = typeof(T);
            var properties = SortedProperties;
            DataTable dt = new DataTable();
            foreach (var propertyInfo in properties)
            {
                if (propertyInfo.PropertyType.IsEnum ||
                    propertyInfo.PropertyType == typeof(bool) ||
                    propertyInfo.PropertyType == typeof(bool?) ||
                    (propertyInfo.PropertyType.IsNullable() && propertyInfo.PropertyType.GetNullableUnderlyingType().IsEnum))
                {
                    dt.Columns.Add(propertyInfo.Name);
                }
                else if (propertyInfo.PropertyType.IsNullable())
                {
                    dt.Columns.Add(propertyInfo.Name,
                         propertyInfo.PropertyType.GetGenericArguments()[0]);
                }
                else
                {
                    dt.Columns.Add(propertyInfo.Name, propertyInfo.PropertyType);
                }
            }

            foreach (var dataItem in dataItems)
            {
                var dr = dt.NewRow();
                foreach (var propertyInfo in properties)
                {
                    var value = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem)?.ToString();
                    if (
                        propertyInfo.PropertyType.IsEnum ||
                        propertyInfo.PropertyType.GetNullableUnderlyingType() != null &&
                        propertyInfo.PropertyType.GetNullableUnderlyingType().IsEnum)
                    {
                        if (value != null)
                        {
                            var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);

                            if (col.MappingValues.Count > 0 && col.MappingValues.ContainsKey(value))
                            {
                                var mapValue = col.MappingValues.FirstOrDefault(f => f.Key == value);
                                dr[propertyInfo.Name] = mapValue.Value;
                            }
                            else
                            {
                                var enumDefinitionList = propertyInfo.PropertyType.GetEnumDefinitionList();
                                if (enumDefinitionList == null)
                                {
                                    enumDefinitionList = propertyInfo.PropertyType.GetNullableUnderlyingType()
                                        .GetEnumDefinitionList();
                                }

                                var tuple = enumDefinitionList.FirstOrDefault(f => f.Item1 == value);
                                if (tuple != null)
                                {
                                    if (!tuple.Item4.IsNullOrWhiteSpace())
                                    {
                                        dr[propertyInfo.Name] = tuple.Item4;
                                    }
                                    else
                                    {
                                        dr[propertyInfo.Name] = tuple.Item2;
                                    }
                                }
                                else
                                {
                                    dr[propertyInfo.Name] = value;
                                }
                            }
                        }
                    }
                    else if (propertyInfo.PropertyType.GetCSharpTypeName() == "Boolean")
                    {
                        var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        var val = Convert.ToBoolean(value);
                        if (col.MappingValues.Count > 0 && col.MappingValues.ContainsKey(val))
                        {
                            var mapValue = col.MappingValues.FirstOrDefault(f => f.Key == val);
                            dr[propertyInfo.Name] = mapValue.Value;
                        }
                        else
                        {
                            dr[propertyInfo.Name] = val;
                        }
                    }
                    else if (propertyInfo.PropertyType.GetCSharpTypeName() == "Nullable<Boolean>")
                    {
                        var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        var val = Convert.ToBoolean(value);
                        if (col.MappingValues.Count > 0 && col.MappingValues.ContainsKey(Convert.ToBoolean(val)))
                        {
                            var mapValue = col.MappingValues.FirstOrDefault(f => f.Key == val);
                            dr[propertyInfo.Name] = mapValue.Value;
                        }
                        else
                        {
                            dr[propertyInfo.Name] = val;
                        }
                    }
                    else if (propertyInfo.PropertyType.GetCSharpTypeName() == "Int32")
                    {
                        var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        var val = Convert.ToInt32(value);

                        if (col.MappingValues.Count > 0 && col.MappingValues.ContainsKey(Convert.ToInt32(val)))
                        {
                            var mapValue = col.MappingValues.FirstOrDefault(f => f.Key == val);
                            dr[propertyInfo.Name] = int.Parse(mapValue.Value);
                        }
                        else
                        {
                            dr[propertyInfo.Name] = val;
                        }
                    }
                    else if (propertyInfo.PropertyType.GetCSharpTypeName() == "DateTimeOffset")
                    {
                        dr[propertyInfo.Name]
                            = DateTimeOffset.Parse(
                                value);
                    }
                    else if (propertyInfo.PropertyType.GetCSharpTypeName() == "Nullable<DateTimeOffset>")
                    {
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            dr[propertyInfo.Name] = DBNull.Value;
                            break;
                        }

                        if (DateTimeOffset.TryParse(value, out var date))
                        {
                            dr[propertyInfo.Name] = date;
                            break;
                        }
                    }
                    else
                    {
                        if (value != null)
                        {
                            dr[propertyInfo.Name]
                                = value;
                        }
                        else
                        {
                            dr[propertyInfo.Name] = DBNull.Value;
                        }
                    }
                }

                dt.Rows.Add(dr);
            }
            return dt;
        }

        /// <summary>
        ///     添加图片
        /// </summary>
        protected void AddPictures(int rowCount)
        {
            int ignoreCount = 0;
            for (var colIndex = 0; colIndex < ExporterHeaderList.Count; colIndex++)
            {
                if (ExporterHeaderList[colIndex].ExporterHeaderAttribute.IsIgnore)
                {
                    ignoreCount++;
                }
                if (ExporterHeaderList[colIndex].ExportImageFieldAttribute != null)
                {
                    for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                    {
                        var cell = CurrentExcelWorksheet.Cells[rowIndex + 1, colIndex + 1];
                        var url = cell.Text;
                        if (File.Exists(url) || url.StartsWith("http", StringComparison.OrdinalIgnoreCase) || url.IsBase64StringValid())
                        {
                            try
                            {
                                cell.Value = string.Empty;
                                Bitmap bitmap;
                                if (url.IsBase64StringValid())
                                {
                                    bitmap = url.Base64StringToBitmap();
                                }
                                else
                                {
                                    bitmap = Extension.GetBitmapByUrl(url);
                                }

                                if (bitmap == null)
                                {
                                    cell.Value = ExporterHeaderList[colIndex].ExportImageFieldAttribute.Alt;
                                }
                                else
                                {
                                    var pic = CurrentExcelWorksheet.Drawings.AddPicture(Guid.NewGuid().ToString(), bitmap);
                                    AddImage((rowIndex + (ExcelExporterSettings.HeaderRowIndex > 1 ? ExcelExporterSettings.HeaderRowIndex : 0)),
                                        colIndex - ignoreCount, pic, ExporterHeaderList[colIndex].ExportImageFieldAttribute.YOffset, ExporterHeaderList[colIndex].ExportImageFieldAttribute.XOffset);

                                    //pic.SetPosition
                                    //    (rowIndex + (ExcelExporterSettings.HeaderRowIndex > 1 ? ExcelExporterSettings.HeaderRowIndex : 0),
                                    //    ExporterHeaderList[colIndex].ExportImageFieldAttribute.Height / 5, colIndex - ignoreCount, 0);

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

        private void AddImage(int rowIndex, int colIndex, ExcelPicture picture, int yOffset, int xOffset)
        {
            if (picture != null)
            {
                picture.From.Column = colIndex;
                picture.From.Row = rowIndex;
                //调整对齐
                picture.From.ColumnOff = Pixel2MTU(xOffset);
                picture.From.RowOff = Pixel2MTU(yOffset);
            }
        }

        private int Pixel2MTU(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
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

            var tbStyle = ExcelExporterSettings.TableStyle;
            //if (!ExcelExporterSettings.TableStyle.IsNullOrWhiteSpace())
            //    tbStyle = (TableStyles)Enum.Parse(typeof(TableStyles), ExcelExporterSettings.TableStyle);

            var er = excelRange.LoadFromDataTable(dataTable, true, tbStyle);
            CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
        }

        /// <summary>
        ///设置x行开始追加内容
        /// </summary>
        private void SetSkipRows()
        {
            if (ExcelExporterSettings.HeaderRowIndex > 1)
            {
                CurrentExcelWorksheet.InsertRow(1, ExcelExporterSettings.HeaderRowIndex);
            }
        }

        /// <summary>
        /// 删除忽略列
        /// </summary>  
        protected void DeleteIgnoreColumns()
        {
            var deletedCount = 0;
            foreach (var exporterHeaderDto in ExporterHeaderList.Where(p => p.ExporterHeaderAttribute != null && p.ExporterHeaderAttribute.IsIgnore))
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
            var isNoneStyle = ExcelExporterSettings.TableStyle == TableStyles.None;

            if (CurrentExcelTable == null && ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable && !isNoneStyle)
            {
                var cols = ExporterHeaderList.Count;
                var range = CurrentExcelWorksheet.Cells[1, 1, CurrentExcelWorksheet.Dimension?.End.Row ?? 10, cols];
                //https://github.com/dotnetcore/Magicodes.IE/issues/66
                CurrentExcelTable = CurrentExcelWorksheet.Tables.Add(range, $"Table{CurrentExcelWorksheet.Index}");
                CurrentExcelTable.ShowHeader = true;
                //Enum.TryParse(ExcelExporterSettings.TableStyle, out TableStyles outStyle);
                CurrentExcelTable.TableStyle = ExcelExporterSettings.TableStyle;
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
                if (exporterHeader.ExporterHeaderAttribute != null && !string.IsNullOrWhiteSpace(exporterHeader.ExporterHeaderAttribute.Format))
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

                if (!ExcelExporterSettings.AutoFitAllColumn && exporterHeader.ExporterHeaderAttribute != null && exporterHeader.ExporterHeaderAttribute.IsAutoFit)
                    col.AutoFit();

                if (exporterHeader.ExportImageFieldAttribute != null)
                {
                    col.Width = exporterHeader.ExportImageFieldAttribute.Width;
                }

                if (exporterHeader.ExporterHeaderAttribute != null)
                {
                    //设置单元格宽度
                    var width = exporterHeader.ExporterHeaderAttribute.Width;
                    if (width > 0)
                    {
                        col.Width = width;
                    }

                    if (exporterHeader.ExporterHeaderAttribute.AutoCenterColumn)
                    {
                        col.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }
            }
        }

        /// <summary>
        /// 获取筛选器
        /// </summary>
        /// <typeparam name="TFilter"></typeparam>
        /// <param name="filterType"></param>
        /// <returns></returns>
        private TFilter GetFilter<TFilter>(Type filterType = null) where TFilter : IFilter
        {
            return filterType.GetFilter<TFilter>(ExcelExporterSettings.IsDisableAllFilter);
        }
    }
}
