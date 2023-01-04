using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.IE.Core;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using Magicodes.IE.Excel.Images;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    /// <summary>
    /// 导出辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public partial class ExportHelper<T> where T : class, new()
    {
        private ExcelExporterAttribute _excelExporterAttribute;
        private ExcelWorksheet _excelWorksheet;
        private ExcelPackage _excelPackage;
        private IList<ExporterHeaderInfo> _exporterHeaderList;
        private Type _type;
        private string _sheetName;

        #region 构造函数
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

            if (typeof(ExpandoObject).Equals(typeof(T)))
            {
                IsExpandoObjectType = true;
            }

            if (existExcelPackage != null)
            {
                this._excelPackage = existExcelPackage;
            }

            _sheetName = sheetName;
        }
        #endregion

        #region 属性
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
                                ExporterHeadersFilter = exporterAttribute.ExporterHeadersFilter,
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

                    ExporterHeaderFilter =
                        GetFilter<IExporterHeaderFilter>(_excelExporterAttribute.ExporterHeaderFilter);
                    ExporterHeadersFilter =
                        GetFilter<IExporterHeadersFilter>(_excelExporterAttribute.ExporterHeadersFilter);

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
        protected IList<ExporterHeaderInfo> ExporterHeaderList
        {
            get
            {
                if (_exporterHeaderList == null)
                {
                    GetExporterHeaderInfoList();
                }

                if ((_exporterHeaderList == null || _exporterHeaderList.Count == 0) && !IsDynamicDatableExport &&
                    !IsExpandoObjectType) throw new ArgumentException(Resource.DefineTheHeader);
                if (_exporterHeaderList.Count(t => t.ExporterHeaderAttribute.IsIgnore == false) == 0 &&
                    _exporterHeaderList.Count != 0) throw new ArgumentException(Resource.DoNotIgnoreAllTheHeader);
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
        /// 表头（集合）筛选器
        /// </summary>
        protected IExporterHeadersFilter ExporterHeadersFilter { get; set; }

        /// <summary>
        /// 是否为动态DataTable导出
        /// </summary>
        protected bool IsDynamicDatableExport { get; set; }

        /// <summary>
        /// 是否为ExpandoObject类型，调用LoadFromDictionaries以支持动态导出
        /// </summary>
        protected bool IsExpandoObjectType { get; set; }

        #endregion

        #region 导出
        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <returns>文件</returns>
        public virtual ExcelPackage Export(ICollection<T> dataItems)
        {
            if (_exporterHeaderList == null) GetExporterHeaderInfoList(null, dataItems);

            //TODO:先获取列头
            if (!IsExpandoObjectType)
            {
                AddDataItems(ParseData(dataItems));
            }
            else
            {
                AddDataItems(dataItems);
            }

            //仅当存在图片表头才渲染图片
            if (ExporterHeaderList.Any(p => p.ExportImageFieldAttribute != null))
            {
                AddPictures(dataItems.Count);
            }

            DisableAutoFitWhenDataRowsIsLarge(dataItems.Count);
            return AddHeaderAndStyles();
        }

        /// <summary>
        ///     导出Excel
        /// </summary>
        public ExcelPackage Export(DataTable dataItems)
        {
            if ((ExporterHeaderList == null || ExporterHeaderList.Count == 0) && IsDynamicDatableExport)
            {
                GetExporterHeaderInfoList(dataItems);
            }

            AddDataItems(dataItems);
            SetSkipRows();
            //TODO:动态导出暂不考虑支持图片导出，后续可以考虑通过约定实现
            //AddPictures(dataItems.Rows.Count);

            DisableAutoFitWhenDataRowsIsLarge(dataItems.Rows.Count);
            return AddHeaderAndStyles();
        }
        #endregion

        #region 表头相关操作
        /// <summary>
        /// 添加导出表头
        /// </summary>
        /// <param name="exporterHeaderInfos"></param>
        public virtual void AddExporterHeaderInfoList(List<ExporterHeaderInfo> exporterHeaderInfos)
        {
            _exporterHeaderList = exporterHeaderInfos;
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
                //执行列头（集合）筛选器
                if (ExporterHeadersFilter != null)
                {
                    _exporterHeaderList = ExporterHeadersFilter.Filter(_exporterHeaderList);
                }
                ReorderHeaders();
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
                var type = _type ?? typeof(T);
                var objProperties = type.GetProperties().ToList();
                if (objProperties.Count == 0)
                    return;
                for (var i = 0; i < objProperties.Count; i++)
                {
                    var exporterHeaderAttribute = (objProperties[i].GetCustomAttributes(typeof(ExporterHeaderAttribute), true) as
                                ExporterHeaderAttribute[])?.FirstOrDefault() ??
                            new ExporterHeaderAttribute(objProperties[i].GetDisplayName() ?? objProperties[i].Name);
                    var item = new ExporterHeaderInfo
                    {
                        Index = i + 1,
                        PropertyName = objProperties[i].Name,
                        ExporterHeaderAttribute = exporterHeaderAttribute,
                        CsTypeName = objProperties[i].PropertyType.GetCSharpTypeName(),
                        ExportImageFieldAttribute = objProperties[i].GetAttribute<ExportImageFieldAttribute>(true),
                        PropertyInfo = objProperties[i],
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
                        (objProperties[i].GetAttribute<IEIgnoreAttribute>(true) == null)
                            ? item.ExporterHeaderAttribute.IsIgnore
                            : objProperties[i].GetAttribute<IEIgnoreAttribute>(true).IsExportIgnore;

                    var itemMappingValues = item.MappingValues;
                    objProperties[i].ValueMapping(ref itemMappingValues);
                    //var mappings = objProperties[i].GetAttributes<ValueMappingAttribute>().ToList();
                    //foreach (var mappingAttribute in mappings.Where(mappingAttribute =>
                    //    !item.MappingValues.ContainsKey(mappingAttribute.Value)))
                    //    item.MappingValues.Add(mappingAttribute.Value, mappingAttribute.Text);

                    ////如果存在自定义映射，则不会生成默认映射
                    //if (!mappings.Any())
                    //{
                    //    if (objProperties[i].PropertyType.IsEnum)
                    //    {
                    //        var propType = objProperties[i].PropertyType;
                    //        var isNullable = propType.IsNullable();
                    //        if (isNullable) propType = propType.GetNullableUnderlyingType();
                    //        var values = propType.GetEnumTextAndValues();

                    //        foreach (var value in values.Where(value => !item.MappingValues.ContainsKey(value.Key)))
                    //            item.MappingValues.Add(value.Value, value.Key);

                    //        if (isNullable)
                    //            if (!item.MappingValues.ContainsKey(string.Empty))
                    //                item.MappingValues.Add(string.Empty, null);

                    //    }
                    //}

                    AddExportHeaderInfo(item);
                }
            }

            //执行列头（集合）筛选器
            if (ExporterHeadersFilter != null)
            {
                _exporterHeaderList = ExporterHeadersFilter.Filter(_exporterHeaderList);
            }

            ReorderHeaders();
        }

        /// <summary>
        /// 表头重新排序
        /// </summary>
        private void ReorderHeaders()
        {
            if (_exporterHeaderList.Any(p => p.ExporterHeaderAttribute?.ColumnIndex >= 0 && p.ExporterHeaderAttribute?.ColumnIndex < 10000))
            {
                //列索引从小到大先进行排序和处理，以防插入前面导致索引不对
                _exporterHeaderList = _exporterHeaderList.OrderBy(p => p.ExporterHeaderAttribute?.ColumnIndex ?? 10000).ToList();
                //修改列位置
                var maxIndex = _exporterHeaderList.Count - 1;
                var todoOrderHeaders = new List<ExporterHeaderInfo>();
                //将存在指定列索引的列筛出来，并进行索引修正（防止误赋值）
                foreach (var item in _exporterHeaderList.Where(p => p.ExporterHeaderAttribute?.ColumnIndex >= 0 && p.ExporterHeaderAttribute?.ColumnIndex < 10000))
                {
                    //10000及以上属于无效索引
                    var index = item.ExporterHeaderAttribute.ColumnIndex;
                    //如果索引小于0，则设置为0
                    if (index < 0)
                        index = 0;
                    //如果索引设置超出当前列数，则插入最后一列
                    if (index > maxIndex)
                        index = maxIndex;
                    item.ExporterHeaderAttribute.ColumnIndex = index;
                    todoOrderHeaders.Add(item);
                }
                //移动位置
                foreach (var item in todoOrderHeaders)
                {
                    _exporterHeaderList.Remove(item);
                    _exporterHeaderList.Insert(item.ExporterHeaderAttribute.ColumnIndex, item);
                }
                //重新编排序号
                for (int i = 0; i < _exporterHeaderList.Count; i++)
                {
                    var item = _exporterHeaderList[i];
                    item.Index = i + 1;
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
        #endregion


        #region Sheet、Row相关操作

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
        /// 添加Sheet
        /// 支持同一个数据拆成多个Sheet
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public ExcelWorksheet AddExcelWorksheet(string name = null)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                name = ExcelExporterSettings?.Name ?? Resource.ExportResult;
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
        #endregion



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

        #region 数据处理
        /// <summary>
        ///     添加导出数据
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="excelRange"></param>
        protected void AddDataItems(IEnumerable<ExpandoObject> dataItems, ExcelRangeBase excelRange = null)
        {
            if (excelRange == null)
                excelRange = CurrentExcelWorksheet.Cells["A1"];

            if (dataItems == null || !dataItems.Any())
            {
                return;
            }

            if (ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable)
            {
                if (IsExpandoObjectType)
                    excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                else
                {
                    //如果TableStyle=None则Table不为null
                    var er = excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                    CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
                }
            }
            else
            {
                //if (IsExpandoObjectType)
                //  excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                //else
                excelRange.LoadFromDictionaries(dataItems, true, ExcelExporterSettings.TableStyle);
                //CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
            }
        }

        /// <summary>
        ///     添加导出数据
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="excelRange"></param>
        protected void AddDataItems(IEnumerable<T> dataItems, ExcelRangeBase excelRange = null)
        {
            if (excelRange == null)
                excelRange = CurrentExcelWorksheet.Cells["A1"];

            if (dataItems == null || !dataItems.Any())
            {
                return;
            }

            if (ExcelExporterSettings.ExcelOutputType == ExcelOutputTypes.DataTable)
            {
                if (IsExpandoObjectType)
                    excelRange.LoadFromDictionaries((IEnumerable<IDictionary<string, object>>)dataItems, true, ExcelExporterSettings.TableStyle);
                else
                {
                    //如果TableStyle=None则Table不为null
                    var er = excelRange.LoadFromCollection(dataItems, true, ExcelExporterSettings.TableStyle);
                    CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
                }
            }
            else
            {
                excelRange.LoadFromCollection(dataItems, true, ExcelExporterSettings.TableStyle);
            }
        }

        protected virtual IEnumerable<ExpandoObject> ParseData(ICollection<T> dataItems)
        {
            var type = typeof(T);
            var properties = ExporterHeaderList
                    ?.OrderBy(p => p.Index)
                    .Select(p => p.PropertyInfo)
                    ?.ToList();
            //IEnumerable<ExpandoObject> list = new List<ExpandoObject>();

            foreach (var dataItem in dataItems)
            {
                dynamic obj = new ExpandoObject();
                foreach (var propertyInfo in properties)
                {
                    bool s = propertyInfo.PropertyType.GetNullableUnderlyingType()?.IsEnum ?? false;
                    if (propertyInfo.PropertyType.IsEnum
                        || (propertyInfo.PropertyType.GetNullableUnderlyingType()?.IsEnum ?? false))
                    {
                        //var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                        //var value = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem)?.ToString();

                        //if (col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(value ?? string.Empty))
                        //{
                        //    var mapValue = col.MappingValues.FirstOrDefault(f => f.Key == value);
                        //    ((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Value;
                        //}
                        //else
                        //{

                        //}
                        if (
                            propertyInfo.PropertyType.IsEnum ||
                            propertyInfo.PropertyType.GetNullableUnderlyingType() != null &&
                            propertyInfo.PropertyType.GetNullableUnderlyingType().IsEnum)
                        {
                            {
                                var value = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem)?.GetHashCode();
                                {
                                    var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);

                                    if (col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(value))
                                    {
                                        var mapValue = col.MappingValues.FirstOrDefault(f => f.Value == value);
                                        //dr[propertyInfo.Name] = mapValue.Value;
                                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Key;
                                    }
                                    else
                                    {
                                        var enumDefinitionList = propertyInfo.PropertyType.GetEnumDefinitionList();
                                        if (enumDefinitionList == null)
                                        {
                                            enumDefinitionList = propertyInfo.PropertyType.GetNullableUnderlyingType()
                                                .GetEnumDefinitionList();
                                        }

                                        var tuple = enumDefinitionList.FirstOrDefault(f => f.Item1 == value.ToString());
                                        if (tuple != null)
                                        {
                                            if (!tuple.Item4.IsNullOrWhiteSpace())
                                            {
                                                //dr[propertyInfo.Name] = tuple.Item4;
                                                ((IDictionary<string, object>)obj)[propertyInfo.Name] = tuple.Item4;
                                            }
                                            else
                                            {
                                                ((IDictionary<string, object>)obj)[propertyInfo.Name] = tuple.Item2;
                                                // dr[propertyInfo.Name] = tuple.Item2;
                                            }
                                        }
                                        else
                                        {
                                            ((IDictionary<string, object>)obj)[propertyInfo.Name] = value;
                                            //dr[propertyInfo.Name] = value;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        var cSharpTypeName = propertyInfo.PropertyType.GetCSharpTypeName();
                        switch (cSharpTypeName)
                        {
                            case "Boolean":
                                {
                                    var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                                    var val = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem).ToString();
                                    bool value = Convert.ToBoolean(val);
                                    if (col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(value))
                                    {
                                        var mapValue = col.MappingValues.FirstOrDefault(f => f.Value.ToString() == value.ToString());
                                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Key;
                                    }
                                    else
                                    {
                                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = value;
                                    }
                                }
                                break;
                            case "Nullable<Boolean>":
                                {
                                    var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                                    var objValue = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem);
                                    if (objValue == null)
                                    {
                                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = null;
                                    }
                                    else
                                    {
                                        var boolValue = (bool?)(objValue);
                                        if (boolValue.HasValue && col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(boolValue))
                                        {
                                            var mapValue = col.MappingValues.FirstOrDefault(f => f.Value == boolValue);
                                            ((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Key;
                                        }
                                        else
                                        {
                                            ((IDictionary<string, object>)obj)[propertyInfo.Name] = objValue;
                                        }

                                    }
                                }
                                break;
                            case "DateTimeOffset":
                                {
                                    var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                                    var val = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem).ToString();
                                    DateTime value = DateTimeOffset.Parse(val).ConvertToDateTime();
                                    if (col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(value))
                                    {
                                        var mapValue = col.MappingValues.FirstOrDefault(f => f.Value.ToString() == value.ToString());
                                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Key;
                                    }
                                    else
                                    {
                                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = value;
                                    }
                                }
                                break;
                            case "Nullable<DateTimeOffset>":
                                {
                                    var col = ExporterHeaderList.First(a => a.PropertyName == propertyInfo.Name);
                                    var objValue = type.GetProperty(propertyInfo.Name)?.GetValue(dataItem);
                                    if (objValue == null)
                                    {
                                        ((IDictionary<string, object>)obj)[propertyInfo.Name] = null;
                                    }
                                    else
                                    {
                                        var dateValue = ((DateTimeOffset?)(objValue)).ConvertToDateTime();
                                        if (dateValue.HasValue && col.MappingValues.Count > 0 && col.MappingValues.ContainsValue(dateValue))
                                        {
                                            var mapValue = col.MappingValues.FirstOrDefault(f => f.Value == dateValue);
                                            ((IDictionary<string, object>)obj)[propertyInfo.Name] = mapValue.Key;
                                        }
                                        else
                                        {
                                            ((IDictionary<string, object>)obj)[propertyInfo.Name] = dateValue;
                                        }

                                    }
                                }
                                break;
                            default:
                                ((IDictionary<string, object>)obj)[propertyInfo.Name] = type.GetProperty(propertyInfo.Name)
                            ?.GetValue(dataItem);
                                break;
                        }
                    }
                }

                yield return obj;
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

            var tbStyle = ExcelExporterSettings.TableStyle;
            //if (!ExcelExporterSettings.TableStyle.IsNullOrWhiteSpace())
            //    tbStyle = (TableStyles)Enum.Parse(typeof(TableStyles), ExcelExporterSettings.TableStyle);

            var er = excelRange.LoadFromDataTable(dataTable, true, tbStyle);
            CurrentExcelTable = CurrentExcelWorksheet.Tables.GetFromRange(er);
        }
        #endregion

        #region 图片处理
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
                                Image image;
                                IImageFormat format;
                                if (url.IsBase64StringValid())
                                {
                                    image = url.Base64StringToImage(out format);
                                }
                                else
                                {
                                    image = url.GetImageByUrl(out format);
                                }

                                if (image == null)
                                {
                                    cell.Value = ExporterHeaderList[colIndex].ExportImageFieldAttribute.Alt;
                                }
                                else
                                {
                                    using (ExcelPicture pic = CurrentExcelWorksheet.Drawings.AddPicture(Guid.NewGuid().ToString(), image, format))
                                    {
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

        internal static void AddImage(int rowIndex, int colIndex, ExcelPicture picture, int yOffset, int xOffset)
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

        internal static int Pixel2MTU(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
        }
        #endregion




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

                    if (exporterHeader.ExporterHeaderAttribute.WrapText)
                    {
                        col.Style.WrapText = exporterHeader.ExporterHeaderAttribute.WrapText;
                    }
                    col.Hidden = exporterHeader.ExporterHeaderAttribute.Hidden;

                    if (exporterHeader.ExporterHeaderAttribute.FontColor != null)
                    {
                        col.Style.Font.Color.SetColor(exporterHeader.ExporterHeaderAttribute.FontColor.Value);
                    }
                }
                if (ExcelExporterSettings.FontSize != 0)
                {
                    col.Style.Font.Size = ExcelExporterSettings.FontSize;
                }

                var headerSize = ExcelExporterSettings.HeaderFontSize;
                if (headerSize == 0 && exporterHeader.ExporterHeaderAttribute != null)
                    headerSize = exporterHeader.ExporterHeaderAttribute.FontSize;

                if (headerSize != 0)
                {
                    var headerCell = CurrentExcelWorksheet.Cells[1, exporterHeader.Index];
                    headerCell.Style.Font.Size = headerSize;
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
