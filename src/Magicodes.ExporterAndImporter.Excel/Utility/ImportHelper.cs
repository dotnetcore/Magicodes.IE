﻿// ======================================================================
// 
//           filename : ImportHelper.cs
//           description :
// 
//           created by 雪雁 at  2019-09-18 16:25
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Reflection;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    /// <summary>
    ///     导入辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ImportHelper<T> : IDisposable where T : class, new()
    {
        private ExcelImporterAttribute _excelImporterAttribute;

        /// <summary>
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="labelingFilePath"></param>
        public ImportHelper(string filePath = null, string labelingFilePath = null)
        {
            FilePath = filePath;
            LabelingFilePath = labelingFilePath;
        }

        /// <summary>
        ///     导入全局设置
        /// </summary>
        protected ExcelImporterAttribute ExcelImporterSettings
        {
            get
            {
                if (_excelImporterAttribute == null)
                {
                    var type = typeof(T);
                    _excelImporterAttribute = type.GetAttribute<ExcelImporterAttribute>(true);
                    if (_excelImporterAttribute != null) return _excelImporterAttribute;

                    var importerAttribute = type.GetAttribute<ImporterAttribute>(true);
                    if (importerAttribute != null)
                    {
                        _excelImporterAttribute = new ExcelImporterAttribute()
                        {
                            HeaderRowIndex = importerAttribute.HeaderRowIndex,
                            MaxCount = importerAttribute.MaxCount,
                            ImportResultFilter = importerAttribute.ImportResultFilter,
                            ImportHeaderFilter = importerAttribute.ImportHeaderFilter
                        };
                    }
                    else
                        _excelImporterAttribute = new ExcelImporterAttribute();

                    return _excelImporterAttribute;
                }

                return _excelImporterAttribute;
            }
            set => _excelImporterAttribute = value;
        }


        /// <summary>
        ///     导入文件路径
        /// </summary>
        protected string FilePath { get; set; }

        /// <summary>
        ///     标注文件路径
        /// </summary>
        public string LabelingFilePath { get; }

        /// <summary>
        ///     导入结果
        /// </summary>
        internal ImportResult<T> ImportResult { get; set; }

        /// <summary>
        ///     列头定义
        /// </summary>
        protected List<ImporterHeaderInfo> ImporterHeaderInfos { get; set; }

        /// <summary>
        /// </summary>
        public void Dispose()
        {
            ExcelImporterSettings = null;
            FilePath = null;
            ImporterHeaderInfos = null;
            ImportResult = null;
            GC.Collect();
        }

        /// <summary>
        ///     导入模型验证数据
        /// </summary>
        /// <returns></returns>
        public Task<ImportResult<T>> Import(string filePath = null)
        {
            if (!string.IsNullOrWhiteSpace(filePath)) FilePath = filePath;
            ImportResult = new ImportResult<T>();
            try
            {
                CheckImportFile(FilePath);
                using (Stream stream = new FileStream(FilePath, FileMode.Open))
                {
                    using (var excelPackage = new ExcelPackage(stream))
                    {
                        #region 检查模板
                        ParseTemplate(excelPackage);
                        if (ImportResult.HasError) return Task.FromResult(ImportResult);
                        #endregion
                        ParseData(excelPackage);
                        #region 数据验证
                        for (var i = 0; i < ImportResult.Data.Count; i++)
                        {
                            var isValid = ValidatorHelper.TryValidate(ImportResult.Data.ElementAt(i),
                                out var validationResults);
                            if (!isValid)
                            {
                                var rowIndex = ExcelImporterSettings.HeaderRowIndex + i + 1;
                                var dataRowError = GetDataRowErrorInfo(rowIndex);
                                foreach (var validationResult in validationResults)
                                {
                                    var key = validationResult.MemberNames.First();
                                    var column = ImporterHeaderInfos.FirstOrDefault(a => a.PropertyName == key);
                                    if (column != null) key = column.Header.Name;

                                    var value = validationResult.ErrorMessage;
                                    if (dataRowError.FieldErrors.ContainsKey(key))
                                        dataRowError.FieldErrors[key] += Environment.NewLine + value;
                                    else
                                        dataRowError.FieldErrors.Add(key, value);
                                }
                            }
                        }
                        RepeatDataCheck();
                        #endregion

                        //执行结果筛选器
                        if (ExcelImporterSettings.ImportResultFilter != null)
                        {
                            var filter = (IImportResultFilter)ExcelImporterSettings.ImportResultFilter.Assembly.CreateInstance(ExcelImporterSettings.ImportResultFilter.FullName);

                            if (filter == null)
                            {
                                throw new Exception("结果筛选器必须实现接口IImportResultFilter！");
                            }
                            ImportResult = filter.Filter(ImportResult);
                        }
                        //生成Excel错误标注
                        LabelingError(excelPackage);
                    }
                }
            }
            catch (Exception ex)
            {
                ImportResult.Exception = ex;
            }

            return Task.FromResult(ImportResult);
        }

        /// <summary>
        ///     获取当前行
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private DataRowErrorInfo GetDataRowErrorInfo(int rowIndex)
        {
            if (ImportResult.RowErrors == null) ImportResult.RowErrors = new List<DataRowErrorInfo>();

            var dataRowError = ImportResult.RowErrors.FirstOrDefault(p => p.RowIndex == rowIndex);
            if (dataRowError == null)
            {
                dataRowError = new DataRowErrorInfo
                {
                    RowIndex = rowIndex
                };
                ImportResult.RowErrors.Add(dataRowError);
            }

            return dataRowError;
        }

        /// <summary>
        ///     检查重复数据
        /// </summary>
        private void RepeatDataCheck()
        {
            //获取需要检查重复数据的列
            var notAllowRepeatCols = ImporterHeaderInfos.Where(p => p.Header.IsAllowRepeat == false).ToList();
            if (notAllowRepeatCols.Count == 0) return;

            var rowIndex = ExcelImporterSettings.HeaderRowIndex;
            var qDataList = ImportResult.Data.Select(p =>
            {
                rowIndex++;
                return new { RowIndex = rowIndex, RowData = p };
            }).ToList().AsQueryable();

            foreach (var notAllowRepeatCol in notAllowRepeatCols)
            {
                //查询指定列
                var qDataByProp = qDataList
                    .Select($"new(RowData.{notAllowRepeatCol.PropertyName} as Value, RowIndex)")
                    .OrderBy("Value").ToDynamicList();

                //重复行的行号
                var listRepeatRows = new List<int>();
                for (var i = 0; i < qDataByProp.Count; i++)
                {
                    //当前行值
                    var currentValue = qDataByProp[i].Value;
                    if (i == 0 || string.IsNullOrEmpty(currentValue?.ToString())) continue;

                    //上一行的值
                    var preValue = qDataByProp[i - 1].Value;
                    if (currentValue == preValue)
                    {
                        listRepeatRows.Add(qDataByProp[i - 1].RowIndex);
                        listRepeatRows.Add(qDataByProp[i].RowIndex);
                        //如果不是最后一行，则继续检测
                        if (i != qDataByProp.Count - 1) continue;
                    }

                    if (listRepeatRows.Count == 0) continue;

                    var errorIndexsStr = string.Join("，", listRepeatRows.Distinct());
                    foreach (var repeatRow in listRepeatRows.Distinct())
                    {
                        var dataRowError = GetDataRowErrorInfo(repeatRow);

                        var key = notAllowRepeatCol.Header?.Name ??
                                  notAllowRepeatCol.PropertyName;
                        var error = $"存在数据重复，请检查！所在行：{errorIndexsStr}。";
                        if (dataRowError.FieldErrors.ContainsKey(key))
                            dataRowError.FieldErrors[key] += Environment.NewLine + error;
                        else
                            dataRowError.FieldErrors.Add(key, error);
                    }

                    listRepeatRows.Clear();
                }
            }
        }

        /// <summary>
        ///     标注错误
        /// </summary>
        /// <param name="excelPackage"></param>
        internal virtual void LabelingError(ExcelPackage excelPackage)
        {
            //是否标注错误
            if (ExcelImporterSettings.IsLabelingError && ImportResult.HasError)
            {
                var worksheet = GetImportSheet(excelPackage);
                //先移除原先的错误标注
                if (worksheet.Comments != null && worksheet.Comments.Count > 0)
                {
                    int length = worksheet.Comments.Count;
                    for (int i = 0; i < length; i++)
                    {
                        worksheet.Comments.RemoveAt(0);
                    }
                }
                //TODO:标注模板错误
                //标注数据错误
                foreach (var item in ImportResult.RowErrors)
                    foreach (var field in item.FieldErrors)
                    {
                        var col = ImporterHeaderInfos.First(p => p.Header.Name == field.Key);
                        var cell = worksheet.Cells[item.RowIndex, col.Header.ColumnIndex];
                        cell.Style.Font.Color.SetColor(Color.Red);
                        cell.Style.Font.Bold = true;
                        //处理 如果存在了Comment后 再出现Comment附加的情况
                        if (cell.Comment == null)
                        {
                            cell.AddComment(string.Join(",", field.Value), col.Header.Author);
                        }
                        else
                        {
                            cell.Comment.Text = field.Value;
                            cell.Comment.Author = col.Header.Author;
                        }
                    }

                var ext = Path.GetExtension(FilePath);
                var filePath = string.IsNullOrWhiteSpace(LabelingFilePath) ? FilePath.Replace(ext, "_" + ext) : LabelingFilePath;
                excelPackage.SaveAs(new FileInfo(filePath));
            }
        }

        /// <summary>
        /// 标注业务错误
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="bussinessErrorDataList"></param>
        /// <param name="filePath">返回错误Excel路径</param>
        internal virtual void LabelingBussinessError(ExcelPackage excelPackage, List<DataRowErrorInfo> bussinessErrorDataList, out string filePath)
        {
            if (bussinessErrorDataList == null)
            {
                filePath = "";
                return;
            }
            this.ImportResult = new ImportResult<T>();
            ParseHeader();
            ParseTemplate(excelPackage);
            //执行结果筛选器
            if (ExcelImporterSettings.ImportResultFilter != null)
            {
                var filter = (IImportResultFilter)ExcelImporterSettings.ImportResultFilter.Assembly.CreateInstance(ExcelImporterSettings.ImportResultFilter.FullName);

                if (filter == null)
                {
                    throw new Exception("结果筛选器必须实现接口IImportResultFilter！");
                }
                ImportResult = filter.Filter(ImportResult);
            }
            //if (ExcelImporterSettings.IsLabelingError && ImportResult.HasError)
            //业务错误必须标注
            var worksheet = GetImportSheet(excelPackage);
            //标注数据错误
            foreach (var item in bussinessErrorDataList)
            {
                item.RowIndex += ExcelImporterSettings.HeaderRowIndex;
                foreach (var field in item.FieldErrors)
                {
                    var col = ImporterHeaderInfos.First(p => p.Header.Name == field.Key);
                    var cell = worksheet.Cells[item.RowIndex, col.Header.ColumnIndex];
                    cell.Style.Font.Color.SetColor(Color.Red);
                    cell.Style.Font.Bold = true;
                    if (cell.Comment == null)
                    {
                        cell.AddComment(string.Join(",", field.Value), col.Header.Author);
                    }
                    else
                    {
                        cell.Comment.Text = field.Value;

                    }
                }
            }

            var ext = Path.GetExtension(FilePath);
            filePath = string.IsNullOrWhiteSpace(LabelingFilePath) ? FilePath.Replace(ext, "_" + ext) : LabelingFilePath;
            excelPackage.SaveAs(new FileInfo(filePath));
        }

        /// <summary>
        ///     检查导入文件路劲
        /// </summary>
        /// <exception cref="ArgumentException">文件路径不能为空! - filePath</exception>
        private static void CheckImportFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("文件路径不能为空!", nameof(filePath));

            //TODO:在Docker容器中存在文件路径找不到问题，暂时先注释掉
            //if (!File.Exists(filePath))
            //{
            //    throw new ImportException("导入文件不存在!");
            //}
        }

        /// <summary>
        ///     解析模板
        /// </summary>
        /// <returns></returns>
        protected virtual void ParseTemplate(ExcelPackage excelPackage)
        {
            ImportResult.TemplateErrors = new List<TemplateErrorInfo>();
            //获取导入实体列定义
            ParseHeader();
            try
            {
                //根据名称获取Sheet，如果不存在则取第一个
                var worksheet = GetImportSheet(excelPackage);
                var excelHeaders = new Dictionary<string, int>();
                var endColumnCount = ExcelImporterSettings.EndColumnCount ?? worksheet.Dimension.End.Column;
                if (!string.IsNullOrWhiteSpace(ExcelImporterSettings.SheetDescription))
                {
                    ExcelImporterSettings.HeaderRowIndex++;
                }
                for (var columnIndex = 1; columnIndex <= endColumnCount; columnIndex++)
                {

                    var header = worksheet.Cells[ExcelImporterSettings.HeaderRowIndex, columnIndex].Text;

                    //如果未设置读取的截止列，则默认指定为出现空格，则读取截止
                    if (ExcelImporterSettings.EndColumnCount.HasValue &&
                        columnIndex > ExcelImporterSettings.EndColumnCount.Value ||
                        string.IsNullOrWhiteSpace(header))
                        break;

                    //不处理空表头
                    if (string.IsNullOrWhiteSpace(header)) continue;

                    if (excelHeaders.ContainsKey(header))
                        ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                        {
                            ErrorLevel = ErrorLevels.Error,
                            ColumnName = header,
                            RequireColumnName = null,
                            Message = "列头重复！"
                        });

                    excelHeaders.Add(header, columnIndex);
                }

                foreach (var item in ImporterHeaderInfos)
                    if (!excelHeaders.ContainsKey(item.Header.Name))
                    {
                        //仅验证必填字段
                        if (item.IsRequired)
                        {
                            ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                            {
                                ErrorLevel = ErrorLevels.Error,
                                ColumnName = null,
                                RequireColumnName = item.Header.Name,
                                Message = "当前导入模板中未找到此字段！"
                            });
                            continue;
                        }

                        ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                        {
                            ErrorLevel = ErrorLevels.Warning,
                            ColumnName = null,
                            RequireColumnName = item.Header.Name,
                            Message = "当前导入模板中未找到此字段！"
                        });
                    }
                    else
                    {
                        item.IsExist = true;
                        //设置列索引
                        if (item.Header.ColumnIndex == 0)
                            item.Header.ColumnIndex = excelHeaders[item.Header.Name];
                    }
            }
            catch (Exception ex)
            {
                ImportResult.TemplateErrors.Add(new TemplateErrorInfo
                {
                    ErrorLevel = ErrorLevels.Error,
                    ColumnName = null,
                    RequireColumnName = null,
                    Message = $"模板出现未知错误：{ex}"
                });
                throw new Exception($"模板出现未知错误：{ex.Message}", ex);
            }
        }

        /// <summary>
        ///     解析头部
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ArgumentException">导入实体没有定义ImporterHeader属性</exception>
        protected virtual bool ParseHeader()
        {
            ImporterHeaderInfos = new List<ImporterHeaderInfo>();
            var objProperties = typeof(T).GetProperties();
            if (objProperties.Length == 0) return false;

            foreach (var propertyInfo in objProperties)
            {
                //TODO:简化并重构
                //如果不设置，则自动使用默认定义
                var importerHeaderAttribute =
                    (propertyInfo.GetCustomAttributes(typeof(ImporterHeaderAttribute), true) as
                        ImporterHeaderAttribute[])?.FirstOrDefault() ?? new ImporterHeaderAttribute
                        {
                            Name = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                        };

                if (string.IsNullOrWhiteSpace(importerHeaderAttribute.Name))
                    importerHeaderAttribute.Name = propertyInfo.GetDisplayName() ?? propertyInfo.Name;

                //忽略字段处理
                if (importerHeaderAttribute.IsIgnore) continue;

                var colHeader = new ImporterHeaderInfo
                {
                    IsRequired = propertyInfo.IsRequired(),
                    PropertyName = propertyInfo.Name,
                    Header = importerHeaderAttribute
                };
                ImporterHeaderInfos.Add(colHeader);

                #region 处理值映射

                var mappings = propertyInfo.GetAttributes<ValueMappingAttribute>().ToList();
                foreach (var mappingAttribute in mappings.Where(mappingAttribute =>
                    !colHeader.MappingValues.ContainsKey(mappingAttribute.Text)))
                    colHeader.MappingValues.Add(mappingAttribute.Text, mappingAttribute.Value);

                //如果存在自定义映射，则不会生成默认映射
                if (mappings.Any()) continue;

                //为bool类型生成默认映射
                switch (propertyInfo.PropertyType.GetCSharpTypeName())
                {
                    case "Boolean":
                    case "Nullable<Boolean>":
                        {
                            if (!colHeader.MappingValues.ContainsKey("是")) colHeader.MappingValues.Add("是", true);
                            if (!colHeader.MappingValues.ContainsKey("否")) colHeader.MappingValues.Add("否", false);
                            break;
                        }
                }

                var type = propertyInfo.PropertyType;
                var isNullable = type.IsNullable();
                if (isNullable) type = type.GetNullableUnderlyingType();
                //为枚举类型生成默认映射
                if (type.IsEnum)
                {
                    var values = type.GetEnumTextAndValues();
                    foreach (var value in values.Where(value => !colHeader.MappingValues.ContainsKey(value.Key)))
                        colHeader.MappingValues.Add(value.Key, value.Value);

                    if (isNullable)
                        if (!colHeader.MappingValues.ContainsKey(string.Empty))
                            colHeader.MappingValues.Add(string.Empty, null);
                }

                #endregion
            }

            //列筛选器
            if (ExcelImporterSettings.ImportHeaderFilter != null)
            {
                var filter = (IImportHeaderFilter)ExcelImporterSettings.ImportHeaderFilter.Assembly.CreateInstance(ExcelImporterSettings.ImportHeaderFilter.FullName);

                if (filter == null)
                {
                    throw new Exception("导入列筛选器必须实现接口IImportHeaderFilter！");
                }
                ImporterHeaderInfos = filter.Filter(ImporterHeaderInfos);
            }

            return true;
        }

        /// <summary>
        ///     构建Excel模板
        /// </summary>
        protected virtual void StructureExcel(ExcelPackage excelPackage)
        {
            var worksheet =
                excelPackage.Workbook.Worksheets.Add(typeof(T).GetDisplayName() ??
                                                     ExcelImporterSettings.SheetName ?? "导入数据");
            if (!ParseHeader()) return;
            //设置列头
            //设置头部描述说明
            if (!string.IsNullOrWhiteSpace(ExcelImporterSettings.SheetDescription))
            {
                ExcelImporterSettings.HeaderRowIndex++;
                worksheet.Cells[1, 1, 1, ImporterHeaderInfos.Count].Merge = true;
                worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[1, 1].Value = ExcelImporterSettings.SheetDescription;
                worksheet.Row(1).Height = ExcelImporterSettings.DescriptionHeight;

            }

            for (var i = 0; i < ImporterHeaderInfos.Count; i++)
            {
                //忽略
                if (ImporterHeaderInfos[i].Header.IsIgnore) continue;

                worksheet.Cells[ExcelImporterSettings.HeaderRowIndex, i + 1].Value =
                    ImporterHeaderInfos[i].Header.Name;
                if (!string.IsNullOrWhiteSpace(ImporterHeaderInfos[i].Header.Description))
                    worksheet.Cells[ExcelImporterSettings.HeaderRowIndex, i + 1].AddComment(
                        ImporterHeaderInfos[i].Header.Description,
                        ImporterHeaderInfos[i].Header.Author);
                //如果必填，则列头标红
                if (ImporterHeaderInfos[i].IsRequired)
                    worksheet.Cells[ExcelImporterSettings.HeaderRowIndex, i + 1].Style.Font.Color.SetColor(Color.Red);

                if (ImporterHeaderInfos[i].MappingValues.Count > 0)
                {
                    //针对枚举类型和Bool类型添加数据约束
                    var range = ExcelCellBase.GetAddress(ExcelImporterSettings.HeaderRowIndex + 1, i + 1,
                        ExcelPackage.MaxRows, i + 1);
                    var dataValidations = worksheet.DataValidations.AddListValidation(range);
                    foreach (var mappingValue in ImporterHeaderInfos[i].MappingValues)
                        dataValidations.Formula.Values.Add(mappingValue.Key);
                }
            }

            worksheet.Cells.AutoFitColumns();
            worksheet.Cells.Style.WrapText = true;
            // worksheet.Cells[worksheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[worksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //绿色太丑了
            worksheet.Cells[worksheet.Dimension.Address].Style.Fill.BackgroundColor.SetColor(Color.White);


        }

        /// <summary>
        ///     解析数据
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ArgumentException">支持最大导入条数限制，默认50000</exception>
        protected virtual void ParseData(ExcelPackage excelPackage)
        {
            var worksheet = GetImportSheet(excelPackage);
            if (worksheet.Dimension.End.Row > ExcelImporterSettings.MaxCount + ExcelImporterSettings.HeaderRowIndex) throw new ArgumentException($"最大允许导入条数不能超过{ExcelImporterSettings.MaxCount}条！");

            ImportResult.Data = new List<T>();
            var propertyInfos = new List<PropertyInfo>(typeof(T).GetProperties());

            for (var rowIndex = ExcelImporterSettings.HeaderRowIndex + 1;
                rowIndex <= worksheet.Dimension.End.Row;
                rowIndex++)
            {
                //跳过空行
                if (worksheet.Cells[rowIndex, 1, rowIndex, worksheet.Dimension.End.Column].All(p => p.Text == string.Empty))
                {
                    continue;
                }
                {
                    var dataItem = new T();
                    foreach (var propertyInfo in propertyInfos.Where(p =>
                        ImporterHeaderInfos.Any(p1 => p1.PropertyName == p.Name && p1.IsExist)))
                    {
                        var col = ImporterHeaderInfos.First(a => a.PropertyName == propertyInfo.Name);

                        var cell = worksheet.Cells[rowIndex, col.Header.ColumnIndex];
                        try
                        {
                            var cellValue = cell.Value?.ToString();
                            if (!cellValue.IsNullOrWhiteSpace())
                                if (col.MappingValues.Count > 0 && col.MappingValues.ContainsKey(cellValue))
                                {
                                    //TODO:进一步缓存并优化
                                    var isEnum = propertyInfo.PropertyType.IsEnum;
                                    var isNullable = propertyInfo.PropertyType.IsNullable();
                                    var type = propertyInfo.PropertyType;
                                    if (isNullable)
                                    {
                                        type = propertyInfo.PropertyType.GetNullableUnderlyingType();
                                        isEnum = type.IsEnum;
                                    }

                                    var value = col.MappingValues[cellValue];

                                    if (isEnum && isNullable && (value is int || value is short) &&
                                        Enum.IsDefined(type, value))
                                        propertyInfo.SetValue(dataItem,
                                            value == null ? null : Enum.ToObject(type, value));
                                    //propertyInfo.SetValue(dataItem,
                                    //    value == null ? null : Convert.ChangeType(value, type));
                                    else
                                        propertyInfo.SetValue(dataItem,
                                            value);
                                    continue;
                                }

                            //
                            if (propertyInfo.PropertyType.IsEnum)
                            {
                                AddRowDataError(rowIndex, col, $"值 {cellValue} 不存在模板下拉选项中");
                                continue;
                            }


                            switch (propertyInfo.PropertyType.GetCSharpTypeName())
                            {
                                case "Boolean":
                                    propertyInfo.SetValue(dataItem, false);
                                    //AddRowDataError(rowIndex, col, $"值 {cellValue} 不存在模板下拉选项中");
                                    break;
                                case "Nullable<Boolean>":
                                    if (string.IsNullOrWhiteSpace(cellValue))
                                        propertyInfo.SetValue(dataItem, null);
                                    else
                                        AddRowDataError(rowIndex, col, $"值 {cellValue} 不合法！");
                                    break;
                                case "String":
                                    //TODO:进一步优化
                                    //移除所有的空格，包括中间的空格
                                    if (col.Header.FixAllSpace)
                                        propertyInfo.SetValue(dataItem, cellValue?.Replace(" ", string.Empty));
                                    else if (col.Header.AutoTrim)
                                        propertyInfo.SetValue(dataItem, cellValue?.Trim());
                                    else
                                        propertyInfo.SetValue(dataItem, cellValue);

                                    break;
                                //long
                                case "Int64":
                                    {
                                        if (!long.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Int64>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }

                                        if (!long.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Int32":
                                    {
                                        if (!int.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Int32>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }

                                        if (!int.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Int16":
                                    {
                                        if (!short.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Int16>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }

                                        if (!short.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Decimal":
                                    {
                                        if (!decimal.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Decimal>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }

                                        if (!decimal.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Double":
                                    {
                                        if (!double.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Double>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }

                                        if (!double.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                //case "float":
                                case "Single":
                                    {
                                        if (!float.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Single>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }

                                        if (!float.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "DateTime":
                                    {
                                        if (!DateTime.TryParse(cellValue, out var date))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的日期时间格式！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, date);
                                    }
                                    break;
                                case "Nullable<DateTime>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }

                                        if (!DateTime.TryParse(cellValue, out var date))
                                        {
                                            AddRowDataError(rowIndex, col, $"值 {cellValue} 无效，请填写正确的日期时间格式！");
                                            break;
                                        }

                                        propertyInfo.SetValue(dataItem, date);
                                    }
                                    break;
                                default:
                                    propertyInfo.SetValue(dataItem, cell.Value);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            AddRowDataError(rowIndex, col, ex.Message);
                        }
                    }

                    ImportResult.Data.Add(dataItem);
                }
            }
        }

        /// <summary>
        ///     获取导入的Sheet
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <returns></returns>
        protected virtual ExcelWorksheet GetImportSheet(ExcelPackage excelPackage)
        {
#if NET461
            return excelPackage.Workbook.Worksheets[typeof(T).GetDisplayName()] ??
                   excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetName] ??
                   excelPackage.Workbook.Worksheets[1];
#else
            return excelPackage.Workbook.Worksheets[typeof(T).GetDisplayName()] ??
                   excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetName] ??
                   excelPackage.Workbook.Worksheets[0];
#endif
        }

        /// <summary>
        ///     添加数据行错误
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="importerHeaderInfo"></param>
        /// <param name="errorMessage"></param>
        protected virtual void AddRowDataError(int rowIndex, ImporterHeaderInfo importerHeaderInfo,
            string errorMessage = "数据格式无效！")
        {
            var rowError = GetDataRowErrorInfo(rowIndex);
            rowError.FieldErrors.Add(importerHeaderInfo.Header.Name, errorMessage);
        }

        /// <summary>
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private static bool GetBooleanValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;

            switch (value.ToLower())
            {
                case "1":
                case "是":
                case "yes":
                case "true":
                    return true;
                case "0":
                case "否":
                case "no":
                case "false":
                default:
                    return false;
            }
        }

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <returns>二进制字节</returns>
        public Task<byte[]> GenerateTemplateByte()
        {
            using (var excelPackage = new ExcelPackage())
            {
                StructureExcel(excelPackage);
                return Task.FromResult(excelPackage.GetAsByteArray());
            }
        }

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">文件名必须填写! - fileName</exception>
        public Task<ExportFileInfo> GenerateTemplate(string fileName = null)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("文件名必须填写!", fileName);

            var fileInfo =
                ExcelHelper.CreateExcelPackage(fileName, excelPackage => { StructureExcel(excelPackage); });
            return Task.FromResult(fileInfo);
        }

        /// <summary>
        /// 将存在的错误数据通过导入模板返回,并且标识业务错误原因
        /// </summary>
        /// <param name="bussinessErrorDataList">错误的业务数据</param>
        /// <param name="msg">成功:错误数据返回路径,失败 返回错误原因</param>
        /// <returns></returns>
        public bool OutputBussinessErrorData(List<DataRowErrorInfo> bussinessErrorDataList, out string msg)
        {
            try
            {
                using (Stream stream = new FileStream(FilePath, FileMode.Open))
                {
                    using (var excelPackage = new ExcelPackage(stream))
                    {
                        //生成Excel错误标注
                        LabelingBussinessError(excelPackage, bussinessErrorDataList, out msg);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                return false;
            }
        }
    }
}