using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Reflection;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{

    /// <summary>
    ///     多Sheet导入辅助类
    /// </summary>
    public class ImportMultipleSheetHelper : IDisposable
    {
        private ExcelImporterAttribute _excelImporterAttribute;

        private Type _importDataType;

        private Stream _excelStream;

        private ExcelPackage _excelPackage;

        private List<PropertyInfo> _sheetPropertyList;


        private ImportMultipleSheetHelper()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="stream"></param>
        public ImportMultipleSheetHelper(Stream stream)
        {
            _excelStream = stream;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        public ImportMultipleSheetHelper(string filePath)
        {
            FilePath = filePath;
            CheckImportFile(FilePath);
            _excelStream = new FileStream(FilePath, FileMode.Open);
        }

        /// <summary>
        /// Sheet属性信息列表
        /// </summary>
        /// <param name="sheetPropertyList"></param>
        public ImportMultipleSheetHelper(List<PropertyInfo> sheetPropertyList)
        {
            _sheetPropertyList = sheetPropertyList;
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
                    var type = _importDataType;
                    _excelImporterAttribute = type.GetAttribute<ExcelImporterAttribute>(true);
                    if (_excelImporterAttribute != null) return _excelImporterAttribute;

                    var importerAttribute = type.GetAttribute<ImporterAttribute>(true);
                    if (importerAttribute != null)
                    {
                        _excelImporterAttribute = new ExcelImporterAttribute()
                        {
                            HeaderRowIndex = importerAttribute.HeaderRowIndex
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
        ///     导入结果
        /// </summary>
        protected ImportResult<object> ImportResult { get; set; }

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
            _excelPackage?.Dispose();
            _excelStream?.Close();
            _excelStream?.Dispose();
            GC.Collect();
        }

        /// <summary>
        ///     导入模型验证数据
        /// </summary>
        /// <returns></returns>
        public Task<ImportResult<object>> Import(string sheetName, int sheetIndex, Type importDataType, bool isSaveLabelingError = true)
        {
            _importDataType = importDataType;
            ImportResult = new ImportResult<object>();
            ExcelImporterSettings.SheetName = sheetName;
            ExcelImporterSettings.SheetIndex = sheetIndex;
            try
            {
                if (_excelPackage == null) _excelPackage = new ExcelPackage(_excelStream);
                {
                    #region 检查模板

                    ParseTemplate(_excelPackage);
                    if (ImportResult.HasError) return Task.FromResult(ImportResult);

                    #endregion

                    ParseData(_excelPackage);

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

                    LabelingError(_excelPackage, isSaveLabelingError);
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
            }).ToList().AsQueryable().ToList();

            foreach (var notAllowRepeatCol in notAllowRepeatCols)
            {
                //查询指定列
                //var qDataByProp = qDataList
                //    .Select($"new(RowData.{notAllowRepeatCol.PropertyName} as Value,RowIndex)")
                //    .OrderBy("Value").ToDynamicList();
                var qDataByProp = qDataList;
                //重复行的行号
                var listRepeatRows = new List<int>();
                for (var i = 0; i < qDataByProp.Count; i++)
                {
                    //当前行值
                    var currentValue = qDataByProp[i].RowData;
                    if (i == 0 || string.IsNullOrEmpty(currentValue?.ToString())) continue;

                    //上一行的值
                    var preValue = qDataByProp[i - 1].RowData;
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
        /// <param name="isSaveLabelingError"></param>
        protected virtual void LabelingError(ExcelPackage excelPackage, bool isSaveLabelingError = true)
        {
            //是否标注错误
            if (ExcelImporterSettings.IsLabelingError && ImportResult.HasError)
            {
                var worksheet = GetImportSheet(excelPackage);
                //TODO:标注模板错误
                //标注数据错误
                foreach (var item in ImportResult.RowErrors)
                    foreach (var field in item.FieldErrors)
                    {
                        var col = ImporterHeaderInfos.First(p => p.Header.Name == field.Key);
                        var cell = worksheet.Cells[item.RowIndex, col.Header.ColumnIndex];
                        cell.Style.Font.Color.SetColor(Color.Red);
                        cell.Style.Font.Bold = true;
                        cell.AddComment(string.Join(",", field.Value), col.Header.Author);
                    }
                if (isSaveLabelingError)
                {
                    var ext = Path.GetExtension(FilePath);
                    excelPackage.SaveAs(new FileInfo(FilePath.Replace(ext, "_" + ext)));
                }
            }
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
            ParseImporterHeader();
            try
            {
                //根据名称获取Sheet，如果不存在则取第一个
                var worksheet = GetImportSheet(excelPackage);
                var excelHeaders = new Dictionary<string, int>();
                var endColumnCount = ExcelImporterSettings.EndColumnCount ?? worksheet.Dimension.End.Column;
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
        protected virtual bool ParseImporterHeader()
        {
            ImporterHeaderInfos = new List<ImporterHeaderInfo>();
            var objProperties = _importDataType.GetProperties();
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

            return true;
        }

        /// <summary>
        ///     解析头部
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ArgumentException">导入实体没有定义ImporterHeader属性</exception>
        protected virtual bool ParseImporterHeader(Type sheetType)
        {
            ImporterHeaderInfos = new List<ImporterHeaderInfo>();
            var objProperties = sheetType.GetProperties();
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

            return true;
        }
        /// <summary>
        ///     构建Excel模板
        /// </summary>
        protected virtual void StructureExcel(ExcelPackage excelPackage)
        {
            foreach (var sheetProperty in _sheetPropertyList)
            {
                var sheetType = sheetProperty.PropertyType;
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();

                var worksheet =
                    excelPackage.Workbook.Worksheets.Add(importerAttribute.SheetName);
                if (!ParseImporterHeader(sheetType)) return;

                //设置列头
                for (var i = 0; i < ImporterHeaderInfos.Count; i++)
                {
                    //忽略
                    if (ImporterHeaderInfos[i].Header.IsIgnore) continue;

                    worksheet.Cells[importerAttribute.HeaderRowIndex, i + 1].Value =
                        ImporterHeaderInfos[i].Header.Name;
                    if (!string.IsNullOrWhiteSpace(ImporterHeaderInfos[i].Header.Description))
                        worksheet.Cells[importerAttribute.HeaderRowIndex, i + 1].AddComment(
                            ImporterHeaderInfos[i].Header.Description,
                            ImporterHeaderInfos[i].Header.Author);
                    //如果必填，则列头标红
                    if (ImporterHeaderInfos[i].IsRequired)
                        worksheet.Cells[importerAttribute.HeaderRowIndex, i + 1].Style.Font.Color.SetColor(Color.Red);

                    if (ImporterHeaderInfos[i].MappingValues.Count > 0)
                    {
                        //针对枚举类型和Bool类型添加数据约束
                        var range = ExcelCellBase.GetAddress(importerAttribute.HeaderRowIndex + 1, i + 1,
                            ExcelPackage.MaxRows, i + 1);
                        var dataValidations = worksheet.DataValidations.AddListValidation(range);
                        foreach (var mappingValue in ImporterHeaderInfos[i].MappingValues)
                            dataValidations.Formula.Values.Add(mappingValue.Key);
                    }
                }

                worksheet.Cells.AutoFitColumns();
                worksheet.Cells.Style.WrapText = true;
                worksheet.Cells[worksheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[worksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells[worksheet.Dimension.Address].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[worksheet.Dimension.Address].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[worksheet.Dimension.Address].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[worksheet.Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[worksheet.Dimension.Address].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[worksheet.Dimension.Address].Style.Fill.BackgroundColor.SetColor(Color.DarkSeaGreen);
            }
        }

        /// <summary>
        ///     解析数据
        /// </summary>
        /// <returns></returns>
        protected virtual void ParseData(ExcelPackage excelPackage)
        {
            var worksheet = GetImportSheet(excelPackage);

            //检查导入最大条数限制
            if (ExcelImporterSettings.MaxCount != 0
               && ExcelImporterSettings.MaxCount != int.MaxValue
               && worksheet.Dimension.End.Row > ExcelImporterSettings.MaxCount + ExcelImporterSettings.HeaderRowIndex
               ) throw new ArgumentException($"最大允许导入条数不能超过{ExcelImporterSettings.MaxCount}条！");

            ImportResult.Data = new List<object>();
            var propertyInfos = new List<PropertyInfo>(_importDataType.GetProperties());

            for (var rowIndex = ExcelImporterSettings.HeaderRowIndex + 1;
                rowIndex <= worksheet.Dimension.End.Row;
                rowIndex++)
            {
                var isNullNumber = 1;
                for (var column = 1; column < worksheet.Dimension.End.Column; column++)
                    if (worksheet.Cells[rowIndex, column].Text == string.Empty)
                        isNullNumber++;

                if (isNullNumber < worksheet.Dimension.End.Column)
                {
                    var dataItem = Activator.CreateInstance(_importDataType);
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
            return excelPackage.Workbook.Worksheets[_importDataType.GetDisplayName()]??(
                 ExcelImporterSettings.SheetName != null?
                 excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetName] ??
                 excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetIndex] :
                 excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetIndex]??
                   excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetIndex]);
#else
            return excelPackage.Workbook.Worksheets[_importDataType.GetDisplayName()] ??(
                 ExcelImporterSettings.SheetName != null?
                 excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetName] ??
                 excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetIndex] :
                 excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetIndex]??
                   excelPackage.Workbook.Worksheets[ExcelImporterSettings.SheetIndex]);
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
    }
}
