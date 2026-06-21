// ======================================================================
//
//           filename : TemplateExportHelper.cs
//           description :
//
//           created by 雪雁 at  2020-01-06 16:13
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
//
// ======================================================================

using DynamicExpresso;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.IE.Core;
using Magicodes.IE.EPPlus;
using Magicodes.IE.Excel.Images;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Magicodes.ExporterAndImporter.Excel.Utility.TemplateExport
{
    /// <summary>
    ///     模板导出辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class TemplateExportHelper<T> : IDisposable where T : class
    {
        /// <summary>
		///     匹配表达式
		/// </summary>
		private const string VariableRegexString = "(\\{\\{)+([\\w_.>|\\?:&=]*)+(\\}\\})";

        /// <summary>
        ///     管道匹配表达式
        /// </summary>
        private const string PipelineVariableRegexString = "(\\{\\{)+(img|image|formula)+(::)+([\\w_.>|\\?:&=]*)+(\\}\\})";

        /// <summary>
        ///     变量正则
        /// </summary>
        private readonly Regex _variableRegex = new Regex(VariableRegexString, RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private readonly Regex _pipeLineVariableRegex = new Regex(PipelineVariableRegexString, RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// 用于缓存表达式
        /// </summary>
        private readonly Dictionary<string, Lambda> cellWriteFuncs = new Dictionary<string, Lambda>();

        /// <summary>
        ///     模板文件路径
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        ///     模板文件路径
        /// </summary>
        public string TemplateFilePath { get; set; }

        /// <summary>
        ///     模板写入器
        /// </summary>
        protected Dictionary<string, List<IWriter>> SheetWriters { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        protected T Data { get; set; }

        /// <summary>
        /// 是否是支持的动态类型（JObject、Dictionary（仅支持key为string类型））
        /// </summary>
        private bool IsDynamicSupportTypes
        {
            get
            {
                //TODO:支持DataTable
                return IsDictionaryType || IsJObjectType || IsExpandoObjectType;
            }
        }

        /// <summary>
        /// 是否是JObject类型
        /// </summary>
        public bool IsJObjectType
        {
            get
            {
                if (isJObjectType.HasValue) return isJObjectType.Value;

                // 使用类型比较替代字符串名称比较，更准确
                var type = typeof(T);
                isJObjectType = type.Name == "JObject" && type.Namespace == "Newtonsoft.Json.Linq";
                return isJObjectType.Value;
            }
        }

        /// <summary>
        /// 是否是符合要求的字典类型
        /// </summary>
        public bool IsDictionaryType
        {
            get
            {
                if (isDictionaryType.HasValue) return isDictionaryType.Value;

                // 使用类型比较替代字符串名称比较
                var type = typeof(T);
                if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Dictionary<,>))
                {
                    isDictionaryType = type.GetGenericArguments()[0] == typeof(string);
                }
                else
                {
                    isDictionaryType = false;
                }
                return isDictionaryType.Value;
            }
        }

        public bool IsExpandoObjectType
        {
            get
            {
                if (isExpandoObjectType.HasValue) return isExpandoObjectType.Value;
                // 使用类型比较替代字符串名称比较
                isExpandoObjectType = typeof(T) == typeof(System.Dynamic.ExpandoObject);
                return isExpandoObjectType.Value;
            }
        }

        private bool? isJObjectType;

        private bool? isDictionaryType;

        private bool? isExpandoObjectType;

        /// <summary>
        /// 像素到MTU的转换系数（1像素 = 9525 MTU）
        /// </summary>
        private const int PIXEL_TO_MTU_FACTOR = 9525;

        /// <summary>
        /// 是否已释放资源
        /// </summary>
        private bool _disposed = false;

        /// <summary>
        ///     根据模板导出Excel
        /// </summary>
        /// <param name="templateFilePath">模板文件路径</param>
        /// <param name="data"></param>
        /// <param name="callback">完成导出后执行的操作，默认导出无操作</param>
        public void Export(string templateFilePath, T data, Action<ExcelPackage> callback = null)
        {
            if (!string.IsNullOrWhiteSpace(templateFilePath)) TemplateFilePath = templateFilePath;
            if (string.IsNullOrWhiteSpace(TemplateFilePath))
                throw new ArgumentException(Resource.TemplateFilePathCannotBeEmpty, nameof(TemplateFilePath));
            using (Stream stream = new FileStream(TemplateFilePath, FileMode.Open))
            {
                Export(stream, data, callback);
            }
        }

        /// <summary>
        ///     根据模板导出Excel
        /// </summary>
        /// <param name="templateStream">模板文件流</param>
        /// <param name="data"></param>
        /// <param name="callback"></param>
        /// <exception cref="ArgumentException">完成导出后执行的操作，默认导出无操作</exception>
        public void Export(Stream templateStream, T data, Action<ExcelPackage> callback)
        {
            Data = data ?? throw new ArgumentException(Resource.DataCannotBeEmpty, nameof(data));
            using (var excelPackage = new ExcelPackage(templateStream))
            {
                ParseTemplateFile(excelPackage);
                ParseData(excelPackage);
                callback?.Invoke(excelPackage);
            }
        }


        /// <summary>
        /// 处理数据
        /// </summary>
        /// <param name="excelPackage"></param>
        private void ParseData(ExcelPackage excelPackage)
        {
            var target = new Interpreter();
            //.Reference(typeof(System.Linq.Enumerable))
            //.Reference(typeof(IEnumerable<>))
            //.Reference(typeof(IDictionary<,>));
            if (IsExpandoObjectType)
            {
                target.SetVariable("data", Data, typeof(IDictionary<string, object>));
            }
            else
                target.SetVariable("data", Data, Data.GetType());

            //表格渲染参数
            var tbParameters = new[] {
                new Parameter("index", typeof(int))
            };

            //TODO:渲染支持自定义处理程序
            foreach (var sheetName in SheetWriters.Keys)
            {
                var sheet = excelPackage.Workbook.Worksheets[sheetName];

                //渲染表格
                RenderTable(target, tbParameters, sheetName, sheet);

                //处理普通单元格模板
                RenderCells(target, sheetName, sheet);

                //重新设置行宽（适应图片）
                RenderRowsHeight(sheet);
            }
        }

        /// <summary>
        /// 处理普通单元格模板
        /// </summary>
        /// <param name="target"></param>
        /// <param name="sheetName"></param>
        /// <param name="sheet"></param>
        private void RenderCells(Interpreter target, string sheetName, ExcelWorksheet sheet)
        {
            foreach (var writer in SheetWriters[sheetName].Where(p => p.WriterType == WriterTypes.Cell))
            {
                RenderCell(target, sheet, writer);
            }
        }

        /// <summary>
        /// 渲染表格
        /// </summary>
        /// <param name="target"></param>
        /// <param name="tbParameters"></param>
        /// <param name="sheetName"></param>
        /// <param name="sheet"></param>
        private void RenderTable(Interpreter target, Parameter[] tbParameters, string sheetName, ExcelWorksheet sheet)
        {
            var tableGroups = SheetWriters[sheetName].Where(p => p.WriterType == WriterTypes.Table)
                .GroupBy(p => p.TableKey);

            var insertRows = 0;
            //支持一行多表格
            //1）获取所有表格的区域范围（列数行数以及坐标）
            var tableInfoList = new List<TemplateTableInfo>(tableGroups.Count());
            foreach (var tableGroup in tableGroups)
            {
                var tableKey = tableGroup.Key;
                var rowCount = 0;
                if (IsDictionaryType || IsExpandoObjectType)
                {
                    IEnumerable<IDictionary<string, object>> tableData = null;
                    try
                    {
                        tableData = target.Eval<IEnumerable<IDictionary<string, object>>>($"data[\"{tableKey}\"]");
                        rowCount = tableData?.Count() ?? 0;
                        if (tableData != null)
                        {
                            target.SetVariable(tableKey, tableData, typeof(IEnumerable<IDictionary<string, object>>));
                        }
                    }
                    catch
                    {
                        rowCount = 0;
                    }
                }
                else
                {
                    try
                    {
                        rowCount = target.Eval<int>($"data.{tableKey}.Count");
                        var tableData = target.Eval<IEnumerable<dynamic>>($"data.{tableKey}");
                        if (tableData != null)
                        {
                            target.SetVariable(tableKey, tableData);
                        }
                    }
                    catch
                    {
                        rowCount = 0;
                    }
                }
                var startCol = tableGroup.OrderBy(p => p.ColIndex).First();
                var rowStart = startCol.RowIndex;
                var tableInfo = new TemplateTableInfo()
                {
                    TableKey = tableKey,
                    RawRowStart = rowStart,
                    NewRowStart = rowStart,
                    RowCount = rowCount,
                    Writers = tableGroup
                };
                tableInfoList.Add(tableInfo);
            }

            var rowTableGroups = tableInfoList.GroupBy(p => p.RawRowStart);
            foreach (var item in rowTableGroups)
            {
                //是否为一行多个Table
                var isManyTable = item.Count() > 1;
                //一行多Table以最大的为准
                TemplateTableInfo table = !isManyTable ? item.First() : item.OrderByDescending(p => p.RowCount).First();

                if (table.RowCount == 0)
                {
                    continue;
                }

                if (isManyTable)
                {
                    foreach (var itemTable in item)
                    {
                        itemTable.NewRowStart += insertRows;
                        // 如果一行有多表格则记录最大行数，用于后续清理多余的行
                        itemTable.SameRowMaxRowCount = table.RowCount;
                    }
                }
                else
                    table.NewRowStart += insertRows;

                //2）统一插入行
                var startRow = table.NewRowStart;
                //插入行
                //插入的目标行号
                var targetRow = table.NewRowStart + 1;
                //插入
                var numRowsToInsert = table.RowCount - 1;
                var refRow = table.NewRowStart;

                if (numRowsToInsert == 0) continue;
                sheet.InsertRow(targetRow, numRowsToInsert);
                //EPPlus的问题。修复如果存在合并的单元格，但是在新插入的行无法生效的问题，具体见 https://stackoverflow.com/questions/31853046/epplus-copy-style-to-a-range/34299694#34299694

                var maxCloumn = sheet.Dimension.End.Column;
                RowCopy(sheet, refRow, refRow, table.RowCount, maxCloumn);

                #region 更新单元格

                var updateCellWriters = SheetWriters[sheetName].Where(p => p.WriterType == WriterTypes.Cell).Where(p => p.RowIndex > table.NewRowStart);
                foreach (var writer in updateCellWriters)
                {
                    writer.RowIndex += table.RowCount - 1;
                }

                #endregion 更新单元格

                //表格渲染完成后更新插入的行数
                insertRows += table.RowCount - 1;
            }

            //4）渲染表格
            foreach (var table in tableInfoList)
            {
                var tableGroup = table.Writers;

                var tableKey = tableGroup.Key;
                //TODO:处理异常“No property or field”

                foreach (var col in tableGroup)
                {
                    var address = new ExcelAddressBase(col.TplAddress);
                    if (table.RowCount == 0)
                    {
                        sheet.Cells[address.Start.Row, address.Start.Column].Value = string.Empty;
                        continue;
                    }

                    RenderTableCells(target, tbParameters, sheet, table.NewRowStart - table.RawRowStart, tableKey, table.RowCount, col, address, table.SameRowMaxRowCount);
                }
            }
        }

        /// <summary>
        /// 多行复制（迭代实现，避免递归导致的栈溢出）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow">复制前的开始行</param>
        /// <param name="endRow">复制前的结束行</param>
        /// <param name="totalRows">总行数</param>
        /// <param name="maxColumnNum">最大列数</param>
        private void RowCopy(ExcelWorksheet sheet, int startRow, int endRow, int totalRows, int maxColumnNum)
        {
            const int maxExcelRows = 1048576; // Excel 2007+ 最大行数

            if (totalRows <= 1)
            {
                return;
            }

            var currentStartRow = startRow;
            var currentEndRow = endRow;

            var desiredEndRow = Math.Min(startRow + totalRows - 1, maxExcelRows);
            if (currentEndRow >= desiredEndRow)
            {
                return;
            }

            while (currentEndRow < desiredEndRow)
            {
                var copiedRows = currentEndRow - currentStartRow + 1;
                if (copiedRows <= 0)
                {
                    break;
                }

                var targetStartRow = currentEndRow + 1;
                var targetEndRow = Math.Min(currentEndRow + copiedRows, desiredEndRow);
                var targetRows = targetEndRow - targetStartRow + 1;
                if (targetRows <= 0)
                {
                    break;
                }

                var sourceEndRow = currentStartRow + targetRows - 1;
                if (sourceEndRow > currentEndRow)
                {
                    sourceEndRow = currentEndRow;
                }

                sheet.Cells[currentStartRow, 1, sourceEndRow, maxColumnNum]
                    .Copy(sheet.Cells[targetStartRow, 1, targetStartRow + (sourceEndRow - currentStartRow), maxColumnNum]);

                currentEndRow = targetEndRow;
            }
        }

        /// <summary>
        /// 重新设置行宽（适应图片）
        /// </summary>
        /// <param name="sheet"></param>
        private static void RenderRowsHeight(ExcelWorksheet sheet)
        {
            var rows = new List<int>();
            foreach (ExcelDrawing item in sheet.Drawings)
            {
                if (item is ExcelPicture pic)
                {
                    var rowIndex = pic.From.Row + 1;
                    if (rows.Contains(rowIndex))
                    {
                        continue;
                    }
                    //https://github.com/dotnetcore/Magicodes.IE/issues/131
                    //sheet.Row(rowIndex).Height = pic.Image.Height;
                    sheet.Row(rowIndex).Height = pic.GetPrivateProperty<int>("_height");
                    rows.Add(rowIndex);
                }
            }
            rows.Clear();
        }

        /// <summary>
        /// 渲染表格单元格
        /// </summary>
        /// <param name="target"></param>
        /// <param name="tbParameters"></param>
        /// <param name="sheet"></param>
        /// <param name="insertRows"></param>
        /// <param name="tableKey"></param>
        /// <param name="rowCount"></param>
        /// <param name="writer"></param>
        /// <param name="address"></param>
        /// <param name="sameRowMaxRowCount"></param>
        private void RenderTableCells(Interpreter target, Parameter[] tbParameters, ExcelWorksheet sheet, int insertRows, string tableKey, int rowCount, IWriter writer, ExcelAddressBase address, int? sameRowMaxRowCount = null)
        {
            var cellString = writer.CellString;
            if (cellString.Contains("{{Table>>"))
            {
                //{{ Table >> BookInfo | RowNo}}
                var parts = cellString.Split('|');
                if (parts.Length > 1)
                {
                    cellString = "{{" + parts[1].Trim();
                }
            }
            else if (cellString.Contains(">>Table}}"))
            {
                //{{Remark|>>Table}}
                var parts = cellString.Split('|');
                if (parts.Length > 0)
                {
                    cellString = parts[0].Trim() + "}}";
                }
            }

            RenderTableCells(target, tbParameters, sheet, insertRows, tableKey, rowCount, cellString, address, sameRowMaxRowCount);
        }

        /// <summary>
        /// 渲染单元格
        /// </summary>
        /// <param name="target"></param>
        /// <param name="sheet"></param>
        /// <param name="writer"></param>
        /// <param name="dataVar"></param>
        /// <param name="cellFunc"></param>
        /// <param name="parameters"></param>
        /// <param name="invokeParams"></param>
        private void RenderCell(Interpreter target, ExcelWorksheet sheet, IWriter writer, string dataVar = "\" + data.", Lambda cellFunc = null, Parameter[] parameters = null, params object[] invokeParams)
        {
            var expression = writer.CellString;
            RenderCell(target, sheet, expression, new ExcelAddress(writer.RowIndex, writer.ColIndex, writer.RowIndex, writer.ColIndex).ToString(), dataVar, cellFunc, parameters, invokeParams);
        }

        /// <summary>
        /// 渲染单元格
        /// </summary>
        /// <param name="target"></param>
        /// <param name="sheet"></param>
        /// <param name="expression"></param>
        /// <param name="cellAddress"></param>
        /// <param name="dataVar"></param>
        /// <param name="cellFunc"></param>
        /// <param name="parameters"></param>
        /// <param name="invokeParams"></param>
        private void RenderCell(Interpreter target, ExcelWorksheet sheet, string expression, string cellAddress, string dataVar = "\" + data.", Lambda cellFunc = null, Parameter[] parameters = null, params object[] invokeParams)
        {
            //处理单元格渲染管道
            RenderCellPipeline(target, sheet, ref expression, cellAddress, cellFunc, parameters, dataVar, invokeParams);
            //如果表达式没有处理，则进行处理
            if (expression.Contains("{{"))
            {
                // 使用StringBuilder优化字符串操作
                var sb = new StringBuilder(expression);
                if (IsDynamicSupportTypes)
                {
                    dataVar = dataVar.TrimEnd('.');
                    sb.Replace("{{", dataVar + "[\"");
                    sb.Replace("}}", "\"] + \"");
                }
                else
                {
                    sb.Replace("{{", dataVar);
                    sb.Replace("}}", " + \"");
                }

                expression = sb.ToString();

                expression = expression.StartsWith("\"")
                        ? expression.TrimStart('\"').TrimStart().TrimStart('+')
                        : "\"" + expression;

                expression = expression.EndsWith("\"")
                    ? expression.TrimEnd('\"').TrimEnd().TrimEnd('+')
                    : expression + "\"";

                cellFunc = CreateOrGetCellFunc(target, cellFunc, expression, parameters);

                try
                {
                    var result = cellFunc.Invoke(invokeParams);
                    sheet.Cells[cellAddress].Value = result;
                }
                catch (Exception ex)
                {
                    // 处理表达式执行异常，记录并设置默认值
                    System.Diagnostics.Debug.WriteLine($"表达式执行失败: {expression}, 错误: {ex.Message}");
                    sheet.Cells[cellAddress].Value = string.Empty;
                }
            }
            else if (!string.IsNullOrWhiteSpace(expression))
            {
                sheet.Cells[cellAddress].Value = expression;
            }
        }

        /// <summary>
        /// 渲染多个单元格(一列数据)
        /// </summary>
        /// <param name="target"></param>
        /// <param name="parameters"></param>
        /// <param name="sheet"></param>
        /// <param name="insertRows"></param>
        /// <param name="tableKey"></param>
        /// <param name="rowCount"></param>
        /// <param name="cellString"></param>
        /// <param name="address"></param>
        /// <param name="sameRowMaxRowCount"></param>
        private void RenderTableCells(Interpreter target, Parameter[] parameters, ExcelWorksheet sheet, int insertRows, string tableKey, int rowCount, string cellString, ExcelAddressBase address, int? sameRowMaxRowCount = null)
        {
            //var dataVar = !IsDynamicSupportTypes ? ("\" + data." + tableKey + "[index].") : ("\" + data[\"" + tableKey + "\"][index]");
            string dataVar;
            if (IsDictionaryType || IsExpandoObjectType)
            {
                dataVar = ($"\" + {tableKey}.Skip(index).First()");
            }
            else if (IsJObjectType)
            {
                dataVar = $"\" + data[\"{tableKey}\"][index]";
            }
            else
            {
                dataVar = $"\" + {tableKey}.Skip(index).First().";
            }

            //渲染一列单元格
            for (var i = 0; i < rowCount; i++)
            {
                var rowIndex = address.Start.Row + i + insertRows;
                var targetAddress = new ExcelAddress(rowIndex, address.Start.Column, rowIndex, address.Start.Column);
                //https://github.com/dotnetcore/Magicodes.IE/issues/155
                sheet.Row(rowIndex).Height = sheet.Row(address.Start.Row).Height;
                RenderCell(target, sheet, cellString, targetAddress.Address, dataVar, null, parameters, i);
            }

            if (sameRowMaxRowCount.HasValue && sameRowMaxRowCount.Value > rowCount)
            {
                // 清理多余的行
                for (var i = rowCount; i < sameRowMaxRowCount.Value; i++)
                {
                    var rowIndex = address.Start.Row + i + insertRows;
                    var targetAddress = new ExcelAddress(rowIndex, address.Start.Column, rowIndex, address.Start.Column);
                    sheet.Cells[targetAddress.Address].Clear();
                }
            }
        }

        /// <summary>
        /// 创建或者从缓存中获取
        /// </summary>
        /// <param name="target"></param>
        /// <param name="cellFunc"></param>
        /// <param name="expression"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        private Lambda CreateOrGetCellFunc(Interpreter target, Lambda cellFunc, string expression, params Parameter[] parameters)
        {
            if (cellFunc == null)
            {
                // 缓存键应包含参数信息，避免不同参数的相同表达式冲突
                var cacheKey = GetCacheKey(expression, parameters);
                if (cellWriteFuncs.ContainsKey(cacheKey))
                {
                    cellFunc = cellWriteFuncs[cacheKey];
                }
                else
                {
                    try
                    {
                        cellFunc = parameters == null || parameters.Length == 0 ? target.Parse(expression) : target.Parse(expression, parameters);
                        cellWriteFuncs.Add(cacheKey, cellFunc);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"{Resource.ErrorBuildingExpression}{expression}。", ex);
                    }
                }
            }
            return cellFunc;
        }

        /// <summary>
        /// 生成缓存键，包含表达式和参数信息
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        private string GetCacheKey(string expression, Parameter[] parameters)
        {
            if (parameters == null || parameters.Length == 0)
            {
                return expression;
            }
            var paramTypes = string.Join(",", parameters.Select(p => p.Type?.FullName ?? p.Name));
            return $"{expression}|Params:{paramTypes}";
        }

        /// <summary>
        /// 渲染单元格管道
        /// </summary>
        /// <param name="target"></param>
        /// <param name="sheet"></param>
        /// <param name="expressionStr"></param>
        /// <param name="cellAddress"></param>
        /// <param name="cellFunc"></param>
        /// <param name="parameters"></param>
        /// <param name="dataVar"></param>
        /// <param name="invokeParams"></param>
        private bool RenderCellPipeline(Interpreter target, ExcelWorksheet sheet, ref string expressionStr, string cellAddress, Lambda cellFunc, Parameter[] parameters, string dataVar, object[] invokeParams)
        {
            if (!expressionStr.Contains("::"))
            {
                return false;
            }
            //匹配所有的管道变量
            var matches = _pipeLineVariableRegex.Matches(expressionStr);
            foreach (Match item in matches)
            {
                // 使用 string.Split 替代 Regex.Split 处理简单分隔符
                var parts = item.Value.Split(new[] { "::" }, StringSplitOptions.None);
                if (parts.Length < 2) continue;
                
                var typeKey = parts[0].TrimStart('{').ToLower();
                //参数使用Url参数语法，不支持编码
                //Demo：
                //{{Image::ImageUrl?Width=50&Height=120&Alt=404}}
                //处理特殊字段
                //自定义渲染，以"::"作为切割。
                //TODO:允许注入自定义管道逻辑
                //支持：
                //图：{{Image::ImageUrl?Width=250&Height=70&Alt=404}}

                switch (typeKey)
                {
                    case "image":
                    case "img":
                        {
                            ProcessImagePipeline(target, sheet, ref expressionStr, cellAddress, cellFunc, parameters, dataVar, invokeParams, item.Value, parts);
                        }
                        break;

                    case "formula":
                        {
                            ProcessFormulaPipeline(sheet, ref expressionStr, cellAddress, item.Value, parts);
                        }
                        break;

                    default:
                        break;
                }
            }

            return true;
        }

        /// <summary>
        /// 处理图片管道
        /// </summary>
        private void ProcessImagePipeline(Interpreter target, ExcelWorksheet sheet, ref string expressionStr, string cellAddress, Lambda cellFunc, Parameter[] parameters, string dataVar, object[] invokeParams, string matchValue, string[] parts)
        {
            var body = parts.Length > 1 ? parts[1].TrimEnd('}') : string.Empty;
            var (expression, alt, height, width, xOffset, yOffset) = ParseImageParameters(body);

            var finalExpression = (dataVar + (IsDynamicSupportTypes ? ("[\"" + expression + "\"]") : (expression))).Trim('\"').Trim().Trim('+');
            cellFunc = CreateOrGetCellFunc(target, cellFunc, finalExpression, parameters);
            //获取图片地址
            string imageUrl = null;
            try
            {
                imageUrl = cellFunc.Invoke(invokeParams)?.ToString();
            }
            catch (KeyNotFoundException)
            {
                // 字典或ExpandoObject中缺少字段，使用默认值
                imageUrl = null;
            }
            catch (Exception ex)
            {
                // 其他异常也记录并继续
                System.Diagnostics.Debug.WriteLine($"获取图片URL失败: {expression}, 错误: {ex.Message}");
                imageUrl = null;
            }
            var cell = sheet.Cells[cellAddress];
            
            // 修正空引用检查逻辑
            if (string.IsNullOrEmpty(imageUrl) || (!File.Exists(imageUrl) && !imageUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase) && !imageUrl.IsBase64StringValid()))
            {
                cell.Value = alt;
            }
            else
            {
                try
                {
                    Image image = null;
                    IImageFormat format = default;
                    
                    if (imageUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    {
                        image = imageUrl.GetImageByUrl(out format);
                    }
                    else if (imageUrl.IsBase64StringValid())
                    {
                        image = imageUrl.Base64StringToImage(out format);
                    }
                    else if (File.Exists(imageUrl))
                    {
                        using (Stream imageStream = File.OpenRead(imageUrl))
                        {
                            image = Image.Load(imageStream);
                            format = image.GetImageFormat(imageStream);
                        }
                    }

                    if (image == null)
                    {
                        cell.Value = alt;
                    }
                    else
                    {
                        // 使用 using 确保 Image 资源正确释放
                        using (image)
                        {
                            if (height == default) height = image.Height;
                            if (width == default) width = image.Width;
                            cell.Value = string.Empty;
                            var excelImage = sheet.Drawings.AddPicture(Guid.NewGuid().ToString(), image, format);
                            var address = new ExcelAddress(cell.Address);
                            ////调整对齐
                            excelImage.From.ColumnOff = Pixel2MTU(xOffset);
                            excelImage.From.RowOff = Pixel2MTU(yOffset);
                            excelImage.From.Column = address.Start.Column - 1;
                            excelImage.From.Row = address.Start.Row - 1;
                            //excelImage.SetPosition(address.Start.Row - 1, 0, address.Start.Column - 1, 0);
                            excelImage.SetSize(width, height);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录异常信息而不是吞掉异常
                    // 注意：这里可能需要日志框架，暂时保留原有行为但添加异常信息
                    System.Diagnostics.Debug.WriteLine($"图片处理失败: {ex.Message}");
                    cell.Value = alt;
                }
            }
            expressionStr = expressionStr.Replace(matchValue, string.Empty);
        }

        /// <summary>
        /// 解析图片参数
        /// </summary>
        private (string expression, string alt, int height, int width, int xOffset, int yOffset) ParseImageParameters(string body)
        {
            var alt = string.Empty;
            var height = 0;
            var width = 0;
            var xOffset = 0;
            var yOffset = 0;
            string expression;

            if (body.Contains("?") && body.Contains("="))
            {
                var arr = body.Split('?');
                expression = arr.Length > 0 ? arr[0] : body;
                //从表达式提取Url参数语法内容
                if (arr.Length > 1)
                {
                    var values = GetNameVaulesFromQueryStringExpresson(arr[1]);

                    //获取高度 - 使用 TryParse 替代 Parse
                    var heightStr = values["h"] ?? values["height"];
                    if (!string.IsNullOrWhiteSpace(heightStr))
                    {
                        int.TryParse(heightStr, out height);
                    }

                    //获取宽度 - 使用 TryParse 替代 Parse
                    var widthStr = values["w"] ?? values["width"];
                    if (!string.IsNullOrWhiteSpace(widthStr))
                    {
                        int.TryParse(widthStr, out width);
                    }

                    //获取XOffset - 使用 TryParse 替代 Parse
                    var xOffsetStr = values["XOffset"] ?? values["x"];
                    if (!string.IsNullOrWhiteSpace(xOffsetStr))
                    {
                        int.TryParse(xOffsetStr, out xOffset);
                    }

                    //获取YOffset - 使用 TryParse 替代 Parse
                    var yOffsetStr = values["YOffset"] ?? values["y"];
                    if (!string.IsNullOrWhiteSpace(yOffsetStr))
                    {
                        int.TryParse(yOffsetStr, out yOffset);
                    }

                    //获取alt文本
                    alt = values["alt"] ?? string.Empty;
                }
            }
            else
            {
                expression = body;
            }

            return (expression, alt, height, width, xOffset, yOffset);
        }

        /// <summary>
        /// 处理公式管道
        /// </summary>
        private void ProcessFormulaPipeline(ExcelWorksheet sheet, ref string expressionStr, string cellAddress, string matchValue, string[] parts)
        {
            var body = parts.Length > 1 ? parts[1].TrimEnd('}') : string.Empty;
            if (body.Contains("?") && body.Contains("="))
            {
                var arr = body.Split('?');
                if (arr.Length > 1)
                {
                    var function = arr[0];
                    var @params = arr[1].Replace("params=", "").Replace("&", ",");
                    var cell = sheet.Cells[cellAddress];
                    cell.Formula = $"={function}({@params})";
                    expressionStr = expressionStr.Replace(matchValue, string.Empty);
                }
            }
        }

        /// <summary>
        /// 将像素转换为MTU（Measurement Unit）
        /// </summary>
        /// <param name="pixels">像素值</param>
        /// <returns>MTU值</returns>
        internal static int Pixel2MTU(int pixels)
        {
            return pixels * PIXEL_TO_MTU_FACTOR;
        }

        /// <summary>
        ///     验证并转换模板
        /// </summary>
        /// <param name="excelPackage"></param>
        protected void ParseTemplateFile(ExcelPackage excelPackage)
        {
            SheetWriters = new Dictionary<string, List<IWriter>>();
            foreach (var worksheet in excelPackage.Workbook.Worksheets)
            {
                if (worksheet.Dimension == null)
                    continue;
                var writers = new List<IWriter>();
                if (!SheetWriters.ContainsKey(worksheet.Name)) SheetWriters.Add(worksheet.Name, writers);
                var endColumnIndex = worksheet.Dimension.End.Column;
                var endRowIndex = worksheet.Dimension.End.Row;

                //获取所有包含表达式的单元格
                var q = (from cell in worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                        endRowIndex, endColumnIndex]
                         where _variableRegex.IsMatch((cell.Value ?? string.Empty).ToString())
                         select cell).ToList();

                var rows = q.GroupBy(p => p.Rows);

                foreach (var rowGroups in rows)
                {
                    var isStartTable = false;
                    string tableKey = null;
                    foreach (var cell in rowGroups)
                    {
                        var cellString = cell.Value.ToString();
                        if (cellString.Contains("{{Table>>"))
                        {
                            isStartTable = true;
                            //{{ Table >> BookInfo | RowNo}}
                            tableKey = Regex.Split(cellString, "{{Table>>")[1].Split('|')[0].Trim();
                        }

                        writers.Add(new Writer
                        {
                            TableKey = tableKey,
                            TplAddress = cell.Address,
                            CellString = cellString,
                            WriterType = isStartTable ? WriterTypes.Table : WriterTypes.Cell,
                            RowIndex = cell.Start.Row,
                            ColIndex = cell.Start.Column
                        });

                        if (isStartTable && cellString.Contains(">>Table}}"))
                        {
                            isStartTable = false;
                            tableKey = null;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 将查询字符串解析转换为名值集合.
        /// </summary>
        /// <param name="queryStringExpresson"></param>
        /// <returns></returns>
        public static NameValueCollection GetNameVaulesFromQueryStringExpresson(string queryStringExpresson)
        {
            queryStringExpresson = queryStringExpresson.Replace("?", string.Empty);
            var result = new NameValueCollection(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(queryStringExpresson))
            {
                int count = queryStringExpresson.Length;
                for (int i = 0; i < count; i++)
                {
                    int startIndex = i;
                    int index = -1;
                    while (i < count)
                    {
                        char item = queryStringExpresson[i];
                        if (item == '=')
                        {
                            if (index < 0)
                            {
                                index = i;
                            }
                        }
                        else if (item == '&')
                        {
                            break;
                        }
                        i++;
                    }

                    string value = null;
                    string key;
                    if (index >= 0)
                    {
                        key = queryStringExpresson.Substring(startIndex, index - startIndex).ToLower();
                        value = queryStringExpresson.Substring(index + 1, (i - index) - 1);
                    }
                    else
                    {
                        key = queryStringExpresson.Substring(startIndex, i - startIndex).ToLower();
                    }
                    result[key] = value;
                    if ((i == (count - 1)) && (queryStringExpresson[i] == '&'))
                    {
                        result[key] = string.Empty;
                    }
                }
            }
            return result;
        }

        /// <summary>执行与释放或重置非托管资源关联的应用程序定义的任务。</summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        /// <param name="disposing">是否释放托管资源</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // 释放托管资源
                    FilePath = null;
                    TemplateFilePath = null;
                    SheetWriters = null;
                    Data = null;
                    cellWriteFuncs?.Clear();
                }

                // 释放非托管资源（如果有）
                // 正则表达式对象会在类销毁时自动释放

                _disposed = true;
            }
        }
    }
}
