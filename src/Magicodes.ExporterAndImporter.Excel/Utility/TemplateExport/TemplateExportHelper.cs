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
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
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
        private readonly Regex _variableRegex = new Regex(VariableRegexString, RegexOptions.IgnoreCase);

        private readonly Regex _pipeLineVariableRegex = new Regex(PipelineVariableRegexString, RegexOptions.IgnoreCase);

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
        ///     值字典
        /// </summary>
        protected Dictionary<string, string> ValuesDictionary { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        protected T Data { get; set; }

        /// <summary>
        ///     表值字典
        /// </summary>
        protected Dictionary<string, List<Dictionary<string, string>>> TableValuesDictionary { get; set; }

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

                var name = typeof(T).Name;
                switch (name)
                {
                    case "JObject":
                        isJObjectType = true;
                        break;

                    default:
                        isJObjectType = false;
                        break;
                }
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

                var name = typeof(T).Name;
                switch (name)
                {
                    case "Dictionary`2":
                        {
                            isDictionaryType = typeof(T).GetGenericArguments()[0].Equals(typeof(string));
                            break;
                        }
                    default:
                        isDictionaryType = false;
                        break;
                }
                return isDictionaryType.Value;
            }
        }

        public bool IsExpandoObjectType
        {
            get
            {
                if (isExpandoObjectType.HasValue) return isExpandoObjectType.Value;
                isExpandoObjectType = typeof(T).Name == "ExpandoObject";
                return isExpandoObjectType.Value;
            }
        }

        private bool? isJObjectType;

        private bool? isDictionaryType;

        private bool? isExpandoObjectType;

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
                throw new ArgumentException("模板文件路径不能为空!", nameof(TemplateFilePath));
            if (callback == null) return;

            Data = data ?? throw new ArgumentException("数据不能为空!", nameof(data));

            using (Stream stream = new FileStream(TemplateFilePath, FileMode.Open))
            {
                using (var excelPackage = new ExcelPackage(stream))
                {
                    ParseTemplateFile(excelPackage);

                    ParseData(excelPackage);
                    callback.Invoke(excelPackage);
                }
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
                target.SetVariable("data", Data, typeof(T));

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
            foreach (var tableGroup in tableGroups)
            {
                var tableKey = tableGroup.Key;
                var startRow = 0;
                //TODO:处理异常“No property or field”

                var rowCount = 0;
                if (IsDictionaryType || IsExpandoObjectType)
                {
                    IEnumerable<IDictionary<string, object>> tableData = null;
                    tableData = target.Eval<IEnumerable<IDictionary<string, object>>>($"data[\"{tableKey}\"]");
                    //if (IsExpandoObjectType)
                    //    tableData = target.Eval<IEnumerable<IDictionary<string, object>>>($"data.{tableKey}");

                    rowCount = tableData.Count();
                    target.SetVariable(tableKey, tableData, typeof(IEnumerable<IDictionary<string, object>>));
                }
                else
                {
                    rowCount = target.Eval<int>($"data.{tableKey}.Count");
                    var tableData = target.Eval<IEnumerable<dynamic>>($"data.{tableKey}");
                    target.SetVariable(tableKey, tableData);
                }


                var isFirst = true;
                foreach (var col in tableGroup)
                {
                    var address = new ExcelAddressBase(col.TplAddress);
                    if (rowCount == 0)
                    {
                        sheet.Cells[address.Start.Row, address.Start.Column].Value = string.Empty;
                        continue;
                    }
                    //TODO:支持同一行多个表格
                    //行数大于1时需要插入行
                    if (isFirst && rowCount > 1)
                    {
                        startRow = address.Start.Row;
                        //插入行
                        //插入的目标行号
                        var targetRow = address.Start.Row + 1 + insertRows;
                        //插入
                        var numRowsToInsert = rowCount - 1;
                        var refRow = address.Start.Row + insertRows;

                        //sheet.InsertRow(targetRow, numRowsToInsert, refRow);
                        sheet.InsertRow(targetRow, numRowsToInsert);
                        //EPPlus的问题。修复如果存在合并的单元格，但是在新插入的行无法生效的问题，具体见 https://stackoverflow.com/questions/31853046/epplus-copy-style-to-a-range/34299694#34299694

                        //逐行复制效率低，改为多行复制
                        //for (var i = 0; i < numRowsToInsert; i++)
                        //{
                        //    sheet.Cells[String.Format("{0}:{0}", refRow)].Copy(sheet.Cells[String.Format("{0}:{0}", targetRow + i)]);
                        //    //sheet.Row(refRow).StyleID = sheet.Row(targetRow + i).StyleID;
                        //}
                        var maxCloumn = sheet.Dimension.End.Column;
                        RowCopy(sheet, refRow, refRow, rowCount, maxCloumn);
                    }
                    RenderTableCells(target, tbParameters, sheet, insertRows, tableKey, rowCount, col, address);

                    if (isFirst)
                    {
                        isFirst = false;
                    }
                }

                #region 更新单元格

                var updateCellWriters = SheetWriters[sheetName].Where(p => p.WriterType == WriterTypes.Cell).Where(p => p.RowIndex > startRow);
                foreach (var item in updateCellWriters)
                {
                    item.RowIndex += rowCount - 1;
                }

                #endregion 更新单元格

                //表格渲染完成后更新插入的行数
                insertRows += rowCount - 1;
            }
        }

        /// <summary>
        /// 多行复制
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow">复制前的开始行</param>
        /// <param name="endRow">复制前的结束行</param>
        /// <param name="totalRows">总行数</param>
        /// <param name="maxColumnNum">最大列数</param>
        private void RowCopy(ExcelWorksheet sheet, int startRow, int endRow, int totalRows, int maxColumnNum)
        {
            //rows表示现有的sheet行数
            int rows = endRow - startRow + 1;
            if (totalRows > rows * 2)
            {
                //行数复制一倍
                sheet.Cells[startRow, 1, endRow, maxColumnNum].Copy(sheet.Cells[endRow + 1, 1, endRow * 2 - startRow + 1, maxColumnNum]);
                //再次循环
                RowCopy(sheet, startRow, endRow * 2 - startRow + 1, totalRows, maxColumnNum);
            }
            else
            {
                //行数复制需要(需要复制 totalRows - rows)
                sheet.Cells[startRow, 1, startRow + (totalRows - rows) - 1, maxColumnNum].Copy(sheet.Cells[endRow + 1, 1, startRow + totalRows, maxColumnNum]);
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
        private void RenderTableCells(Interpreter target, Parameter[] tbParameters, ExcelWorksheet sheet, int insertRows, string tableKey, int rowCount, IWriter writer, ExcelAddressBase address)
        {
            var cellString = writer.CellString;
            if (cellString.Contains("{{Table>>"))
                //{{ Table >> BookInfo | RowNo}}
                cellString = "{{" + cellString.Split('|')[1].Trim();
            else if (cellString.Contains(">>Table}}"))
                //{{Remark|>>Table}}
                cellString = cellString.Split('|')[0].Trim() + "}}";

            RenderTableCells(target, tbParameters, sheet, insertRows, tableKey, rowCount, cellString, address);
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
            var expresson = writer.CellString;
            RenderCell(target, sheet, expresson, new ExcelAddress(writer.RowIndex, writer.ColIndex, writer.RowIndex, writer.ColIndex).ToString(), dataVar, cellFunc, parameters, invokeParams);
        }

        /// <summary>
        /// 渲染单元格
        /// </summary>
        /// <param name="target"></param>
        /// <param name="sheet"></param>
        /// <param name="expresson"></param>
        /// <param name="cellAddress"></param>
        /// <param name="dataVar"></param>
        /// <param name="cellFunc"></param>
        /// <param name="parameters"></param>
        /// <param name="invokeParams"></param>
        private void RenderCell(Interpreter target, ExcelWorksheet sheet, string expresson, string cellAddress, string dataVar = "\" + data.", Lambda cellFunc = null, Parameter[] parameters = null, params object[] invokeParams)
        {
            //处理单元格渲染管道
            RenderCellPipeline(target, sheet, ref expresson, cellAddress, cellFunc, parameters, dataVar, invokeParams);
            //如果表达式没有处理，则进行处理
            if (expresson.Contains("{{"))
            {
                if (IsDynamicSupportTypes)
                {
                    dataVar = dataVar.TrimEnd('.');
                    expresson = expresson
                                .Replace("{{", dataVar + "[\"")
                                .Replace("}}", "\"] + \"");
                }
                else
                {
                    expresson = expresson
                                .Replace("{{", dataVar)
                                .Replace("}}", " + \"");
                }

                expresson = expresson.StartsWith("\"")
                        ? expresson.TrimStart('\"').TrimStart().TrimStart('+')
                        : "\"" + expresson;

                expresson = expresson.EndsWith("\"")
                    ? expresson.TrimEnd('\"').TrimEnd().TrimEnd('+')
                    : expresson + "\"";

                cellFunc = CreateOrGetCellFunc(target, cellFunc, expresson, parameters);

                var result = cellFunc.Invoke(invokeParams);
                sheet.Cells[cellAddress].Value = IsDynamicSupportTypes ? result?.ToString() : result;
            }
            else if (!string.IsNullOrWhiteSpace(expresson))
            {
                sheet.Cells[cellAddress].Value = expresson;
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
        private void RenderTableCells(Interpreter target, Parameter[] parameters, ExcelWorksheet sheet, int insertRows, string tableKey, int rowCount, string cellString, ExcelAddressBase address)
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
        }

        /// <summary>
        /// 创建或者从缓存中获取
        /// </summary>
        /// <param name="target"></param>
        /// <param name="cellFunc"></param>
        /// <param name="expresson"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        private Lambda CreateOrGetCellFunc(Interpreter target, Lambda cellFunc, string expresson, params Parameter[] parameters)
        {
            if (cellFunc == null)
            {
                if (cellWriteFuncs.ContainsKey(expresson))
                {
                    cellFunc = cellWriteFuncs[expresson];
                }
                else
                {
                    try
                    {
                        cellFunc = parameters == null ? target.Parse(expresson) : target.Parse(expresson, parameters);
                        cellWriteFuncs.Add(expresson, cellFunc);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"构建表达式时出错：{expresson}。", ex);
                    }
                }
            }
            return cellFunc;
        }

        /// <summary>
        /// 渲染单元格管道
        /// </summary>
        /// <param name="target"></param>
        /// <param name="sheet"></param>
        /// <param name="expressonStr"></param>
        /// <param name="cellAddress"></param>
        /// <param name="cellFunc"></param>
        /// <param name="parameters"></param>
        /// <param name="dataVar"></param>
        /// <param name="invokeParams"></param>
        private bool RenderCellPipeline(Interpreter target, ExcelWorksheet sheet, ref string expressonStr, string cellAddress, Lambda cellFunc, Parameter[] parameters, string dataVar, object[] invokeParams)
        {
            if (!expressonStr.Contains("::"))
            {
                return false;
            }
            //匹配所有的管道变量
            var matches = _pipeLineVariableRegex.Matches(expressonStr);
            foreach (Match item in matches)
            {
                var typeKey = Regex.Split(item.Value, "::").First().TrimStart('{').ToLower();
                //参数使用Url参数语法，不支持编码
                //Demo：
                //{{Image::ImageUrl?Width=50&Height=120&Alt=404}}
                //处理特殊字段
                //自定义渲染，以“::”作为切割。
                //TODO:允许注入自定义管道逻辑
                //支持：
                //图：{{Image::ImageUrl?Width=250&Height=70&Alt=404}}

                string body, expresson;
                switch (typeKey)
                {
                    case "image":
                    case "img":
                        {
                            body = Regex.Split(item.Value, "::").Last().TrimEnd('}');
                            var alt = string.Empty;
                            var height = 0;
                            var width = 0;
                            if (body.Contains("?") && body.Contains("="))
                            {
                                var arr = body.Split('?');
                                expresson = arr[0];
                                //从表达式提取Url参数语法内容
                                var values = GetNameVaulesFromQueryStringExpresson(arr[1]);

                                //获取高度
                                var heightStr = values["h"] ?? values["height"];
                                if (!string.IsNullOrWhiteSpace(heightStr))
                                {
                                    height = int.Parse(heightStr);
                                }

                                //获取宽度
                                var widthStr = values["w"] ?? values["width"];
                                if (!string.IsNullOrWhiteSpace(widthStr))
                                {
                                    width = int.Parse(widthStr);
                                }

                                //获取alt文本
                                alt = values["alt"];
                            }
                            else
                            {
                                expresson = body;
                            }
                            expresson = (dataVar + (IsDynamicSupportTypes ? ("[\"" + expresson + "\"]") : (expresson))).Trim('\"').Trim().Trim('+');
                            cellFunc = CreateOrGetCellFunc(target, cellFunc, expresson, parameters);
                            //获取图片地址
                            var imageUrl = cellFunc.Invoke(invokeParams)?.ToString();
                            var cell = sheet.Cells[cellAddress];
                            if (imageUrl == null || (!File.Exists(imageUrl) && !imageUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase)))
                            {
                                cell.Value = alt;
                            }
                            else
                            {
                                try
                                {
                                    var bitmap = Extension.GetBitmapByUrl(imageUrl);
                                    if (bitmap == null)
                                    {
                                        cell.Value = alt;
                                    }
                                    else
                                    {
                                        if (height == default) height = bitmap.Height;
                                        if (width == default) width = bitmap.Width;
                                        cell.Value = string.Empty;
                                        var excelImage = sheet.Drawings.AddPicture(Guid.NewGuid().ToString(), bitmap);
                                        var address = new ExcelAddress(cell.Address);

                                        excelImage.SetPosition(address.Start.Row - 1, 0, address.Start.Column - 1, 0);
                                        excelImage.SetSize(width, height);
                                    }
                                }
                                catch (Exception)
                                {
                                    cell.Value = alt;
                                }
                            }
                            expressonStr = expressonStr.Replace(item.Value, string.Empty);
                        }
                        break;

                    case "formula":
                        body = Regex.Split(item.Value, "::").Last().TrimEnd('}');
                        if (body.Contains("?") && body.Contains("="))
                        {
                            var arr = body.Split('?');
                            var @function = arr[0];
                            var @params = arr[1].Replace("params=", "").Replace("&", ",");
                            var cell = sheet.Cells[cellAddress];
                            cell.Formula = $"={@function}({@params})";
                            expressonStr = expressonStr.Replace(item.Value, string.Empty);
                        }
                        break;

                    default:
                        break;
                }
            }

            return true;
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
            FilePath = null;
            SheetWriters = null;
            ValuesDictionary = null;
            TableValuesDictionary = null;
            cellWriteFuncs.Clear();
        }
    }
}