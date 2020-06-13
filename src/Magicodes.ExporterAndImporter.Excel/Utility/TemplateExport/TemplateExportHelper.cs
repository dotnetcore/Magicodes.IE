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

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DynamicExpresso;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

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
		///     变量正则
		/// </summary>
		private readonly Regex _variableRegex = new Regex(VariableRegexString, RegexOptions.IgnoreCase);

		/// <summary>
		/// 用于缓存表达式
		/// </summary>
		private Dictionary<string, Lambda> cellWriteFuncs = new Dictionary<string, Lambda>();

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
			if (data == null) throw new ArgumentException("数据不能为空!", nameof(data));
			if (callback == null) return;
           
			Data = data;

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
			target.SetVariable("data", Data, typeof(T));

			//表格渲染参数
			var tbParameters = new[] {
				new Parameter("index", typeof(int))
			};

			//TODO:渲染支持自定义处理程序
			foreach (var sheetName in SheetWriters.Keys)
			{
				var sheet = excelPackage.Workbook.Worksheets[sheetName];
				//处理普通单元格模板
				RenderCells(target, sheetName, sheet);
				//渲染表格
				RenderTable(target, tbParameters, sheetName, sheet);
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
				RenderCell(target, sheet, writer.CellString, writer.Address);
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
				//TODO:处理异常“No property or field”
				var rowCount = target.Eval<int>($"data.{tableKey}.Count");
				Console.WriteLine($"正在处理表格【{tableKey}】，行数：{rowCount}。");
				var isFirst = true;
				foreach (var col in tableGroup)
				{
					var address = new ExcelAddressBase(col.Address);
					if (rowCount == 0)
					{
						sheet.Cells[address.Start.Row, address.Start.Column].Value = string.Empty;
						continue;
					}
					//TODO:支持同一行多个表格
					//行数大于1时需要插入行
					if (isFirst && rowCount > 1)
					{
						//插入行
						//插入的目标行号
						var targetRow = address.Start.Row + 1 + insertRows;
						//插入
						var numRowsToInsert = rowCount - 1;
						var refRow = address.Start.Row + insertRows;

						//sheet.InsertRow(targetRow, numRowsToInsert, refRow);
						sheet.InsertRow(targetRow, numRowsToInsert);
						//EPPlus的问题。修复如果存在合并的单元格，但是在新插入的行无法生效的问题，具体见 https://stackoverflow.com/questions/31853046/epplus-copy-style-to-a-range/34299694#34299694
						for (var i = 0; i < numRowsToInsert; i++)
						{
							sheet.Cells[String.Format("{0}:{0}", refRow)].Copy(sheet.Cells[String.Format("{0}:{0}", targetRow + i)]);
							//sheet.Row(refRow).StyleID = sheet.Row(targetRow + i).StyleID;
						}
					}

					RenderTableCells(target, tbParameters, sheet, insertRows, tableKey, rowCount, col, address);

					if (isFirst)
					{
						isFirst = false;
					}
				}
				//表格渲染完成后更新插入的行数
				insertRows += rowCount - 1;
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
				if (item is ExcelPicture)
				{
					var pic = item as ExcelPicture;
					var rowIndex = item.From.Row + 1;
					if (rows.Contains(rowIndex))
					{
						continue;
					}

					sheet.Row(rowIndex).Height = pic.Image.Height;
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

			RenderCells(target, tbParameters, sheet, insertRows, tableKey, rowCount, cellString, address);
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
			if (!RenderCellPipeline(target, sheet, expresson, cellAddress, cellFunc, parameters, dataVar, invokeParams))
			{
				//如果表达式没有处理，则进行处理
				if (expresson.Contains("{{"))
				{
					expresson = expresson
									.Replace("{{", dataVar)
									.Replace("}}", " + \"");

					expresson = expresson.StartsWith("\"")
						? expresson.TrimStart('\"').TrimStart().TrimStart('+')
						: "\"" + expresson;

					expresson = expresson.EndsWith("\"")
						? expresson.TrimEnd('\"').TrimEnd().TrimEnd('+')
						: expresson + "\"";
				}
				cellFunc = CreateOrGetCellFunc(target, cellFunc, expresson, parameters);

				var result = cellFunc.Invoke(invokeParams);
				sheet.Cells[cellAddress].Value = result;
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
		private void RenderCells(Interpreter target, Parameter[] parameters, ExcelWorksheet sheet, int insertRows, string tableKey, int rowCount, string cellString, ExcelAddressBase address)
		{
			var dataVar = "\" + data." + tableKey + "[index].";
			//渲染一列单元格
			for (var i = 0; i < rowCount; i++)
			{
				var rowIndex = address.Start.Row + i + insertRows;
				var targetAddress = new ExcelAddress(rowIndex, address.Start.Column, rowIndex, address.Start.Column);
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
		private bool RenderCellPipeline(Interpreter target, ExcelWorksheet sheet, string expressonStr, string cellAddress, Lambda cellFunc, Parameter[] parameters, string dataVar, object[] invokeParams)
		{
			if (!expressonStr.Contains("::"))
			{
				return false;
			}
			//参数使用Url参数语法，不支持编码
			//Demo：
			//{{Image::ImageUrl?Width=50&Height=120&Alt=404}}
			//处理特殊字段
			//自定义渲染，以“::”作为切割。
			//TODO:允许注入自定义管道逻辑
			var typeKey = Regex.Split(expressonStr, "::").First().TrimStart('{').ToLower();
			switch (typeKey)
			{
				case "image":
				case "img":
					{
						var body = Regex.Split(expressonStr, "::").Last().TrimEnd('}');
						var expresson = string.Empty;
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
						expresson = (dataVar + expresson).Trim('\"').Trim().Trim('+');
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
					}
					break;
				case "formula":
                    string str = "执行到这里了哈哈哈";
					
                    break;
                default:
					break;
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
							Address = cell.Address,
							CellString = cellString,
							WriterType = isStartTable ? WriterTypes.Table : WriterTypes.Cell,
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