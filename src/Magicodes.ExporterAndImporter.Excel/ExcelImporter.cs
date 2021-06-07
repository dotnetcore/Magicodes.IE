// ======================================================================
// 
//           filename : ExcelImporter.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core.Extension;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     Excel导入
    /// </summary>
    public class ExcelImporter : IExcelImporter
    {
        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">文件名必须填写! - fileName</exception>
        public Task<ExportFileInfo> GenerateTemplate<T>(string fileName) where T : class, new()
        {
            fileName.CheckExcelFileName();
            var isMultipleSheetType = false;
            var tableType = typeof(T);
            List<PropertyInfo> sheetPropertyList = new List<PropertyInfo>();
            var sheetProperties = tableType.GetProperties();

            for (var i = 0; i < sheetProperties.Length; i++)
            {
                var sheetProperty = sheetProperties[i];
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                if (importerAttribute == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(importerAttribute.SheetName))
                {
                    isMultipleSheetType = true;
                    sheetPropertyList.Add(sheetProperty);
                }
            }

            if (isMultipleSheetType)
            {
                using (var importer = new ImportMultipleSheetHelper(sheetPropertyList))
                {
                    return importer.GenerateTemplate(fileName);
                }
            }
            {
                using (var importer = new ImportHelper<T>())
                {
                    return importer.GenerateTemplate(fileName);
                }
            }
        }

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>二进制字节</returns>
        public Task<byte[]> GenerateTemplateBytes<T>() where T : class, new()
        {
            var isMultipleSheetType = false;
            var tableType = typeof(T);
            List<PropertyInfo> sheetPropertyList = new List<PropertyInfo>();
            var sheetProperties = tableType.GetProperties();

            for (var i = 0; i < sheetProperties.Length; i++)
            {
                var sheetProperty = sheetProperties[i];
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                if (importerAttribute == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(importerAttribute.SheetName))
                {
                    isMultipleSheetType = true;
                    sheetPropertyList.Add(sheetProperty);
                }
            }

            if (isMultipleSheetType)
            {
                using (var importer = new ImportMultipleSheetHelper(sheetPropertyList))
                {
                    return importer.GenerateTemplateByte();
                }
            }
            else
            {
                using (var importer = new ImportHelper<T>())
                {
                    return importer.GenerateTemplateByte();
                }
            }
        }

        /// <summary>
        ///     导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="labelingFilePath"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(string filePath, string labelingFilePath = null) where T : class, new()
        {
            filePath.CheckExcelFileName();
            using (var importer = new ImportHelper<T>(filePath, labelingFilePath))
            {
                return importer.Import();
            }
        }

        /// <summary>
        ///     导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(Stream stream) where T : class, new()
        {
            using (var importer = new ImportHelper<T>(stream))
            {
                return importer.Import();
            }
        }

        /// <summary>
        /// 导出业务错误数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">文件路径</param>
        /// <param name="bussinessErrorDataList">错误数据</param>
        /// <param name="msg">成功:错误数据返回路径,失败 返回错误原因</param>
        /// <returns></returns>
        public bool OutputBussinessErrorData<T>(string filePath, List<DataRowErrorInfo> bussinessErrorDataList, out string msg) where T : class, new()
        {
            filePath.CheckExcelFileName();
            using (var importer = new ImportHelper<T>(filePath, null))
            {
                return importer.OutputBussinessErrorData(bussinessErrorDataList, out msg);
            }
        }

        /// <summary>
        /// 导出业务错误数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">流</param>
        /// <param name="bussinessErrorDataList">错误数据</param>
        /// <param name="fileByte">成功:错误数据返回文件流字节,失败 返回null</param>
        /// <returns></returns>
        public bool OutputBussinessErrorData<T>(Stream stream, List<DataRowErrorInfo> bussinessErrorDataList, out byte[] fileByte) where T : class, new()
        {
            using (var importer = new ImportHelper<T>())
            {
                return importer.OutputBussinessErrorDataByte(stream, bussinessErrorDataList, out fileByte);
            }
        }

        /// <summary>
        /// 导入多个Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <param name="filePath"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型的object装箱，使用时做强转</returns>
        public async Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(string filePath) where T : class, new()
        {
            filePath.CheckExcelFileName();
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }
            var resultList = new Dictionary<string, ImportResult<object>>();
            var tableType = typeof(T);
            var sheetProperties = tableType.GetProperties();
            using (var importer = new ImportMultipleSheetHelper(filePath))
            {
                for (var i = 0; i < sheetProperties.Length; i++)
                {
                    var sheetProperty = sheetProperties[i];
                    var importerAttribute =
                        (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                    if (importerAttribute == null)
                    {
                        throw new Exception($"Sheet属性{sheetProperty.Name}没有标注ExcelImporterAttribute特性");
                    }
                    //if (string.IsNullOrEmpty(importerAttribute.SheetName))
                    //{
                    //    throw new Exception($"Sheet属性{sheetProperty.Name}的ExcelImporterAttribute特性没有设置SheetName");
                    //}
                    bool isSaveLabelingError = i == sheetProperties.Length - 1;
                    //最后一个属性才保存标注的错误,避免多次保存
                    var result = await importer.Import(importerAttribute.SheetName, importerAttribute.SheetIndex,sheetProperty.PropertyType, isSaveLabelingError);
                    resultList.Add(importerAttribute.SheetName ??
                        importerAttribute.SheetIndex.ToString(), result);
                }
            }
            return resultList;
        }

        /// <summary>
        /// 导入多个Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <param name="stream"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型的object装箱，使用时做强转</returns>
        public async Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(Stream stream) where T : class, new()
        {
            var resultList = new Dictionary<string, ImportResult<object>>();
            var tableType = typeof(T);
            var sheetProperties = tableType.GetProperties();
            using (var importer = new ImportMultipleSheetHelper(stream))
            {
                for (var i = 0; i < sheetProperties.Length; i++)
                {
                    var sheetProperty = sheetProperties[i];
                    var importerAttribute =
                        (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                    if (importerAttribute == null)
                    {
                        throw new Exception($"Sheet属性{sheetProperty.Name}没有标注ExcelImporterAttribute特性");
                    }
                    //if (string.IsNullOrEmpty(importerAttribute.SheetName))
                    //{
                    //    throw new Exception($"Sheet属性{sheetProperty.Name}的ExcelImporterAttribute特性没有设置SheetName");
                    //}
                    bool isSaveLabelingError = i == sheetProperties.Length - 1;
                    //最后一个属性才保存标注的错误,避免多次保存
                    var result = await importer.Import(importerAttribute.SheetName, importerAttribute.SheetIndex, sheetProperty.PropertyType, isSaveLabelingError);
                    resultList.Add(importerAttribute.SheetName ??
                        importerAttribute.SheetIndex.ToString(), result);
                }
            }
            return resultList;
        }


        /// <summary>
        /// 导入多个相同类型的Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <typeparam name="TSheet">Sheet类</typeparam>
        /// <param name="filePath"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型TSheet</returns>
        public async Task<Dictionary<string, ImportResult<TSheet>>> ImportSameSheets<T, TSheet>(string filePath)
            where T : class, new() where TSheet : class, new()
        {
            filePath.CheckExcelFileName();
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }
            var resultList = new Dictionary<string, ImportResult<TSheet>>();
            var tableType = typeof(T);
            var sheetProperties = tableType.GetProperties();
            using (var importer = new ImportMultipleSheetHelper(filePath))
            {
                for (var i = 0; i < sheetProperties.Length; i++)
                {
                    var sheetProperty = sheetProperties[i];
                    var importerAttribute =
                        (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                    if (importerAttribute == null)
                    {
                        throw new Exception($"Sheet属性{sheetProperty.Name}没有标注ExcelImporterAttribute特性");
                    }
                    //if (string.IsNullOrEmpty(importerAttribute.SheetName))
                    //{
                    //    throw new Exception($"Sheet属性{sheetProperty.Name}的ExcelImporterAttribute特性没有设置SheetName");
                    //}
                    bool isSaveLabelingError = i == sheetProperties.Length - 1;
                    //最后一个属性才保存标注的错误,避免多次保存
                    var result = await importer.Import(importerAttribute.SheetName, importerAttribute.SheetIndex, sheetProperty.PropertyType, isSaveLabelingError);
                    var tResult = new ImportResult<TSheet>();
                    tResult.Data = new List<TSheet>();
                    if (result.Data.Count > 0)
                    {
                        foreach (var item in result.Data)
                        {
                            tResult.Data.Add((TSheet)item);
                        }
                    }
                    tResult.Exception = result.Exception;
                    tResult.RowErrors = result.RowErrors;
                    tResult.TemplateErrors = result.TemplateErrors;
                    resultList.Add(
                        importerAttribute.SheetName ??
                        importerAttribute.SheetIndex.ToString(),
                        tResult);
                }
            }
            return resultList;
        }


        /// <summary>
        /// 判断Dto类型是否为多Sheet类
        /// </summary>
        /// <typeparam name="T">Dto类型</typeparam>
        /// <returns></returns>
        private bool DtoTypeIsMultipleSheet<T>()
        {
            var tableType = typeof(T);
            var sheetProperties = tableType.GetProperties();

            for (var i = 0; i < sheetProperties.Length; i++)
            {
                var sheetProperty = sheetProperties[i];
                var importerAttribute =
                    (sheetProperty.GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[])?.FirstOrDefault();
                if (importerAttribute == null)
                {
                    continue;
                }
                if (!string.IsNullOrEmpty(importerAttribute.SheetName))
                {
                    return true;
                }
            }
            return false;
        }
    }
}