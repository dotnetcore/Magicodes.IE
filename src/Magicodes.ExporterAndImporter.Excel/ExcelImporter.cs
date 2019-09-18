using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Magicodes.ExporterAndImporter.Core.Extension;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     Excel导入
    /// </summary>
    public class ExcelImporter : IImporter
    {
        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">文件名必须填写! - fileName</exception>
        public Task<TemplateFileInfo> GenerateTemplate<T>(string fileName) where T : class, new()
        {
            using (var importer = new ImportHelper<T>())
            {
                return importer.GenerateTemplate(fileName);
            }
        }

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>二进制字节</returns>
        public Task<byte[]> GenerateTemplateByte<T>() where T : class, new()
        {
            using (var importer = new ImportHelper<T>())
            {
                return importer.GenerateTemplateByte();
            }
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(string filePath) where T : class, new()
        {
            using (var importer = new ImportHelper<T>(filePath))
            {
                return importer.Import();
            }
        }
    }
}