using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.IO;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 导入图片字段特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ImportImageFieldAttribute : Attribute
    {
        /// <summary>
        ///     图片存储路径（默认存储到临时目录） 
        /// </summary>
        public string ImageDirectory { get; set; } = Path.GetTempPath();

        /// <summary>
        ///     图片导出方式（默认Base64）
        /// </summary>
        public ImportImageTo ImportImageTo { get; set; } = ImportImageTo.Base64;

        /// <summary>
        /// 
        /// </summary>
        public ImportImageFieldAttribute()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="imageDirectory"></param>
        public ImportImageFieldAttribute(string imageDirectory)
        {
            this.ImportImageTo = ImportImageTo.TempFolder;
            this.ImageDirectory = imageDirectory ?? Path.GetTempPath();
        }
    }
}
