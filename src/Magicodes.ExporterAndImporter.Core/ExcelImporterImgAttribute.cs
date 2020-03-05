using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.IO;

namespace Magicodes.ExporterAndImporter.Core
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ExcelImporterImgAttribute : Attribute
    {
        /// <summary>
        ///     是否是图片
        /// </summary>
        public bool IsImg { get; set; }
        /// <summary>
        ///     图片存储路径 
        /// </summary>
        public string FilePath { get; set; }
        /// <summary>
        ///     图片导出方式
        /// </summary>
        public EnumImg EnumImg { get; set; }

        public ExcelImporterImgAttribute()
        {
            FilePath = Directory.GetCurrentDirectory();
            EnumImg = EnumImg.Url;
        }

    }
}
