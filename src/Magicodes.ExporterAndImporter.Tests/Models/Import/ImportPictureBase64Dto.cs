using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using System;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public class ImportPictureBase64Dto
    {
        [ImporterHeader(Name = "加粗文本")]
        public string Text { get; set; }
        [ImporterHeader(Name = "普通文本")]
        public string Text2 { get; set; }

        /// <summary>
        /// 将图片导入为base64（默认为base64）
        /// </summary>
        [ImportImageField(ImportImageTo = ImportImageTo.Base64)]
        [ImporterHeader(Name = "图1")]
        public string Img1 { get; set; }

        [ImporterHeader(Name = "数值")]
        public string Number { get; set; }
        [ImporterHeader(Name = "名称")]
        public string Name { get; set; }
        [ImporterHeader(Name = "日期")]
        public DateTime Time { get; set; }

        /// <summary>
        /// 将图片导入到临时目录
        /// </summary>
        [ImportImageField(ImportImageTo = ImportImageTo.TempFolder)]
        [ImporterHeader(Name = "图")]
        public string Img { get; set; }
    }
}
