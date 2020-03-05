using Magicodes.ExporterAndImporter.Core;
using System;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public class ImportPictureBase64Dto
    {
        [ImporterHeader(Name = "加粗文本")]
        public string Text { get; set; }
        [ImporterHeader(Name = "普通文本")]
        public string Text2 { get; set; }
        [ExcelImporterImg(IsImg = true,EnumImg =Core.Models.EnumImg.Base64)]
        [ImporterHeader(Name = "图1")]
        public string Img1 { get; set; }
        [ImporterHeader(Name = "数值")]
        public string Number { get; set; }
        [ImporterHeader(Name = "名称")]
        public string Name { get; set; }
        [ImporterHeader(Name = "日期")]
        public DateTime Time { get; set; }
        [ExcelImporterImg(IsImg = true)]
        [ImporterHeader(Name = "图")]
        public string Img { get; set; }
    }
}
