using Magicodes.ExporterAndImporter.Excel;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    [ExcelImporter(IsLabelingError = true, HeaderRowIndex = 2)]
    public class Issue236
    {
        public string 标题 { get; set; }

        public string 外链 { get; set; }

        public string 封面图 { get; set; }

        public string 标签 { get; set; }

        public string 摘要 { get; set; }

        public string 内容 { get; set; }

        public string 浏览次数 { get; set; }

        public string 排序 { get; set; }

        [Required(ErrorMessage = "分类ID是必填的")]
        public string 分类ID { get; set; }

        public string 创建时间 { get; set; }

    }
}
