using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import.OneColumnImporter_Test
{
    /// <summary>
    /// </summary>
    [ExcelImporter(IsLabelingError = true)]
    public class OneColumnImporterDto
    {
        /// <summary>
        ///     姓名
        /// </summary>
        [ImporterHeader(Name = "姓名")]
        [Required(ErrorMessage = "学生姓名不能为空")]
        [MaxLength(50, ErrorMessage = "名称字数超出最大限制,请修改!")]
        public string Name { get; set; }
    }
}
