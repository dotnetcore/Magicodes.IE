using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace Magicodes.Benchmarks.Models
{
    /// <summary>
    /// 导入学生数据Dto
    /// IsLabelingError：是否标注数据错误
    /// </summary>
    [ExcelImporter(IsLabelingError = true)]
    public class ImportStudentDto
    {
        /// <summary>
        ///     序号
        /// </summary>
        [ImporterHeader(Name = "序号")]
        [ExporterHeader("序号")]
        public long SerialNumber { get; set; }

        /// <summary>
        ///     学籍号
        /// </summary>
        [ImporterHeader(Name = "学籍号", IsAllowRepeat = false)]
        [ExporterHeader("学籍号")]
        public string StudentCode { get; set; }

        /// <summary>
        ///     姓名
        /// </summary>
        [ImporterHeader(Name = "姓名")]
        [ExporterHeader("姓名")]
        public string Name { get; set; }

        /// <summary>
        ///     身份证号码
        /// </summary>
        [ImporterHeader(Name = "身份证号", IsAllowRepeat = false)]
        [ExporterHeader("身份证号")]
        public string IdCard { get; set; }

        /// <summary>
        ///     联系电话
        /// </summary>
        [ImporterHeader(Name = "学生联系电话")]
        [ExporterHeader("学生联系电话")]
        public string Phone { get; set; }

    }
}
