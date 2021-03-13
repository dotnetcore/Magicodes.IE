using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using OfficeOpenXml.Table;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{

    /// <summary>
    /// 导入管轴模型
    /// </summary>
    [ExcelImporter(IsLabelingError = true, ImportDescription = @"导入说明：
    1、* 代表必填
    2、管轴编号严格区分大小写且具备固定格式，不得使用特殊字符，如果是数字编号轴则10以下的数值必须使用0补位，示例：B-01、B-02、P1-01、P1-15a、P1-30New等
    3、管线编号严格区分大小写，一般使用大写字母，比如：01-150-HW-05A111-A1E-H等，且不得使用除[-]以外的特殊字符，系统不存在该管线编号时将会导入失败。
    4、层级填写内容为数值（整数）格式，剖面在该管廊的第几层，示例：1、2、3、4、5、-1等。
    5、序号填写内容为数值（整数）格式，剖面在该层级的排列顺序，示例：1、2、3、4、5、6、7等。
    6、位置标识填写内容为文本格式，示例：左、中、右。
    7、是否为吊架填写内容为文本格式，示例：是、否。
    8、吊架百分比填写内容为文本格式，示例：*、50%，20%。"
, DescriptionHeight = 200)]
    [ExcelExporter(Name = "管轴模型", TableStyle = TableStyles.Light10, AutoFitAllColumn = true)]

    public class ImportGalleryAxisDto
    {

        ///<summary>
        ///管轴编号
        ///</summary>
        [ImporterHeader(Name = "管轴编号", IsAllowRepeat = false, Description = "仅支持大写字母,格式如下A-01,数字请至少保证有两位,如01")]
        [MaxLength(length: 100, ErrorMessage = "管轴编号字数超出最大限制,请修改!")]
        [Required(ErrorMessage = "管轴编号不能为空")]
        [ExporterHeader(DisplayName = "管轴编号")]
        public string AxisNumber { get; set; } = string.Empty;

        ///<summary>
        ///管廊编号
        ///</summary>
        [ImporterHeader(Name = "管廊编号", Description = "仅支持大写字母,请在系统中确认该管廊编号已存在,如 A ")]
        [Required(ErrorMessage = "管廊编号不能为空")]
        [ExporterHeader(DisplayName = "管廊编号")]
        public string GalleryNumber { get; set; }

        /// <summary>
        /// 责任区域
        /// </summary>
        [ImporterHeader(Name = "责任区域", Description = "请输入在系统中确认的责任区域")]
        [Required(ErrorMessage = "责任区域不能为空")]
        [ExporterHeader(DisplayName = "责任区域")]
        public string DutyRegion { get; set; }

        ///<summary>
        ///位置标识
        ///</summary>
        [ImporterHeader(Name = "位置标识")]
        [MaxLength(length: 20, ErrorMessage = "位置标识超出最大限制,请修改!")]
        [ValueMapping("左", "左")]
        [ValueMapping("中", "中")]
        [ValueMapping("右", "右")]
        [ExporterHeader(DisplayName = "位置标识")]
        public string Position { get; set; } = string.Empty;

        ///<summary>
        ///是否为吊架
        ///</summary>
        [ImporterHeader(Name = "是否为吊架", Description = "")]
        [ValueMapping("是", "是")]
        [ValueMapping("否", "否")]
        [ExporterHeader(DisplayName = "是否为吊架")]
        public string IsCradle { get; set; }

        ///<summary>
        ///百分比
        ///</summary>
        [ImporterHeader(Name = "吊架百分比")]
        [MaxLength(length: 50, ErrorMessage = "百分比长度超出最大限制,请修改!")]
        [ExporterHeader(DisplayName = "百分比")]
        public string Percentage { get; set; } = string.Empty;

        ///<summary>
        ///排序
        ///</summary>
        [ImporterHeader(Name = "排序", Description = "请输入数字排序")]
        [ExporterHeader(DisplayName = "排序")]
        public int? OrderBy { get; set; }

        /// <summary>
        /// 导入错误反馈
        /// </summary>
        [ImporterHeader(Name = " 导入错误反馈", IsIgnore = true)]
        [ExporterHeader(DisplayName = "导入错误反馈")]
        public string ExportRemark { get; set; }
    }
}
