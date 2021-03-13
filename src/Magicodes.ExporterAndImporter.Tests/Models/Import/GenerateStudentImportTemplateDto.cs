using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{

    /// <summary>
    /// 导入列头筛选器测试
    /// 1）测试修改列头
    /// 2）测试修改值映射
    /// </summary>
    public class ImportStudentDtoHeaderFilterTest : IImportHeaderFilter
    {
        public List<ImporterHeaderInfo> Filter(List<ImporterHeaderInfo> importerHeaderInfos)
        {
            foreach (var item in importerHeaderInfos)
            {
                if (item.PropertyName == "Gender")
                {
                    item.MappingValues = new Dictionary<string, dynamic>()
                    {
                        {"男0",0 },
                        {"女1",1 }
                    };
                }
            }
            return importerHeaderInfos;
        }
    }

    [ExcelImporter(IsLabelingError = true, IsDisableAllFilter = false,
        ImportHeaderFilter = typeof(ImportStudentDtoHeaderFilterTest))]
    public class GenerateStudentImportTemplateDto
    {
        /// <summary>
        ///     序号
        /// </summary>
        [ImporterHeader(Name = "序号")]
        public long SerialNumber { get; set; }

        /// <summary>
        ///     学籍号
        /// </summary>
        [ImporterHeader(Name = "学籍号", IsAllowRepeat = false)]
        [MaxLength(30, ErrorMessage = "学籍号字数超出最大限制,请修改!")]
        public string StudentCode { get; set; }

        /// <summary>
        ///     姓名
        /// </summary>
        [ImporterHeader(Name = "姓名")]
        [Required(ErrorMessage = "学生姓名不能为空")]
        [MaxLength(50, ErrorMessage = "名称字数超出最大限制,请修改!")]
        public string Name { get; set; }

        /// <summary>
        ///     身份证号码
        /// </summary>
        [ImporterHeader(Name = "身份证号", IsAllowRepeat = false)]
        [Required(ErrorMessage = "身份证号不能为空")]
        [MaxLength(18, ErrorMessage = "身份证字数超出最大限制,请修改!")]
        public string IdCard { get; set; }

        /// <summary>
        ///     性别
        /// </summary>
        [ImporterHeader(Name = "性别")]
        [Required(ErrorMessage = "性别不能为空")]
        [ValueMapping("男", 0)]
        [ValueMapping("女", 1)]
        public Genders Gender { get; set; }

        /// <summary>
        ///     家庭地址
        /// </summary>
        [ImporterHeader(Name = "家庭住址")]
        [Required(ErrorMessage = "家庭地址不能为空")]
        [MaxLength(200, ErrorMessage = "家庭地址字数超出最大限制,请修改!")]
        public string Address { get; set; }

        /// <summary>
        ///     家长姓名
        /// </summary>
        [ImporterHeader(Name = "家长姓名")]
        [Required(ErrorMessage = "家长姓名不能为空")]
        [MaxLength(50, ErrorMessage = "家长姓名数超出最大限制,请修改!")]
        public string Guardian { get; set; }

        /// <summary>
        ///     家长联系电话
        /// </summary>
        [ImporterHeader(Name = "家长联系电话")]
        [MaxLength(20, ErrorMessage = "家长联系电话字数超出最大限制,请修改!")]
        public string GuardianPhone { get; set; }

        /// <summary>
        ///     学号
        /// </summary>
        [ImporterHeader(Name = "学号")]
        [MaxLength(30, ErrorMessage = "学号字数超出最大限制,请修改!")]
        public string StudentNub { get; set; }

        /// <summary>
        ///     宿舍号
        /// </summary>
        [ImporterHeader(Name = "宿舍号")]
        [MaxLength(20, ErrorMessage = "宿舍号字数超出最大限制,请修改!")]
        public string DormitoryNo { get; set; }

        /// <summary>
        ///     QQ
        /// </summary>
        [ImporterHeader(Name = "QQ号")]
        [MaxLength(30, ErrorMessage = "QQ号字数超出最大限制,请修改!")]
        public string QQ { get; set; }

        /// <summary>
        ///     民族
        /// </summary>
        [ImporterHeader(Name = "民族")]
        [MaxLength(2, ErrorMessage = "民族字数超出最大限制,请修改!")]
        public string Nation { get; set; }

        /// <summary>
        ///     户口性质
        /// </summary>
        [ImporterHeader(Name = "户口性质")]
        [MaxLength(10, ErrorMessage = "户口性质字数超出最大限制,请修改!")]
        public string HouseholdType { get; set; }

        /// <summary>
        ///     联系电话
        /// </summary>
        [ImporterHeader(Name = "学生联系电话")]
        [MaxLength(20, ErrorMessage = "手机号码字数超出最大限制,请修改!")]
        public string Phone { get; set; }

        /// <summary>
        ///     状态
        ///     测试可为空的枚举类型
        /// </summary>
        [ImporterHeader(Name = "状态")]
        public StudentStatus? Status { get; set; }

        /// <summary>
        ///     备注
        /// </summary>
        [ImporterHeader(Name = "备注")]
        [MaxLength(200, ErrorMessage = "备注字数超出最大限制,请修改!")]
        public string Remark { get; set; }

        /// <summary>
        ///     是否住校(宿舍)
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        public bool? IsBoarding { get; set; }

        /// <summary>
        ///     所属班级id
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        public Guid ClassId { get; set; }

        /// <summary>
        ///     学校Id
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        public Guid? SchoolId { get; set; }

        /// <summary>
        ///     校区Id
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        public Guid? CampusId { get; set; }

        /// <summary>
        ///     专业Id
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        public Guid? MajorsId { get; set; }

        /// <summary>
        ///     年级Id
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        public Guid? GradeId { get; set; }
    }


}
