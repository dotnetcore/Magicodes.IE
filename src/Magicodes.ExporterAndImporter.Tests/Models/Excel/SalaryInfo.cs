using Magicodes.ExporterAndImporter.Core;
using System;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Excel
{
    public class SalaryInfo
    {
        [ImporterHeader(Name = "工资月份")]
        [ExporterHeader(DisplayName = "工资月份", Format = "yyyy-MM-dd")]
        [Required]
        public DateTime SalaryDate { get; set; }

        /// <summary>
        /// 员工姓名
        /// </summary>
        [ImporterHeader(Name = "员工姓名")]
        [ExporterHeader(DisplayName = "员工姓名")]
        [Required]
        [MaxLength(50, ErrorMessage = "员工姓名字数超过最大长度50的限制")]
        public string EmpName { get; set; }

        /// <summary>
        /// 岗级工资
        /// </summary>
        [ImporterHeader(Name = "岗级工资")]
        [ExporterHeader(DisplayName = "岗级工资")]
        [Required]
        public decimal PostSalary { get; set; }

        public DateTime? TestNullDate1 { get; set; }

        [ExporterHeader(DisplayName = "TestNullDate2", Format = "yyyy-MM-dd")]
        public DateTime? TestNullDate2 { get; set; }

        public DateTimeOffset TestDateTimeOffset1 { get; set; }

        public DateTimeOffset? TestDateTimeOffset2 { get; set; }
    }
}
