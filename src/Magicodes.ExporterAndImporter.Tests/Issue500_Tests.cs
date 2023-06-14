using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Newtonsoft.Json;
using OfficeOpenXml;
using Shouldly;
using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    /// <summary>
    /// 导入学生数据Dto
    /// IsLabelingError：是否标注数据错误
    /// </summary>
    [ExcelImporter(IsLabelingError = true, IsOnlyErrorRows = true, HeaderRowIndex = 2)]
    public class ImportWithOnlyError
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

    public class ExcelImporter_Tests_500 : TestBase
    {
        public ExcelImporter_Tests_500(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
        }

        private readonly ITestOutputHelper _testOutputHelper;
        public IExcelImporter Importer = new ExcelImporter();

        [Fact(DisplayName = "带导入说明行的 Excel 自动标注错误后导出文件里表头缺失")]
        public async Task Issue500_Test1()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入带描述头测试下.xlsx");

            var errorStream = new MemoryStream();
            var fs = new FileStream(filePath, FileMode.Open);

            var import = await Importer.Import<ImportWithOnlyError>(fs, errorStream);

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

            var filePath1 = GetTestFilePath($"{nameof(Issue500_Test1)}.xlsx");
            DeleteFile(filePath1);

            using (var fileStream = new FileStream(filePath1, FileMode.Create))
            {
                errorStream.Seek(0, SeekOrigin.Begin);
                await errorStream.CopyToAsync(fileStream);
            };

            using (var pck = new ExcelPackage(new FileInfo(filePath1)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells[1,1].Text.ShouldBe("序号");
                sheet.Cells[2, 1].Text.ShouldNotBeNullOrWhiteSpace();
            }
        }

        //[Fact(DisplayName = "异常流增加自定义异常后导出 fileByte 为 null")]
        //public async Task Issue998_Test()
        //{
        //    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入带描述头测试下.xlsx");

        //    var errorStream = new MemoryStream();
        //    var fs = new FileStream(filePath, FileMode.Open);

        //    var import = await Importer.Import<ImportWithOnlyError>(fs, errorStream);

        //    if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

        //    foreach (var item in import.Data.ToList())
        //    {
        //        var errorInfo = new DataRowErrorInfo()
        //        {
        //            //由于 Index 从开始
        //            RowIndex = import.Data.ToList().FindIndex(o => o.Equals(item)) + 3,
        //        };

        //        errorInfo.FieldErrors.Add("序号", "数据库已重复");

        //        import.RowErrors.Add(errorInfo);
        //    }

        //    bool result = Importer.OutputBussinessErrorData<ImportWithOnlyError>(errorStream, import.RowErrors.ToList(), out byte[] msg);

        //    msg.ShouldNotBeNull();
        //}
    }
}
