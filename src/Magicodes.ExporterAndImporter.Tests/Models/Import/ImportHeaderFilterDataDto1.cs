// ======================================================================
// 
//           filename : ImportStudentDto.cs
//           description :
// 
//           created by 雪雁 at  2019-11-05 20:02
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{

    /// <summary>
    /// 导入列头筛选器测试
    /// 1）测试修改列头
    /// 2）测试修改值映射
    /// </summary>
    public class ImportHeaderFilterTest : IImportHeaderFilter
    {
        public List<ImporterHeaderInfo> Filter(List<ImporterHeaderInfo> importerHeaderInfos)
        {
            foreach (var item in importerHeaderInfos)
            {
                if (item.PropertyName == "Name")
                {
                    item.Header.Name = "Student";
                }
                else if (item.PropertyName == "Gender")
                {
                    item.MappingValues = new Dictionary<string, dynamic>()
                    {
                        {"男",0 },
                        {"女",1 }
                    };
                }
            }
            return importerHeaderInfos;
        }
    }

    /// <summary>
    /// 导入学生数据Dto
    /// IsLabelingError：是否标注数据错误
    /// </summary>
    [ExcelImporter(IsLabelingError = true, ImportHeaderFilter = typeof(ImportHeaderFilterTest))]
    public class ImportHeaderFilterDataDto1
    {

        /// <summary>
        ///     姓名
        /// </summary>
        [ImporterHeader(Name = "姓名", Author = "雪雁")]
        [Required(ErrorMessage = "学生姓名不能为空")]
        [MaxLength(50, ErrorMessage = "名称字数超出最大限制,请修改!")]
        public string Name { get; set; }

        /// <summary>
        ///     性别
        /// </summary>
        [ImporterHeader(Name = "性别")]
        [Required(ErrorMessage = "性别不能为空")]
        public Genders Gender { get; set; }

    }

    /// <summary>
    /// 导入学生数据Dto
    /// IsLabelingError：是否标注数据错误
    /// </summary>
    [ExcelImporter(IsLabelingError = true)]
    public class DIImportHeaderFilterDataDto1
    {

        /// <summary>
        ///     姓名
        /// </summary>
        [ImporterHeader(Name = "姓名", Author = "雪雁")]
        [Required(ErrorMessage = "学生姓名不能为空")]
        [MaxLength(50, ErrorMessage = "名称字数超出最大限制,请修改!")]
        public string Name { get; set; }

        /// <summary>
        ///     性别
        /// </summary>
        [ImporterHeader(Name = "性别")]
        [Required(ErrorMessage = "性别不能为空")]
        public Genders Gender { get; set; }

    }
}