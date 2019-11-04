using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    /// <summary>
    /// 学生状态 正常、流失、休学、勤工俭学、顶岗实习、毕业、参军
    /// </summary>
    public enum StudentStatus
    {
        /// <summary>
        /// 正常
        /// </summary>
        [Display(Name = "正常")]
        Normal = 0,

        /// <summary>
        /// 流失
        /// </summary>
        [Description("流水")]
        PupilsAway = 1,

        /// <summary>
        /// 休学
        /// </summary>
        Suspension = 2,

        /// <summary>
        /// 勤工俭学
        /// </summary>
        WorkStudy = 3,

        /// <summary>
        /// 顶岗实习
        /// </summary>
        PostPractice = 4,

        /// <summary>
        /// 毕业
        /// </summary>
        Graduation = 5,

        /// <summary>
        /// 参军
        /// </summary>
        JoinTheArmy = 6,
    }
}
