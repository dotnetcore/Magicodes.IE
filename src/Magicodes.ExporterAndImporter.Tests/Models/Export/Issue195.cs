using Magicodes.ExporterAndImporter.Core;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    class Issue195
    {
        [ExporterHeader("分数")]
        public int Score { get; set; }

        [ExporterHeader(DisplayName = "性别")]
        public Sex Sex { get; set; }
    }
    enum Sex
    {
        /// <summary>
        /// 男
        /// </summary>
        [Display(Name = "男")]
        [Description("男")]
        boy = 1,
        /// <summary>
        /// 女
        /// </summary>
        [Display(Name = "女")]
        [Description("女")]
        girl = 2
    }


}
