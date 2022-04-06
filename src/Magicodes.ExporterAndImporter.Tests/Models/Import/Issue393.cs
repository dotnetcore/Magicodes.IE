using Magicodes.ExporterAndImporter.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.IE.Tests.Models.Import
{
    public class Issue393
    {
        [Required(ErrorMessage = "中心编号不能为空")]
        [ImporterHeader(Name = "中心编号", AutoTrim = true)]
        public string SiteCode { get; set; } = string.Empty;


        [Required(ErrorMessage = "受试者筛选号不能为空")]
        [ImporterHeader(Name = "受试者筛选号", AutoTrim = true)]
        public string SubjectCode { get; set; } = string.Empty;

        [Required(ErrorMessage = "访视名称不能为空")]
        [ImporterHeader(Name = "访视名称", AutoTrim = true)]
        public string VisitName { get; set; } = string.Empty;

        [Required(ErrorMessage = "检查日期不能为空")]
        //[CanConvertToTime(ErrorMessage = "检查日期格式有问题")]
        [ImporterHeader(Name = "检查日期", AutoTrim = true)]
        public string StudyDate { get; set; } = string.Empty;

        [Required(ErrorMessage = "Modality不能为空")]
        [ImporterHeader(Name = "检查技术", AutoTrim = true)]
        public string Modality { get; set; } = string.Empty;
    }
}
