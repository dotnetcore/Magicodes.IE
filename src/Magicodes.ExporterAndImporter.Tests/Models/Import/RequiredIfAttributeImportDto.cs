using Magicodes.ExporterAndImporter.Core;
using Magicodes.IE.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Magicodes.IE.Tests.Models.Import
{
    public class RequiredIfAttributeImportDto
    {
        [ImporterHeader(Name = "名称是否必填")]
        [Required(ErrorMessage = "名称是否必填不能为空")]
        [ValueMapping("是", true)]
        [ValueMapping("否", false)]
        public bool IsNameRequired { get; set; }

        [ImporterHeader(Name = "名称")]
        [RequiredIf("IsNameRequired", "True", ErrorMessage = "名称不能为空")]
        [MaxLength(10, ErrorMessage = "名称字数超出最大值：10")]
        public string Name { get; set; }
    }
}
