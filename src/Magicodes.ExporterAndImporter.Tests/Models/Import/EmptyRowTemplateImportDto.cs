#if NET6_0_OR_GREATER
using System;
using System.ComponentModel.DataAnnotations;
using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    /// <summary>
    /// #582: 模板导入空行测试 DTO，空行导致错误行号偏移。
    /// </summary>
    public class EmptyRowTemplateImportDto
    {
        [ImporterHeader(Name = "姓名")]
        [Required]
        public string Name { get; set; }

        [ImporterHeader(Name = "年龄")]
        public int Age { get; set; }

        [ImporterHeader(Name = "城市")]
        public string City { get; set; }
    }
}
#endif