using Magicodes.ExporterAndImporter.Core;
using Magicodes.IE.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Magicodes.IE.Tests.Models.Import
{
    public class DynamicStringLengthImportDto
    {
        [ImporterHeader(Name = "名称")]
        [Required(ErrorMessage = "名称不能为空")]
        [DynamicStringLength(typeof(DynamicStringLengthImportDtoConsts), nameof(DynamicStringLengthImportDtoConsts.MaxNameLength), ErrorMessage = "名称字数不能超过{1}")]
        public string Name { get; set; }
    }

    public static class DynamicStringLengthImportDtoConsts
    {
        public static int MaxNameLength { get; set; } = 3;
    }
}
