using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.IE.Tests.Models.Import;
/// <summary>
/// 可空枚举问题测试
/// </summary>
public class NullableEnumTestImportDto
{
    /// <summary>
    /// 编号
    /// </summary>
    [ImporterHeader(Name = "编号")]
    public string No { get; set; }
    /// <summary>
    /// 等级
    /// </summary>
    [ImporterHeader(Name = "等级")]
    public EnumTest? Level { get; set; }
}
public enum EnumTest
{
    A,
    B,
}