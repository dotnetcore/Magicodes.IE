using Newtonsoft.Json;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using System.Reflection;
using System;

namespace Magicodes.IE.Stash.Import
{
    /// <summary>
    /// 导入映射定义
    /// </summary>
    public class ImportMapDefinition
    {
        /// <summary>
        /// 数据模型类型名称
        /// <para>通过反射找到对应的模型,然后反射实例化和赋值</para>
        /// </summary>
        public string DtoTypeName { get; set; }
        /// <summary>
        /// Dto的Clr类型
        /// </summary>
        public Type DtoType { get; set; }
        /// <summary>
        /// 引入的命名空间
        /// <para>默认引用 System \ System.Text \System.Linq \ System.Collections.Generic</para>
        /// </summary>
        public List<string> Namespaces { get; set; } = new();
        /// <summary>
        /// 变量定义,变量可以被映射定义中的其它地方引用.变更可以是常量或可计算语句
        /// </summary>
        public List<VariableItem> Variables { get; set; } = new();
        /// <summary>
        /// 数据映射定义
        /// </summary>
        public List<MapItem> Maps { get; set; } = new();
    }
}
