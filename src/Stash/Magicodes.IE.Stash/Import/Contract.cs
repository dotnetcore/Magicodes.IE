using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.IE.Stash.Import
{
    public class Contract
    {
        /// <summary>
        /// 默认注入的命名空间
        /// <para>TODO: 这里要改造一下,允许用户来管理自己需要默认注入的命名空间,以减少模型定义时的工作</para>
        /// </summary>
        public static readonly List<string> DefaultNameSpaces = new()
        {
            "System",
            "System.Text",
            "System.Linq",
            "System.Collections.Generic"
        };


        /// <summary>
        /// 数据文件内联定义的SheetName
        /// </summary>
        public static string InlineDefinitionSheetName { get; set; } = "$definition$";

        #region 将模板表中的定界符号提到到这里,以便用户自定义

        /// <summary>
        /// 模型类型
        /// </summary>
        public static string TokenForDtoType { get; set; } = "模型类型:";
        /// <summary>
        /// 命名空间
        /// </summary>
        public static string TokenForNamespaces { get; set; } = "命名空间:";
        /// <summary>
        /// 变量
        /// </summary>
        public static string TokenForVariable { get; set; } = "变量:";
        /// <summary>
        /// 序号
        /// </summary>
        public static string TokenForMapStart { get; set; } = "序号";
        /// <summary>
        /// 定义完
        /// </summary>
        public static string TokenForMapEnd { get; set; } = "^定义完^";
        /// <summary>
        /// 数据源列
        /// </summary>
        public static string TokenForMapColumn { get; set; } = "数据源列";
        /// <summary>
        /// 模型属性名
        /// </summary>
        public static string TokenForMapDtoProperty { get; set; } = "模型属性名";
        /// <summary>
        /// 默认值
        /// </summary>
        public static string TokenForMapDefaultValue { get; set; } = "默认值";
        /// <summary>
        /// 异常处理方式
        /// </summary>
        public static string TokenForMapFail { get; set; } = "异常处理方式";
        /// <summary>
        /// 转换器
        /// </summary>
        public static string TokenForMapPipe { get; set; } = "转换器";


        #endregion
    }
}
