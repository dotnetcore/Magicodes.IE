using System.Collections.Generic;

namespace Magicodes.IE.Stash.Import
{
    /// <summary>
    /// 映射项
    /// </summary>
    public class MapItem
    {
        /// <summary>
        /// 序号
        /// </summary>
        public string Index { get; set; }
        /// <summary>
        /// 数据源列寻址,标题\列号\全部\未定义之外的所有列
        /// </summary>
        public string Column { get; set; }
        /// <summary>
        /// 数据类型
        /// </summary>
        public string Type { get; set; }
        /// <summary>
        /// Dto属性名
        /// </summary>
        public string Property { get; set; }
        /// <summary>
        /// 默认值
        /// </summary>
        public object Default { get; set; }
        /// <summary>
        /// 读取或转换失败后的处理方式
        /// </summary>
        public string Fail { get; set; }
        /// <summary>
        /// 数据转换器管道,按顺序依次执行
        /// </summary>
        public List<Pipe> Pipes { get; set; } = new();
    }
}
