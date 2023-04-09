using System;

namespace Magicodes.IE.Stash.Import
{
    /// <summary>
    /// 动态计算上下文
    /// </summary>

    //TODO: 这个上下文,感觉没啥意义,味道不太对
    public class ImportContext
    {
        public Type DtoType { get; set; }
        /// <summary>
        /// 已解析的变量
        /// </summary>
        public dynamic Variables { get; set; }

        /// <summary>
        /// 映射定义
        /// </summary>
        public ImportMapDefinition Def { get; set; }
    }

    /// <summary>
    /// 导入映射上下文
    /// </summary>
    public class ImportMapContext : ImportContext
    {
        /// <summary>
        /// 当前值
        /// </summary>
        public dynamic Value { get; set; }
        /// <summary>
        /// 当前映射项
        /// </summary>
        public MapItem Map { get; set; }
        /// <summary>
        /// 当前处理器的原始代码
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 转换Dto实例
        /// </summary>
        public object DtoObj { get; set; }
        /// <summary>
        /// 当前转换过程中单元格数据
        /// </summary>
        public List<CellValue> Cells { get; set; }
    }
}
