namespace Magicodes.IE.Stash.Import
{
    /// <summary>
    /// 变量项
    /// </summary>
    public class VariableItem
    {
        /// <summary>
        /// 变量名,变量可以被映射中其它地方引用.
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 值语句,最终会被解析为值
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 经过编译器处理过后的,给cli执行的代码.
        /// </summary>
        public string FullCode { get; set; }
        /// <summary>
        /// 编译后的变量值计算器委托
        /// </summary>
        [Newtonsoft.Json.JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public IFunc Func { get; set; }
        /// <summary>
        /// 经过计算之后的最终值
        /// </summary>
        [Newtonsoft.Json.JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public object Value { get; set; }
    }
}
