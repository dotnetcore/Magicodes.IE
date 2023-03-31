namespace Magicodes.JustCode.Import
{
    /// <summary>
    /// 变量项
    /// </summary>
    public class VariableItem
    {
        /// <summary>
        /// 变量名,变更可以被映射中其它地方引用,引用时用双花括号定界,如: {{变量名}}
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 值定义,可以是常量,也可以是C#语句
        /// </summary>
        public string Code { get; set; }
        public string FullCode { get; set; }
        [Newtonsoft.Json.JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public IFunc Func { get; set; }
        /// <summary>
        /// 解析值定义之后的最终值
        /// </summary>
        [Newtonsoft.Json.JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public object Value { get; set; }
    }
}
