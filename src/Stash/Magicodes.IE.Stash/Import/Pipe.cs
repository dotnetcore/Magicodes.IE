namespace Magicodes.IE.Stash.Import
{
    public class Pipe
    {
        public string Code { get; set; }
        public string FullCode { get; set; }
        [Newtonsoft.Json.JsonIgnore]
        [System.Text.Json.Serialization.JsonIgnore]
        public IFunc Func { get; set; }
    }
}
