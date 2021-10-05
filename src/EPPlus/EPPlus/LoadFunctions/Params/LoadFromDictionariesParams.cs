using System.Collections.Generic;

namespace OfficeOpenXml.LoadFunctions.Params
{
    /// <summary>
    /// Parameters for the LoadFromDictionaries method
    /// </summary>
    public class LoadFromDictionariesParams : LoadFunctionFunctionParamsBase
    {
        /// <summary>
        /// If set, only these keys will be included in the dataset
        /// </summary>
        public IEnumerable<string> Keys { get; private set; }

        /// <summary>
        /// The keys supplied to this function will be included in the dataset, all others will be ignored.
        /// </summary>
        /// <param name="keys">The keys to include</param>
        public void SetKeys(params string[] keys)
        {
            Keys = keys;
        }

        /// <summary>
        /// Sets how headers should be parsed before added to the worksheet, see <see cref="HeaderParsingTypes"/>
        /// </summary>
        public HeaderParsingTypes HeaderParsingType { get; set; } = HeaderParsingTypes.UnderscoreToSpace;
    }
}
