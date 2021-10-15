using OfficeOpenXml.Table;

namespace OfficeOpenXml.LoadFunctions
{
    public abstract class LoadFunctionFunctionParamsBase
    {
        /// <summary>
        /// If true a row with headers will be added above the data
        /// </summary>
        public bool PrintHeaders
        {
            get; set;
        }

        /// <summary>
        /// If set to another value than TableStyles.None the data will be added to a
        /// table with the specified style
        /// </summary>
        public TableStyles TableStyle
        {
            get; set;
        }
    }
}
