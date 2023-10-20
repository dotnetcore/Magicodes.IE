using OfficeOpenXml.LoadFunctions;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
#if !NET35 && !NET40
#endif
namespace OfficeOpenXml
{
    public partial class ExcelRangeBase
    {

        #region LoadFromDictionaries
        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items)
        {
            return LoadFromDictionaries(items, false, TableStyles.None, null);
        }

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, bool printHeaders)
        {
            return LoadFromDictionaries(items, printHeaders, TableStyles.None, null);
        }

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/></param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, bool printHeaders, TableStyles tableStyle)
        {
            return LoadFromDictionaries(items, printHeaders, tableStyle, null);
        }

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries</param>
        /// <param name="printHeaders">If true the key names from the first instance will be used as headers</param>
        /// <param name="tableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="keys">Keys that should be used, keys omitted will not be included</param>
        /// <returns>The filled range</returns>
        /// <example>
        /// <code>
        ///  var items = new List&lt;IDictionary&lt;string, object&gt;&gt;()
        ///    {
        ///        new Dictionary&lt;string, object&gt;()
        ///        { 
        ///            { "Id", 1 },
        ///            { "Name", "TestName 1" }
        ///        },
        ///        new Dictionary&lt;string, object&gt;()
        ///        {
        ///            { "Id", 2 },
        ///            { "Name", "TestName 2" }
        ///        }
        ///    };
        ///    using(var package = new ExcelPackage())
        ///    {
        ///        var sheet = package.Workbook.Worksheets.Add("test");
        ///        var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None, null);
        ///    }
        /// </code>
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, bool printHeaders, TableStyles tableStyle, IEnumerable<string> keys)
        {
            var param = new LoadFromDictionariesParams
            {
                PrintHeaders = printHeaders,
                TableStyle = tableStyle
            };
            if (keys != null && keys.Any())
            {
                param.SetKeys(keys.ToArray());
            }
            var func = new LoadFromDictionaries(this, items, param);
            return func.Load();
        }

        /// <summary>
        /// Load a collection of dictionaries (or dynamic/ExpandoObjects) into the worksheet starting from the top left row of the range.
        /// These dictionaries should have the same set of keys.
        /// </summary>
        /// <param name="items">A list of dictionaries/ExpandoObjects</param>
        /// <param name="paramsConfig"><see cref="Action{LoadFromDictionariesParams}"/> to provide parameters to the function</param>
        /// <example>
        /// sheet.Cells["C1"].LoadFromDictionaries(items, c =>
        /// {
        ///     c.PrintHeaders = true;
        ///     c.TableStyle = TableStyles.Dark1;
        /// });
        /// </example>
        public ExcelRangeBase LoadFromDictionaries(IEnumerable<IDictionary<string, object>> items, Action<LoadFromDictionariesParams> paramsConfig)
        {
            var param = new LoadFromDictionariesParams();
            paramsConfig.Invoke(param);
            var func = new LoadFromDictionaries(this, items, param);
            return func.Load();
        }
        #endregion
    }
}
