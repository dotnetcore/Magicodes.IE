using OfficeOpenXml.Table;

namespace OfficeOpenXml.LoadFunctions
{
    /// <summary>
    /// Base class for ExcelRangeBase.LoadFrom[...] functions
    /// </summary>
    internal abstract class LoadFunctionBase
    {
        public LoadFunctionBase(ExcelRangeBase range, LoadFunctionFunctionParamsBase parameters)
        {
            Range = range;
            PrintHeaders = parameters.PrintHeaders;
            TableStyle = parameters.TableStyle;
        }

        /// <summary>
        /// The range to which the data should be loaded
        /// </summary>
        protected ExcelRangeBase Range { get; }

        /// <summary>
        /// If true a header row will be printed above the data
        /// </summary>
        protected bool PrintHeaders { get; }

        /// <summary>
        /// If value is other than TableStyles.None the data will be added to a table in the worksheet.
        /// </summary>
        protected TableStyles TableStyle { get; set; }

        /// <summary>
        /// Returns how many rows there are in the range (header row not included)
        /// </summary>
        /// <returns></returns>
        protected abstract int GetNumberOfRows();

        /// <summary>
        /// Returns how many columns there are in the range
        /// </summary>
        /// <returns></returns>
        protected abstract int GetNumberOfColumns();

        protected abstract void LoadInternal(object[,] values);

        /// <summary>
        /// Loads the data into the worksheet
        /// </summary>
        /// <returns></returns>
        internal ExcelRangeBase Load()
        {
            var nRows = PrintHeaders ? GetNumberOfRows() + 1 : GetNumberOfRows();
            var nCols = GetNumberOfColumns();
            var values = new object[nRows, nCols];
            LoadInternal(values);
            var ws = Range.Worksheet;
            ws.SetRangeValueInner(Range._fromRow, Range._fromCol, Range._fromRow + nRows - 1, Range._fromCol + nCols - 1, values);

            //Must have at least 1 row, if header is shown
            if (nRows == 1 && PrintHeaders)
            {
                nRows++;
            }

            var r = ws.Cells[Range._fromRow, Range._fromCol, Range._fromRow + nRows - 1, Range._fromCol + nCols - 1];

            if (TableStyle != TableStyles.None)
            {
                var tbl = ws.Tables.Add(r, "");
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }

            return r;
        }
    }
}
