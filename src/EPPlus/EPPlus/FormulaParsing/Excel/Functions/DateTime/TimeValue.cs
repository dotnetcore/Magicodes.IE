using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    /// <summary>
    /// Simple implementation of TimeValue function, just using .NET built-in
    /// function System.DateTime.TryParse, based on current culture
    /// </summary>
    public class TimeValue : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var dateString = ArgToString(arguments, 0);
            return Execute(dateString);
        }

        internal CompileResult Execute(string dateString)
        {
            System.DateTime result;
            System.DateTime.TryParse(dateString, out result);
            return result != System.DateTime.MinValue ?
                CreateResult(GetTimeValue(result), DataType.Date) :
                CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
        }

        private double GetTimeValue(System.DateTime result)
        {
            return (int)result.TimeOfDay.TotalSeconds == 0 ? 0d : result.TimeOfDay.TotalSeconds / (3600 * 24);
        }
    }
}
