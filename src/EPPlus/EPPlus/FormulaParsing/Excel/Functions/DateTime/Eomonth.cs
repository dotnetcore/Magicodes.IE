using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Eomonth : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var date = System.DateTime.FromOADate(ArgToDecimal(arguments, 0));
            var monthsToAdd = ArgToInt(arguments, 1);
            var resultDate = new System.DateTime(date.Year, date.Month, 1).AddMonths(monthsToAdd + 1).AddDays(-1);
            return CreateResult(resultDate.ToOADate(), DataType.Date);
        }
    }
}
