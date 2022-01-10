using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Fact : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToDecimal(arguments, 0);
            ThrowExcelErrorValueExceptionIf(() => number < 0, eErrorType.NA);
            var result = 1d;
            for (var x = 1; x < number; x++)
            {
                result *= x;
            }
            return CreateResult(result, DataType.Integer);
        }
    }
}
