namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class LookupValueMatcher : ValueMatcher
    {
        protected override int CompareObjectToString(object o1, string o2)
        {
            return IncompatibleOperands;
        }
    }
}
