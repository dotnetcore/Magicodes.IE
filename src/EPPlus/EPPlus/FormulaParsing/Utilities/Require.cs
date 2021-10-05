namespace OfficeOpenXml.FormulaParsing.Utilities
{
    public static class Require
    {
        public static ArgumentInfo<T> That<T>(T arg)
        {
            return new ArgumentInfo<T>(arg);
        }
    }
}
