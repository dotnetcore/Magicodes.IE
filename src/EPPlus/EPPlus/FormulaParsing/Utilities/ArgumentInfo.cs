namespace OfficeOpenXml.FormulaParsing.Utilities
{
    public class ArgumentInfo<T>
    {
        public ArgumentInfo(T val)
        {
            Value = val;
        }

        public T Value { get; private set; }

        public string Name { get; private set; }

        public ArgumentInfo<T> Named(string argName)
        {
            Name = argName;
            return this;
        }
    }
}
