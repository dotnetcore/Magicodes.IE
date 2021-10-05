using System;

namespace OfficeOpenXml.FormulaParsing.Utilities
{
    public class IntegerIdProvider : IdProvider
    {
        private int _lastId = int.MinValue;

        public override object NewId()
        {
            if (_lastId >= int.MaxValue)
            {
                throw new InvalidOperationException("IdProvider run out of id:s");
            }
            return _lastId++;
        }
    }
}
