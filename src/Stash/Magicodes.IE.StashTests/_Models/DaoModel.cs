using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.IE.StashTests._Models
{
    public class DaoModel
    {
        //TODO: 还没搞子属性
        public DateTime? NullableTime { get; set; }
        public DateTime Time { get; set; }
        public int Int { get; set; }
        public int? NullableInt { get; set; }
        public string String { get; set; }
        public string String2 { get; set; }
        public string String3 { get; set; }
        public bool Bool { get; set; }
        public bool? NullableBool { get; set; }
        public decimal Decimal { get; set; }
        public decimal? NullableDecimal { get; set; }
        public double Double { get; set; }
        public double? NullableDouble { get; set; }
        public float Float { get; set; }
        public float? NullableFloat { get; set; }
        public long Long { get; set; }
        public long? NullableLong { get; set; }
        public short Short { get; set; }
        public short? NullableShort { get; set; }
        public byte Byte { get; set; }
        public byte? NullableByte { get; set; }
        public Guid Guid { get; set; }
        public Guid? NullableGuid { get; set; }
    }
}
