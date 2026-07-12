using System;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Base class for cell converters. Most converters should derive from <see cref="CellConverter{T}"/> instead.
    /// </summary>
    public abstract class CellConverter
    {
        /// <summary>
        /// Gets the <see cref="Type"/> this converter handles.
        /// </summary>
        public abstract Type Type { get; }

        internal virtual bool TryRead(string cell, out object? value)
        {
            value = null;
            return false;
        }
    }

    /// <summary>
    /// Converts the text of a cell into a value of type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    public abstract class CellConverter<T> : CellConverter
    {
        /// <summary>
        /// Attempts to parse <paramref name="cell"/> into a <typeparamref name="T"/>. Returns <see langword="false"/> if the value cannot be parsed.
        /// </summary>
        public abstract bool Read(string cell, out T value);

        internal sealed override bool TryRead(string cell, out object? value)
        {
            var result = Read(cell, out var typedValue);
            value = typedValue;
            return result;
        }

        public sealed override Type Type => typeof(T);
    }
}
