using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Setter generated for a source-generated property.
    /// </summary>
    public delegate bool XlsxGeneratedPropertySetter<T>(
        T item,
        string cell,
        IReadOnlyList<CellConverter>? converters,
        bool date1904);

    /// <summary>
    /// Read/write metadata generated for one property.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class XlsxGeneratedPropertyMetadata<T>
    {
        /// <summary>
        /// Gets the property name.
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// Gets the header name used for import matching.
        /// </summary>
        public string? ImportHeader { get; }
        /// <summary>
        /// Gets the display name used for the export header.
        /// </summary>
        public string? DisplayName { get; }
        /// <summary>
        /// Gets the description of the property.
        /// </summary>
        public string? Description { get; }
        /// <summary>
        /// Gets the number format string.
        /// </summary>
        public string? Format { get; }
        /// <summary>
        /// Gets the column width.
        /// </summary>
        public double? Width { get; }
        /// <summary>
        /// Gets the export ordering index.
        /// </summary>
        public int ExportIndex { get; }
        /// <summary>
        /// Gets a value indicating whether the property is excluded from export/import.
        /// </summary>
        public bool IsIgnored { get; }
        /// <summary>
        /// Gets the name of the target type.
        /// </summary>
        public string TargetTypeName { get; }
        /// <summary>
        /// Gets the generated value accessor.
        /// </summary>
        public Func<T, CellValue> Getter { get; }
        /// <summary>
        /// Gets the object-based value accessor.
        /// </summary>
        public Func<object?, CellValue> ObjectGetter { get; }
        /// <summary>
        /// Gets the generated cell writer, when one is available.
        /// </summary>
        public Action<XlsxWriter.XlsxRowWriter, T, int>? CellWriter { get; }
        /// <summary>
        /// Gets the generated property setter, when one is available.
        /// </summary>
        public XlsxGeneratedPropertySetter<T>? Setter { get; }

        public XlsxGeneratedPropertyMetadata(
            string name,
            string? importHeader,
            string? displayName,
            string? description,
            string? format,
            double? width,
            int exportIndex,
            bool isIgnored,
            string targetTypeName,
            Func<T, CellValue> getter,
            Action<XlsxWriter.XlsxRowWriter, T, int>? cellWriter,
            XlsxGeneratedPropertySetter<T>? setter)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            ImportHeader = importHeader;
            DisplayName = displayName;
            Description = description;
            Format = format;
            Width = width;
            ExportIndex = exportIndex;
            IsIgnored = isIgnored;
            TargetTypeName = targetTypeName ?? throw new ArgumentNullException(nameof(targetTypeName));
            Getter = getter ?? throw new ArgumentNullException(nameof(getter));
            ObjectGetter = item => item is T typed ? getter(typed) : CellValue.Null;
            CellWriter = cellWriter;
            Setter = setter;
        }
    }

    /// <summary>
    /// Registry for source-generated type metadata.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class XlsxGeneratedTypeMetadataRegistry
    {
        private static readonly ConcurrentDictionary<Type, object> _metadata = new();

        /// <summary>
        /// Registers metadata produced for <typeparamref name="T"/>.
        /// </summary>
        public static void Register<T>(Func<IReadOnlyList<XlsxGeneratedPropertyMetadata<T>>> factory)
        {
            if (factory is null) throw new ArgumentNullException(nameof(factory));
            var metadata = factory();
            if (metadata is null || metadata.Count == 0) return;
            _metadata[typeof(T)] = metadata;
        }

        /// <summary>
        /// Returns metadata for <typeparamref name="T"/>, if generated metadata was registered.
        /// </summary>
        public static IReadOnlyList<XlsxGeneratedPropertyMetadata<T>>? TryGet<T>()
        {
            return _metadata.TryGetValue(typeof(T), out var value)
                ? value as IReadOnlyList<XlsxGeneratedPropertyMetadata<T>>
                : null;
        }
    }

    /// <summary>
    /// Shared parsing helper used by generated readers.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class XlsxGeneratedReadHelper
    {
        /// <summary>
        /// Parses a cell value using the generated reader's conversion rules.
        /// </summary>
        public static bool TryParse<T>(
            string cell,
            out T value,
            IReadOnlyList<CellConverter>? converters,
            bool date1904)
        {
            if (!CellValueParser.TryParse(cell, typeof(T), out var parsed, converters, date1904))
            {
                value = default!;
                return false;
            }

            if (parsed is null)
            {
                value = default!;
                return true;
            }

            value = parsed is T typed ? typed : (T)parsed;
            return true;
        }
    }
}
