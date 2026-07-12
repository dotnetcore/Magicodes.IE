
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// Configuration used to build an export.
    /// The profile is frozen before the first worksheet is written.
    /// </summary>
    public sealed class ExportProfile<T>
    {
        private string? _sheetName;
        private float? _fontSize;
        private bool _freezeHeader = true;
        private double? _defaultRowHeight;
        private bool _autoSst;
        private Action<ColumnMeta>? _headerFilter;
        private Func<T, bool>? _rowFilter;
        private List<string>? _mergeCells;
        private string? _autoFilterRef;
        private readonly List<(string Ref, string Uri)> _hyperlinks = new();
        private readonly Dictionary<string, ColumnConfig> _columns = new(StringComparer.Ordinal);
        private readonly HashSet<string> _ignored = new(StringComparer.Ordinal);
        private bool _frozen;

        /// <summary>
        /// Gets the name of the worksheet.
        /// </summary>
        public string? SheetName => _sheetName;
        /// <summary>
        /// Gets the column configurations, keyed by property name.
        /// </summary>
        public IReadOnlyDictionary<string, ColumnConfig> Columns => _columns;
        /// <summary>
        /// Gets the property names that are excluded from the export.
        /// </summary>
        public IReadOnlyCollection<string> Ignored => _ignored;
        /// <summary>
        /// Gets the default font size, in points, applied to all columns.
        /// </summary>
        public float? FontSize => _fontSize;
        /// <summary>
        /// Gets a value indicating whether the header row is frozen. The default is <see langword="true"/>.
        /// </summary>
        public bool FreezeHeader => _freezeHeader;
        /// <summary>
        /// Gets the default row height, in points.
        /// </summary>
        public double? DefaultRowHeight => _defaultRowHeight;
        /// <summary>
        /// Gets a value indicating whether the shared string table is enabled.
        /// </summary>
        public bool AutoSst => _autoSst;
        /// <summary>
        /// Gets the cell ranges that are merged.
        /// </summary>
        public IReadOnlyList<string>? MergeCells => _mergeCells;
        /// <summary>
        /// Gets the auto-filter range.
        /// </summary>
        public string? AutoFilterRef => _autoFilterRef;
        /// <summary>
        /// Gets the hyperlinks to write.
        /// </summary>
        public IReadOnlyList<(string Ref, string Uri)> Hyperlinks => _hyperlinks;
        /// <summary>
        /// Gets the callback invoked after the header columns are resolved, to further adjust column metadata.
        /// </summary>
        public Action<ColumnMeta>? HeaderFilter => _headerFilter;
        /// <summary>
        /// Gets the predicate that selects which rows are exported; rows for which it returns <see langword="false"/> are skipped.
        /// </summary>
        public Func<T, bool>? RowFilter => _rowFilter;

        /// <summary>
        /// Sets the worksheet name. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> Sheet(string name)
        {
            EnsureNotFrozen();
            _sheetName = name;
            return this;
        }

        /// <summary>
        /// Configures a single column. The <paramref name="selector"/> must be a direct property access, for example <c>x =&gt; x.Name</c>. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> Column<TProp>(Expression<Func<T, TProp>> selector, Action<ColumnConfig> configure)
        {
            EnsureNotFrozen();
            if (selector is null) throw new ArgumentNullException(nameof(selector));
            if (configure is null) throw new ArgumentNullException(nameof(configure));

            var propName = ExtractPropertyName(selector);
            var cfg = new ColumnConfig();
            configure(cfg);
            _columns[propName] = cfg;
            return this;
        }

        /// <summary>
        /// Excludes the specified property from the export. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> Ignore<TProp>(Expression<Func<T, TProp>> selector)
        {
            EnsureNotFrozen();
            if (selector is null) throw new ArgumentNullException(nameof(selector));
            var propName = ExtractPropertyName(selector);
            _ignored.Add(propName);
            return this;
        }

        /// <summary>
        /// Sets the default font size for all columns. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> WithFontSize(float size)
        {
            EnsureNotFrozen();
            _fontSize = size;
            return this;
        }

        /// <summary>
        /// Sets whether the header row is frozen. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> WithFreezeHeader(bool value = true)
        {
            EnsureNotFrozen();
            _freezeHeader = value;
            return this;
        }

        /// <summary>
        /// Sets the default row height. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> WithDefaultRowHeight(double height)
        {
            EnsureNotFrozen();
            _defaultRowHeight = height;
            return this;
        }

        /// <summary>
        /// Sets whether the shared string table is enabled. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> WithAutoSst(bool value = true)
        {
            EnsureNotFrozen();
            _autoSst = value;
            return this;
        }

        /// <summary>
        /// Adds a cell range to merge. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> WithMergeCells(string range)
        {
            EnsureNotFrozen();
            _mergeCells ??= new List<string>();
            _mergeCells.Add(range);
            return this;
        }

        /// <summary>
        /// Sets the auto-filter range. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> WithAutoFilter(string cellRange)
        {
            EnsureNotFrozen();
            _autoFilterRef = cellRange;
            return this;
        }

        /// <summary>
        /// Adds a hyperlink to a cell. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> WithHyperlink(string cellRef, string uri)
        {
            EnsureNotFrozen();
            _hyperlinks.Add((cellRef, uri));
            return this;
        }

        /// <summary>
        /// Registers a callback that adjusts the resolved header columns. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> WithHeader(Action<ColumnMeta> filter)
        {
            EnsureNotFrozen();
            _headerFilter = filter ?? throw new ArgumentNullException(nameof(filter));
            return this;
        }

        /// <summary>
        /// Exports only the rows that satisfy the predicate. Returns this instance for chaining.
        /// </summary>
        public ExportProfile<T> Where(Func<T, bool> predicate)
        {
            EnsureNotFrozen();
            _rowFilter = predicate ?? throw new ArgumentNullException(nameof(predicate));
            return this;
        }

        /// <summary>
        /// Freezes the configuration. Called by the write pipeline; once frozen, the profile cannot be modified.
        /// </summary>
        public void Freeze()
        {
            _frozen = true;
        }

        /// <summary>
        /// Gets a value indicating whether the configuration is frozen.
        /// </summary>
        public bool IsFrozen => _frozen;

        private void EnsureNotFrozen()
        {
            if (_frozen) throw new InvalidOperationException("ExportProfile is frozen and cannot be mutated. Complete configuration before freezing.");
        }

        private static string ExtractPropertyName<TProp>(Expression<Func<T, TProp>> selector)
        {

            if (selector.Body is MemberExpression m && m.Expression is ParameterExpression)
                return m.Member.Name;

            if (selector.Body is UnaryExpression { Operand: MemberExpression m2 } && m2.Expression is ParameterExpression)
                return m2.Member.Name;
            throw new ArgumentException("selector must be a direct property access, e.g. x => x.Foo", nameof(selector));
        }
    }

    /// <summary>
    /// Column accessors and styles prepared for the writer.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public class RowPlan
    {
        /// <summary>
        /// Columns included in the export.
        /// </summary>
        public ColumnMeta[] Columns { get; }
        /// <summary>
        /// Object-based accessors, in column order.
        /// </summary>
        public Func<object?, CellValue>[] Getters { get; }
        /// <summary>
        /// Style IDs, in column order.
        /// </summary>
        public int[] StyleIds { get; }

        /// <summary>
        /// Creates a row plan from the prepared column accessors and styles.
        /// </summary>
        public RowPlan(ColumnMeta[] columns, Func<object?, CellValue>[] getters, int[] styleIds)
        {
            Columns = columns; Getters = getters; StyleIds = styleIds;
        }

        public string[] BuildNumFmts()
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var result = new List<string>();
            for (int i = 0; i < Columns.Length; i++)
            {
                if (Columns[i].Format is string fmt && seen.Add(fmt))
                    result.Add(fmt);
            }
            return result.ToArray();
        }
    }

    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class TypedRowPlan<T> : RowPlan
    {
        /// <summary>
        /// Strongly typed accessors, in column order.
        /// </summary>
        public Func<T, CellValue>[] TypedGetters { get; }

        /// <summary>
        /// Generated cell writers, when available.
        /// </summary>
        public Action<XlsxWriter.XlsxRowWriter, T, int>?[] TypedCellWriters { get; }

        /// <summary>
        /// Indicates whether any column writes a formula.
        /// </summary>
        public bool HasFormulas { get; }

        /// <summary>
        /// Creates a strongly typed row plan.
        /// </summary>
        public TypedRowPlan(ColumnMeta[] columns, Func<object?, CellValue>[] objectGetters, Func<T, CellValue>[] typedGetters, int[] styleIds, Action<XlsxWriter.XlsxRowWriter, T, int>?[] typedCellWriters, bool hasFormulas)
            : base(columns, objectGetters, styleIds)
        {
            TypedGetters = typedGetters;
            TypedCellWriters = typedCellWriters;
            HasFormulas = hasFormulas;
        }
    }

    /// <summary>
    /// Builds and caches row plans for export profiles.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class RowPlanBuilder
    {
        private static class PlanCache<TPlan>
        {
            internal static readonly System.Runtime.CompilerServices.ConditionalWeakTable<ExportProfile<TPlan>, RowPlan> Untyped = new();
            internal static readonly System.Runtime.CompilerServices.ConditionalWeakTable<ExportProfile<TPlan>, TypedRowPlan<TPlan>> Typed = new();
        }

        private static readonly bool _isDynamicCodeAvailable =
#if NETSTANDARD2_0
            true;
#else
            System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeCompiled;
#endif

        /// <summary>
        /// Builds the reflection-based row plan for <typeparamref name="T"/>.
        /// </summary>
#if !NETSTANDARD2_0
        [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reflection over T's properties is intentional; user opts in by calling Build<T>")]
        [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Expression.Compile used for typed fast path; fallback to reflection path in AOT")]
#endif
        public static RowPlan Build<T>(ExportProfile<T> profile)
        {
            return PlanCache<T>.Untyped.GetValue(profile, static p => ReflectColumns<T>(p, typed: false));
        }

        /// <summary>
        /// Builds the strongly typed row plan for <typeparamref name="T"/>.
        /// </summary>
#if !NETSTANDARD2_0
        [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Reflection over T's properties is intentional; user opts in by calling BuildTyped<T>")]
        [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Expression.Compile used for typed fast path; fallback to reflection path in AOT")]
#endif
        public static TypedRowPlan<T> BuildTyped<T>(ExportProfile<T> profile)
        {
            return PlanCache<T>.Typed.GetValue(profile, static p => (TypedRowPlan<T>)ReflectColumns<T>(p, typed: true));
        }

        private static RowPlan ReflectColumns<T>(ExportProfile<T> profile, bool typed)
        {
            var generated = XlsxGeneratedTypeMetadataRegistry.TryGet<T>();
            if (generated is not null)
                return BuildGeneratedColumns(profile, generated, typed);

            var type = typeof(T);
            var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var columns = new List<ColumnMeta>(props.Length);
            var objectGetters = new List<Func<object?, CellValue>>(props.Length);
            var typedGetters = new List<Func<T, CellValue>>(props.Length);
            var propUnderlyings = new List<Type>(props.Length);
            var propInfos = new List<PropertyInfo>(props.Length);
            var styleIds = new List<int>(props.Length);
            bool hasFormulas = false;

            var formatToId = new Dictionary<string, int>(StringComparer.Ordinal);
            var ordered = new List<(PropertyInfo Prop, ColumnConfig? Cfg, int DeclOrder)>(props.Length);
            int declOrder = 0;
            foreach (var p in props)
            {
                if (!p.CanRead) continue;
                var name = p.Name;
                var exporterAttr = p.GetCustomAttribute<ExporterHeaderAttribute>(inherit: true);
                if (exporterAttr is { IsIgnore: true }) continue;
                if (profile.Ignored.Contains(name)) continue;
                profile.Columns.TryGetValue(name, out var cfg);
                ordered.Add((p, cfg, declOrder++));
            }
            ordered.Sort((a, b) =>
            {
                int ai = a.Cfg?.Index
                    ?? (a.Prop.GetCustomAttribute<ExporterHeaderAttribute>()?.Index is int ix && ix >= 0 ? ix : int.MaxValue);
                int bi = b.Cfg?.Index
                    ?? (b.Prop.GetCustomAttribute<ExporterHeaderAttribute>()?.Index is int iy && iy >= 0 ? iy : int.MaxValue);
                // Stable tiebreaker by declaration order — List.Sort is not stable, so without
                // this, columns sharing an Index (or all defaulting to int.MaxValue) would have
                // non-deterministic order across runs/platforms.
                int c = ai.CompareTo(bi);
                return c != 0 ? c : a.DeclOrder.CompareTo(b.DeclOrder);
            });

            int index = 0;
            foreach (var (p, cfg, _) in ordered)
            {
                var exporterAttr = p.GetCustomAttribute<ExporterHeaderAttribute>(inherit: true);

                var format = cfg?.Format ?? exporterAttr?.Format;
                if (format is null)
                {
                    var attr = p.GetCustomAttribute<DisplayFormatAttribute>(inherit: true);
                    if (attr is not null && !string.IsNullOrEmpty(attr.DataFormatString))
                        format = attr.DataFormatString;
                }

                int styleId = 0;
                if (format is not null)
                {
                    if (!formatToId.TryGetValue(format, out styleId))
                    {
                        styleId = formatToId.Count + 1;
                        formatToId[format] = styleId;
                    }
                }

                string displayName = cfg?.Name
                    ?? exporterAttr?.Name
                    ?? p.GetCustomAttribute<DisplayAttribute>(inherit: true)?.GetName()
                    ?? p.GetCustomAttribute<DescriptionAttribute>(inherit: true)?.Description
                    ?? p.Name;

                double? width = cfg?.Width;
                if (width is null && exporterAttr is { Width: > 0 })
                {
                    width = exporterAttr.Width;
                }

                var hidden = cfg?.Hidden ?? false;
                bool? bold = cfg?.Bold;
                bool? wrap = cfg?.Wrap;
                float? fontSize = cfg?.FontSize ?? profile.FontSize;
                bool? autoCenter = cfg?.AutoCenter;
                string? bgColor = cfg?.BackgroundColor;
                string? fontColor = cfg?.FontColor;
                string? fontName = cfg?.FontName;
                Magicodes.IE.IO.BorderStyle? borderStyle = cfg?.BorderStyle;
                string? borderColor = cfg?.BorderColor;
                bool? italic = cfg?.Italic;
                bool? underline = cfg?.Underline;
                bool? strikeThrough = cfg?.StrikeThrough;
                Magicodes.IE.IO.VerticalAlignment? verticalAlignment = cfg?.VerticalAlignment;
                double? rowHeight = cfg?.RowHeight;

                var meta = new ColumnMeta(p.Name, displayName, format, width, hidden, styleId, index,
                    bold: bold, wrap: wrap, fontSize: fontSize, autoCenter: autoCenter,
                    backgroundColor: bgColor, fontColor: fontColor, fontName: fontName,
                    borderStyle: borderStyle, borderColor: borderColor,
                    italic: italic, underline: underline, strikeThrough: strikeThrough,
                    verticalAlignment: verticalAlignment, rowHeight: rowHeight, formula: cfg?.Formula);
                profile.HeaderFilter?.Invoke(meta);

                columns.Add(meta);
                var (objGetter, tyGetter) = BuildGettersPair<T>(p);
                objectGetters.Add(objGetter);
                typedGetters.Add(tyGetter);
                propUnderlyings.Add(p.PropertyType);
                propInfos.Add(p);
                styleIds.Add(styleId);
                if (meta.Formula is not null) hasFormulas = true;
                index++;
            }

            if (typed)
            {
                var cellUnderlying = new Type[typedGetters.Count];
                for (int i = 0; i < cellUnderlying.Length; i++)
                    cellUnderlying[i] = Nullable.GetUnderlyingType(propUnderlyings[i]) ?? propUnderlyings[i];
                var cellWriters = BuildTypedCellWriters<T>(propInfos.ToArray(), cellUnderlying);
                return new TypedRowPlan<T>(
                    columns.ToArray(),
                    objectGetters.ToArray(),
                    typedGetters.ToArray(),
                    styleIds.ToArray(),
                    cellWriters,
                    hasFormulas);
            }
            return new RowPlan(columns.ToArray(), objectGetters.ToArray(), styleIds.ToArray());
        }

        private static RowPlan BuildGeneratedColumns<T>(
            ExportProfile<T> profile,
            IReadOnlyList<XlsxGeneratedPropertyMetadata<T>> properties,
            bool typed)
        {
            var ordered = new List<XlsxGeneratedPropertyMetadata<T>>(properties.Count);
            for (int i = 0; i < properties.Count; i++)
            {
                var property = properties[i];
                if (property.IsIgnored || profile.Ignored.Contains(property.Name))
                    continue;
                ordered.Add(property);
            }

            ordered.Sort((a, b) =>
            {
                int ai = profile.Columns.TryGetValue(a.Name, out var ac) && ac?.Index is int aix
                    ? aix
                    : a.ExportIndex >= 0 ? a.ExportIndex : int.MaxValue;
                int bi = profile.Columns.TryGetValue(b.Name, out var bc) && bc?.Index is int bix
                    ? bix
                    : b.ExportIndex >= 0 ? b.ExportIndex : int.MaxValue;
                return ai.CompareTo(bi);
            });

            var columns = new List<ColumnMeta>(ordered.Count);
            var objectGetters = new List<Func<object?, CellValue>>(ordered.Count);
            var typedGetters = new List<Func<T, CellValue>>(ordered.Count);
            var styleIds = new List<int>(ordered.Count);
            var formatToId = new Dictionary<string, int>(StringComparer.Ordinal);
            bool hasFormulas = false;

            for (int i = 0; i < ordered.Count; i++)
            {
                var property = ordered[i];
                profile.Columns.TryGetValue(property.Name, out var cfg);

                var format = cfg?.Format ?? property.Format;
                int styleId = 0;
                if (format is not null)
                {
                    if (!formatToId.TryGetValue(format, out styleId))
                    {
                        styleId = formatToId.Count + 1;
                        formatToId[format] = styleId;
                    }
                }

                var displayName = cfg?.Name ?? property.DisplayName ?? property.Description ?? property.Name;
                var width = cfg?.Width ?? property.Width;
                var meta = new ColumnMeta(
                    property.Name,
                    displayName,
                    format,
                    width,
                    cfg?.Hidden ?? false,
                    styleId,
                    i,
                    bold: cfg?.Bold,
                    wrap: cfg?.Wrap,
                    fontSize: cfg?.FontSize ?? profile.FontSize,
                    autoCenter: cfg?.AutoCenter,
                    backgroundColor: cfg?.BackgroundColor,
                    fontColor: cfg?.FontColor,
                    fontName: cfg?.FontName,
                    borderStyle: cfg?.BorderStyle,
                    borderColor: cfg?.BorderColor,
                    italic: cfg?.Italic,
                    underline: cfg?.Underline,
                    strikeThrough: cfg?.StrikeThrough,
                    verticalAlignment: cfg?.VerticalAlignment,
                    rowHeight: cfg?.RowHeight,
                    formula: cfg?.Formula);
                profile.HeaderFilter?.Invoke(meta);

                columns.Add(meta);
                objectGetters.Add(property.ObjectGetter);
                typedGetters.Add(property.Getter);
                styleIds.Add(styleId);
                if (meta.Formula is not null) hasFormulas = true;
            }

            if (!typed)
                return new RowPlan(columns.ToArray(), objectGetters.ToArray(), styleIds.ToArray());

            var cellWriters = new Action<XlsxWriter.XlsxRowWriter, T, int>?[ordered.Count];
            for (int i = 0; i < ordered.Count; i++)
                cellWriters[i] = ordered[i].CellWriter;

            return new TypedRowPlan<T>(
                columns.ToArray(),
                objectGetters.ToArray(),
                typedGetters.ToArray(),
                styleIds.ToArray(),
                cellWriters,
                hasFormulas);
        }

        private static Action<XlsxWriter.XlsxRowWriter, T, int>?[] BuildTypedCellWriters<T>(PropertyInfo[] props, Type[] underlyingTypes)
        {
            var generated = XlsxGeneratedRowWritersRegistry.TryGet<T>();
            if (generated is not null)
            {
                var arr = new Action<XlsxWriter.XlsxRowWriter, T, int>?[underlyingTypes.Length];
                for (int i = 0; i < underlyingTypes.Length; i++)
                {
                    if (generated.TryGetValue(props[i].Name, out var cw))
                        arr[i] = cw;
                }
                return arr;
            }

            int n = props.Length;
            var writers = new Action<XlsxWriter.XlsxRowWriter, T, int>?[n];
            for (int i = 0; i < n; i++)
            {
                var prop = props[i];
                var underlying = underlyingTypes[i];
                var nullableUnderlying = Nullable.GetUnderlyingType(prop.PropertyType);
                if (nullableUnderlying is not null)
                {
                    if (nullableUnderlying == typeof(int))
                    {
                        var getter = BuildTypedGetter<T, int?>(prop);
                        writers[i] = (w, item, sid) =>
                        {
                            var value = getter(item);
                            if (value.HasValue) w.WriteNumberCell(sid, value.Value);
                            else w.WriteEmptyCell(sid);
                        };
                    }
                    else if (nullableUnderlying == typeof(long))
                    {
                        var getter = BuildTypedGetter<T, long?>(prop);
                        writers[i] = (w, item, sid) =>
                        {
                            var value = getter(item);
                            if (value.HasValue) w.WriteNumberCell(sid, value.Value);
                            else w.WriteEmptyCell(sid);
                        };
                    }
                    else if (nullableUnderlying == typeof(double))
                    {
                        var getter = BuildTypedGetter<T, double?>(prop);
                        writers[i] = (w, item, sid) =>
                        {
                            var value = getter(item);
                            if (value.HasValue) w.WriteNumberCell(sid, value.Value);
                            else w.WriteEmptyCell(sid);
                        };
                    }
                    else if (nullableUnderlying == typeof(decimal))
                    {
                        var getter = BuildTypedGetter<T, decimal?>(prop);
                        writers[i] = (w, item, sid) =>
                        {
                            var value = getter(item);
                            if (value.HasValue) w.WriteNumberCell(sid, (double)value.Value);
                            else w.WriteEmptyCell(sid);
                        };
                    }
                    else if (nullableUnderlying == typeof(bool))
                    {
                        var getter = BuildTypedGetter<T, bool?>(prop);
                        writers[i] = (w, item, sid) =>
                        {
                            var value = getter(item);
                            if (value.HasValue) w.WriteBoolCell(sid, value.Value);
                            else w.WriteEmptyCell(sid);
                        };
                    }
                    else if (nullableUnderlying == typeof(DateTime))
                    {
                        var getter = BuildTypedGetter<T, DateTime?>(prop);
                        writers[i] = (w, item, sid) =>
                        {
                            var value = getter(item);
                            if (value.HasValue) w.WriteNumberCell(sid, value.Value.ToOADate());
                            else w.WriteEmptyCell(sid);
                        };
                    }
                    else if (nullableUnderlying == typeof(DateTimeOffset))
                    {
                        var getter = BuildTypedGetter<T, DateTimeOffset?>(prop);
                        writers[i] = (w, item, sid) =>
                        {
                            var value = getter(item);
                            if (value.HasValue) w.WriteStringCell(sid, value.Value.ToString("O", System.Globalization.CultureInfo.InvariantCulture));
                            else w.WriteEmptyCell(sid);
                        };
                    }
                    else if (nullableUnderlying.IsEnum)
                    {
                        var getter = BuildNullableEnumNumberGetter<T>(prop);
                        writers[i] = (w, item, sid) =>
                        {
                            var value = getter(item);
                            if (value.HasValue) w.WriteNumberCell(sid, value.Value);
                            else w.WriteEmptyCell(sid);
                        };
                    }
                }
                else if (underlying == typeof(string))
                {
                    var getter = BuildTypedGetter<T, string?>(prop);
                    writers[i] = (w, item, sid) =>
                    {
                        var value = getter(item);
                        if (value is null) w.WriteEmptyCell(sid);
                        else w.WriteStringCell(sid, value);
                    };
                }
                else if (underlying == typeof(int))
                {
                    var getter = BuildTypedGetter<T, int>(prop);
                    writers[i] = (w, item, sid) => w.WriteNumberCell(sid, getter(item));
                }
                else if (underlying == typeof(long))
                {
                    var getter = BuildTypedGetter<T, long>(prop);
                    writers[i] = (w, item, sid) => w.WriteNumberCell(sid, getter(item));
                }
                else if (underlying == typeof(double))
                {
                    var getter = BuildTypedGetter<T, double>(prop);
                    writers[i] = (w, item, sid) =>
                    {
                        w.WriteNumberCell(sid, getter(item));
                    };
                }
                else if (underlying == typeof(decimal))
                {
                    var getter = BuildTypedGetter<T, decimal>(prop);
                    writers[i] = (w, item, sid) =>
                    {
                        w.WriteNumberCell(sid, (double)getter(item));
                    };
                }
                else if (underlying == typeof(bool))
                {
                    var getter = BuildTypedGetter<T, bool>(prop);
                    writers[i] = (w, item, sid) =>
                    {
                        w.WriteBoolCell(sid, getter(item));
                    };
                }
                else if (underlying == typeof(DateTime))
                {
                    var getter = BuildTypedGetter<T, DateTime>(prop);
                    writers[i] = (w, item, sid) => w.WriteNumberCell(sid, getter(item).ToOADate());
                }
                else if (underlying == typeof(DateTimeOffset))
                {
                    var getter = BuildTypedGetter<T, DateTimeOffset>(prop);
                    writers[i] = (w, item, sid) => w.WriteStringCell(sid, getter(item).ToString("O", System.Globalization.CultureInfo.InvariantCulture));
                }
                else if (underlying.IsEnum)
                {
                    var getter = BuildEnumNumberGetter<T>(prop);
                    writers[i] = (w, item, sid) =>
                    {
                        w.WriteNumberCell(sid, getter(item));
                    };
                }
            }
            return writers;
        }

        private static (Func<object?, CellValue> objectGetter, Func<T, CellValue> typedGetter) BuildGettersPair<T>(PropertyInfo p)
        {
            var generated = XlsxGeneratedGettersRegistry.TryGet(typeof(T));
            if (generated is not null && generated.TryGetValue(p.Name, out var genGetter))
            {
                if (XlsxGeneratedTypedGettersRegistry.TryGet(typeof(T)) is { } typedReg
                    && typedReg.TryGetValue(p.Name, out var typedDel))
                {
                    return (genGetter, (Func<T, CellValue>)typedDel);
                }

                var typed = BuildTypedBridge<T>(genGetter);
                return (genGetter, typed);
            }

            var propertyType = p.PropertyType;
            var underlying = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
            var isNullableValueType = Nullable.GetUnderlyingType(propertyType) is not null;

            if (!_isDynamicCodeAvailable)
            {
                return BuildReflectionGettersPair<T>(p, underlying);
            }

            var objectCompiled = BuildExpressionGetter(p);
            Func<T, CellValue> typedGetter = default!;
            if (p.DeclaringType!.IsValueType)
            {
                var objFromT = BuildExpressionGetterBoxT<T>(p);
                typedGetter = item =>
                {
                    var boxed = objFromT(item);
                    return boxed is null ? CellValue.Null : ToCellValue(boxed, underlying);
                };
            }
            else if (isNullableValueType)
            {
                if (underlying == typeof(int))
                {
                    typedGetter = BuildNullableValueGetter<T, int>(p, v => CellValue.FromInteger(v));
                }
                else if (underlying == typeof(long))
                {
                    typedGetter = BuildNullableValueGetter<T, long>(p, CellValue.FromInteger);
                }
                else if (underlying == typeof(ulong))
                {
                    typedGetter = BuildNullableValueGetter<T, ulong>(p, v => CellValue.FromNumber((double)v));
                }
                else if (underlying == typeof(double))
                {
                    typedGetter = BuildNullableValueGetter<T, double>(p, CellValue.FromNumber);
                }
                else if (underlying == typeof(decimal))
                {
                    typedGetter = BuildNullableValueGetter<T, decimal>(p, v => CellValue.FromNumber((double)v));
                }
                else if (underlying == typeof(DateTime))
                {
                    typedGetter = BuildNullableValueGetter<T, DateTime>(p, CellValue.FromDateTime);
                }
                else if (underlying == typeof(bool))
                {
                    typedGetter = BuildNullableValueGetter<T, bool>(p, CellValue.FromBool);
                }
                else
                {
                    typedGetter = item =>
                    {
                        var boxed = objectCompiled(item);
                        return boxed is null ? CellValue.Null : ToCellValue(boxed, underlying);
                    };
                }
            }
            else if (underlying == typeof(string))
            {
                var strGetter = (Func<T, string?>)BuildExpressionGetterTyped<T>(p);
                typedGetter = item => strGetter(item) is { } s ? CellValue.FromString(s) : CellValue.Null;
            }
            else if (underlying == typeof(int))
            {
                var intGetter = (Func<T, int>)BuildExpressionGetterTyped<T>(p);
                typedGetter = item => CellValue.FromInteger(intGetter(item));
            }
            else if (underlying == typeof(long))
            {
                var lngGetter = (Func<T, long>)BuildExpressionGetterTyped<T>(p);
                typedGetter = item => CellValue.FromInteger(lngGetter(item));
            }
            else if (underlying == typeof(ulong))
            {
                var ulngGetter = (Func<T, ulong>)BuildExpressionGetterTyped<T>(p);
                typedGetter = item => CellValue.FromNumber((double)ulngGetter(item));
            }
            else if (underlying == typeof(double))
            {
                var dblGetter = (Func<T, double>)BuildExpressionGetterTyped<T>(p);
                typedGetter = item => CellValue.FromNumber(dblGetter(item));
            }
            else if (underlying == typeof(decimal))
            {
                var decGetter = (Func<T, decimal>)BuildExpressionGetterTyped<T>(p);
                typedGetter = item => CellValue.FromNumber((double)decGetter(item));
            }
            else if (underlying == typeof(DateTime))
            {
                var dtGetter = (Func<T, DateTime>)BuildExpressionGetterTyped<T>(p);
                typedGetter = item => CellValue.FromDateTime(dtGetter(item));
            }
            else if (underlying == typeof(bool))
            {
                var bGetter = (Func<T, bool>)BuildExpressionGetterTyped<T>(p);
                typedGetter = item => CellValue.FromBool(bGetter(item));
            }
            else if (underlying.IsEnum)
            {
                var getter = BuildEnumNumberGetter<T>(p);
                typedGetter = item => CellValue.FromNumber(getter(item));
            }
            else
            {
                typedGetter = item =>
                {
                    var boxed = objectCompiled(item);
                    return boxed is null ? CellValue.Null : ToCellValue(boxed, underlying);
                };
            }
            var objectGetter = (Func<object?, CellValue>)(item =>
            {
                var boxed = objectCompiled(item);
                if (boxed is null) return CellValue.Null;
                return ToCellValue(boxed, underlying);
            });
            return (objectGetter, typedGetter);
        }

        private static Func<object?, object?> BuildExpressionGetter(PropertyInfo p)
        {
            var itemParam = Expression.Parameter(typeof(object), "item");
            Expression body;
            if (p.DeclaringType!.IsValueType)
            {
                var unboxed = Expression.Unbox(itemParam, p.DeclaringType);
                var access = Expression.Property(unboxed, p);
                body = Expression.Convert(access, typeof(object));
            }
            else
            {
                var cast = Expression.Convert(itemParam, p.DeclaringType);
                var access = Expression.Property(cast, p);
                body = Expression.Convert(access, typeof(object));
            }
            return Expression.Lambda<Func<object?, object?>>(body, itemParam).Compile();
        }

        private static Func<T, object?> BuildExpressionGetterBoxT<T>(PropertyInfo p)
        {
            var itemParam = Expression.Parameter(typeof(T), "item");
            var access = Expression.Property(itemParam, p);
            var body = Expression.Convert(access, typeof(object));
            return Expression.Lambda<Func<T, object?>>(body, itemParam).Compile();
        }

        private static Func<T, CellValue> BuildNullableValueGetter<T, TValue>(PropertyInfo p, Func<TValue, CellValue> convert)
            where TValue : struct
        {
            var getter = (Func<T, TValue?>)BuildExpressionGetterTyped<T>(p);
            return item =>
            {
                var value = getter(item);
                return value.HasValue ? convert(value.Value) : CellValue.Null;
            };
        }

        private static Delegate BuildExpressionGetterTyped<T>(PropertyInfo p)
        {
            var itemParam = Expression.Parameter(typeof(T), "item");
            var access = Expression.Property(itemParam, p);
            return Expression.Lambda(access, itemParam).Compile();
        }

        private static Func<T, TValue> BuildTypedGetter<T, TValue>(PropertyInfo p)
        {
            var itemParam = Expression.Parameter(typeof(T), "item");
            var access = Expression.Property(itemParam, p);
            return Expression.Lambda<Func<T, TValue>>(access, itemParam).Compile();
        }

        private static (Func<object?, CellValue> objectGetter, Func<T, CellValue> typedGetter) BuildReflectionGettersPair<T>(PropertyInfo p, Type underlying)
        {
            Func<object?, CellValue> objectGetter = item =>
            {
                if (item is null) return CellValue.Null;
                object? boxed;
                try { boxed = p.GetValue(item); }
                catch (System.Reflection.TargetException) { return CellValue.Null; }
                return boxed is null ? CellValue.Null : ToCellValue(boxed, underlying);
            };
            Func<T, CellValue> typedGetter;
            if (typeof(T).IsValueType)
            {
                typedGetter = item =>
                {
                    object? boxed = item;
                    return objectGetter(boxed);
                };
            }
            else
            {
                typedGetter = item => objectGetter(item);
            }
            return (objectGetter, typedGetter);
        }

        private static Func<T, CellValue> BuildTypedBridge<T>(Func<object?, CellValue> objGetter)
        {
            if (typeof(T).IsValueType)
            {
                return item =>
                {
                    object? boxed = item;
                    return objGetter(boxed);
                };
            }
            return item => objGetter(item);
        }

        private static Func<T, double> BuildEnumNumberGetter<T>(PropertyInfo p)
        {
            var itemParam = Expression.Parameter(typeof(T), "item");
            var access = Expression.Property(itemParam, p);
            // Route enums through double: Excel stores numbers as IEEE754 double anyway, and
            // this avoids the (long) overflow for ulong-backed enums with values > long.MaxValue.
            var body = Expression.Convert(access, typeof(double));
            return Expression.Lambda<Func<T, double>>(body, itemParam).Compile();
        }

        private static Func<T, double?> BuildNullableEnumNumberGetter<T>(PropertyInfo p)
        {
            var itemParam = Expression.Parameter(typeof(T), "item");
            var access = Expression.Property(itemParam, p);
            var hasValue = Expression.Property(access, "HasValue");
            var value = Expression.Property(access, "Value");
            var doubleValue = Expression.Convert(value, typeof(double));
            var body = Expression.Condition(
                hasValue,
                Expression.Convert(doubleValue, typeof(double?)),
                Expression.Constant(null, typeof(double?)));
            return Expression.Lambda<Func<T, double?>>(body, itemParam).Compile();
        }

        internal static CellValue ToCellValue(object boxed, Type underlying)
        {
            if (underlying.IsEnum) return CellValue.FromNumber(Convert.ToDouble(boxed, System.Globalization.CultureInfo.InvariantCulture));
            if (underlying == typeof(string)) return CellValue.FromString((string)boxed);
            if (underlying == typeof(bool)) return CellValue.FromBool((bool)boxed);
            if (underlying == typeof(DateTime)) return CellValue.FromDateTime((DateTime)boxed);
            if (underlying == typeof(DateTimeOffset)) return CellValue.FromString(((DateTimeOffset)boxed).ToString("O", System.Globalization.CultureInfo.InvariantCulture));
            if (underlying == typeof(int) || underlying == typeof(long) || underlying == typeof(short) || underlying == typeof(byte))
                return CellValue.FromInteger(Convert.ToInt64(boxed, System.Globalization.CultureInfo.InvariantCulture));
            if (underlying == typeof(ulong))
                // Excel stores numbers as IEEE754 double; routing through double avoids the
                // (long) overflow for values > long.MaxValue.
                return CellValue.FromNumber((double)(ulong)boxed);
            if (underlying == typeof(uint) || underlying == typeof(ushort) || underlying == typeof(sbyte))
            {
                long l = underlying switch
                {
                    Type t when t == typeof(uint) => (long)(uint)boxed!,
                    Type t when t == typeof(ushort) => (long)(ushort)boxed!,
                    Type t when t == typeof(sbyte) => (long)(sbyte)boxed!,
                    _ => 0L,
                };
                return CellValue.FromInteger(l);
            }
            if (underlying == typeof(float) || underlying == typeof(double) || underlying == typeof(decimal))
                return CellValue.FromNumber(Convert.ToDouble(boxed, System.Globalization.CultureInfo.InvariantCulture));
            if (underlying == typeof(Guid)) return CellValue.FromString(boxed.ToString());
            return CellValue.FromString(boxed.ToString());
        }
    }
}
