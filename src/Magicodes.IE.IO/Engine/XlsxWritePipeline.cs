
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace Magicodes.IE.IO
{
    internal static class XlsxWritePipeline
    {

        public static void Run<T>(Stream output, IEnumerable<T> data, Action<ExportProfile<T>>? configure, XlsxWriteOptions? options)
        {
            var profile = new ExportProfile<T>();
            configure?.Invoke(profile);
            Run(output, data, profile, options);
        }

        public static void Run<T>(Stream output, IEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options)
        {
            var compression = options?.Compression ?? CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;

            profile.Freeze();

            var plan = RowPlanBuilder.BuildTyped(profile);
            var actualSheetName = profile.SheetName ?? typeof(T).Name;

            IEnumerable<T> filtered = profile.RowFilter is null ? data : FilterRows(data, profile.RowFilter);

            bool autoSst = profile.AutoSst;
            // SST needs a short probe before the writer is created; after that the sequence is
            // replayed so the probe does not consume rows from the actual export.
            if (autoSst && compression != CompressionLevel.NoCompression)
                filtered = AutoSstBufferAndDetect(filtered, plan, out autoSst);
            else if (autoSst)
                autoSst = false;

            var numFmts = plan.BuildNumFmts();

            using var writer = new XlsxWriter(output, actualSheetName, compression, defaultRowHeight: 0, strictCellReferences);
            if (profile.DefaultRowHeight is double dh)
                writer.SetDefaultRowHeight(dh);
            if (autoSst) writer.EnableSharedStrings();
            ApplySheetFeatures(writer, profile);
            writer.SetNumFmts(numFmts);
            writer.ResolveColumnStyles(plan.Columns);
            writer.WriteSheetMeta(plan.Columns, profile.FreezeHeader);
            writer.WriteHeader(plan.Columns);
            writer.WriteRows(filtered, plan);
            writer.Complete();
        }

        public static Task RunAsync<T>(XlsxWriter writer, IAsyncEnumerable<T> data, Action<ExportProfile<T>>? configure, CancellationToken cancellationToken)
        {
            var profile = new ExportProfile<T>();
            configure?.Invoke(profile);
            return RunAsync(writer, data, profile, cancellationToken);
        }

        public static async Task RunAsync<T>(XlsxWriter writer, IAsyncEnumerable<T> data, ExportProfile<T> profile, CancellationToken cancellationToken)
        {
            profile.Freeze();

            var plan = RowPlanBuilder.BuildTyped(profile);
            var actualSheetName = profile.SheetName ?? typeof(T).Name;
            IAsyncEnumerable<T> preparedData = profile.RowFilter is null
                ? data
                : FilterRowsAsync(data, profile.RowFilter, cancellationToken);
            bool autoSst = false;
            if (profile.AutoSst && writer.SupportsAutoSharedStrings)
            {
                var prepared = await AutoSstBufferAndDetectAsync(data, plan, cancellationToken).ConfigureAwait(false);
                preparedData = prepared.Data;
                autoSst = prepared.Enable;
            }
            try
            {
                if (autoSst) writer.EnableSharedStrings();

                var numFmts = plan.BuildNumFmts();

                if (profile.DefaultRowHeight is double dh) writer.SetDefaultRowHeight(dh);
                writer.AddSheet(actualSheetName);
                ApplySheetFeatures(writer, profile);
                writer.SetNumFmts(numFmts);
                writer.ResolveColumnStyles(plan.Columns);
                writer.WriteSheetMeta(plan.Columns, profile.FreezeHeader);
                writer.WriteHeader(plan.Columns);
                await writer.WriteRowsAsync(preparedData, plan, cancellationToken).ConfigureAwait(false);
                await writer.CompleteAsync(cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                if (preparedData is IAsyncDisposable disposable)
                    await disposable.DisposeAsync().ConfigureAwait(false);
                throw;
            }
        }

        public static Task RunAsync<T>(XlsxWriter writer, IEnumerable<T> data, Action<ExportProfile<T>>? configure, CancellationToken cancellationToken)
        {
            var profile = new ExportProfile<T>();
            configure?.Invoke(profile);
            return RunAsync(writer, data, profile, cancellationToken);
        }

        public static Task RunAsync<T>(XlsxWriter writer, IEnumerable<T> data, ExportProfile<T> profile, CancellationToken cancellationToken)
        {
            if (data is IList<T> list)
                return RunFromListAsync(writer, list, profile, cancellationToken);

            return RunAsync(writer, ToAsyncEnumerable(data), profile, cancellationToken);
        }

        private static async Task RunFromListAsync<T>(XlsxWriter writer, IList<T> data, ExportProfile<T> profile, CancellationToken cancellationToken)
        {
            profile.Freeze();

            var plan = RowPlanBuilder.BuildTyped(profile);
            var actualSheetName = profile.SheetName ?? typeof(T).Name;

            IList<T> filtered = data;
            if (profile.RowFilter is { } rowFilter)
            {
                var list = new List<T>(data.Count);
                for (int i = 0; i < data.Count; i++)
                {
                    var x = data[i];
                    if (rowFilter(x)) list.Add(x);
                }
                filtered = list;
            }

            bool autoSst = false;
            if (profile.AutoSst && writer.SupportsAutoSharedStrings)
            {
                int probeRows = Math.Min(64, filtered.Count);
                if (probeRows >= 16)
                    DetectSharedStringRatio(filtered, plan, ref autoSst, probeRows);
            }
            if (autoSst) writer.EnableSharedStrings();

            var numFmts = plan.BuildNumFmts();

            if (profile.DefaultRowHeight is double dh) writer.SetDefaultRowHeight(dh);
            writer.AddSheet(actualSheetName);
            ApplySheetFeatures(writer, profile);
            writer.SetNumFmts(numFmts);
            writer.ResolveColumnStyles(plan.Columns);
            writer.WriteSheetMeta(plan.Columns, profile.FreezeHeader);
            writer.WriteHeader(plan.Columns);
            await writer.WriteRowsFromListAsync(filtered, plan, cancellationToken).ConfigureAwait(false);
            await writer.CompleteAsync(cancellationToken).ConfigureAwait(false);
        }

        private static async IAsyncEnumerable<T> ToAsyncEnumerable<T>(IEnumerable<T> source)
        {
            foreach (var item in source) yield return item;
        }

        private static async IAsyncEnumerable<T> FilterRowsAsync<T>(
            IAsyncEnumerable<T> source,
            Func<T, bool> predicate,
            [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            var enumerator = source.GetAsyncEnumerator(cancellationToken);
            try
            {
                while (await enumerator.MoveNextAsync().ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (predicate(enumerator.Current))
                        yield return enumerator.Current;
                }
            }
            finally
            {
                await enumerator.DisposeAsync().ConfigureAwait(false);
            }
        }

        public static void BuildAndWrite(XlsxWriter writer, SheetBase sheet)
        {
            if (writer is null) throw new ArgumentNullException(nameof(writer));
            if (sheet is null) throw new ArgumentNullException(nameof(sheet));
            sheet.WriteTo(writer);
        }

        public static void WriteSheet<T>(XlsxWriter writer, string sheetName, IEnumerable<T> data, ExportProfile<T>? profile)
        {
            if (writer is null) throw new ArgumentNullException(nameof(writer));
            if (data is null) throw new ArgumentNullException(nameof(data));
            profile ??= new ExportProfile<T>().Sheet(sheetName);
            profile.Freeze();
            var plan = RowPlanBuilder.BuildTyped(profile);
            IEnumerable<T> filtered = profile.RowFilter is null ? data : FilterRows(data, profile.RowFilter);
            bool autoSst = profile.AutoSst && writer.SupportsAutoSharedStrings;
            if (autoSst)
                filtered = AutoSstBufferAndDetect(filtered, plan, out autoSst);
            writer.AddSheet(sheetName);
            if (autoSst) writer.EnableSharedStrings();
            ApplySheetFeatures(writer, profile);
            ApplyProfileToWriter(writer, profile, plan);
            writer.WriteRows(filtered, plan);
        }

        public static void WriteSheetUntyped(XlsxWriter writer, string sheetName, IEnumerable data)
        {
            if (writer is null) throw new ArgumentNullException(nameof(writer));
            if (data is null) throw new ArgumentNullException(nameof(data));
            var values = new List<object?>();
            Type? elementType = null;
            foreach (var item in data)
            {
                values.Add(item);
                if (item is null) continue;
                var itemType = item.GetType();
                if (elementType is null)
                {
                    elementType = itemType;
                    continue;
                }
                if (itemType != elementType)
                    throw new ArgumentException("All non-null items in a non-generic Sheet must have the same runtime type.", nameof(data));
            }

            if (elementType is null)
            {
                writer.AddSheet(sheetName);
                return;
            }
            if (elementType.IsValueType && values.Any(static value => value is null))
                throw new ArgumentException("Null items are not supported for value-type rows in a non-generic Sheet.", nameof(data));

            var method = typeof(XlsxWritePipeline).GetMethod(nameof(BoxMultiSheet), BindingFlags.NonPublic | BindingFlags.Static)!
                .MakeGenericMethod(elementType);
            try
            {
                method.Invoke(null, new object?[] { writer, sheetName, values });
            }
            catch (TargetInvocationException ex) when (ex.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw;
            }
        }

        private static void BoxMultiSheet<T>(XlsxWriter writer, string sheetName, IEnumerable data)
        {
            var typed = CastEnumerable<T>(data);
            var profile = new ExportProfile<T>().Sheet(sheetName);
            profile.Freeze();
            var plan = RowPlanBuilder.BuildTyped(profile);

            writer.AddSheet(sheetName);
            ApplySheetFeatures(writer, profile);
            ApplyProfileToWriter(writer, profile, plan);
            writer.WriteRows(typed, plan);
        }

        private static IEnumerable<T> CastEnumerable<T>(IEnumerable src)
        {
            foreach (var item in src) yield return (T)item!;
        }

        private static void ApplyProfileToWriter<T>(XlsxWriter writer, ExportProfile<T> profile, TypedRowPlan<T> plan)
        {
            if (profile.DefaultRowHeight is double dh) writer.SetDefaultRowHeight(dh);
            writer.SetNumFmts(plan.BuildNumFmts());
            writer.ResolveColumnStyles(plan.Columns);
            writer.WriteSheetMeta(plan.Columns, profile.FreezeHeader);
            writer.WriteHeader(plan.Columns);
        }

        private static void ApplySheetFeatures<T>(XlsxWriter writer, ExportProfile<T> profile)
        {
            if (profile.MergeCells is not null)
                foreach (var r in profile.MergeCells) writer.MergeCells(r);
            if (profile.AutoFilterRef is not null) writer.SetAutoFilter(profile.AutoFilterRef);
            foreach (var (refH, uri) in profile.Hyperlinks) writer.AddHyperlink(refH, uri);
        }

        private static IEnumerable<T> FilterRows<T>(IEnumerable<T> data, Func<T, bool> predicate)
        {
            foreach (var item in data)
                if (predicate(item)) yield return item;
        }

        private static IEnumerable<T> AutoSstBufferAndDetect<T>(IEnumerable<T> data, TypedRowPlan<T> plan, out bool enableSst)
        {
            enableSst = false;
            if (data is IList<T> list)
            {
                int probeRows = Math.Min(64, list.Count);
                if (probeRows >= 16)
                    DetectSharedStringRatio(list, plan, ref enableSst, probeRows);
                return list;
            }

            var enumerator = data.GetEnumerator();
            var probe = new List<T>(64);
            try
            {
                while (probe.Count < 64 && enumerator.MoveNext())
                    probe.Add(enumerator.Current);
            }
            catch
            {
                enumerator.Dispose();
                throw;
            }

            try
            {
                DetectSharedStringRatio(probe, plan, ref enableSst);
                return ContinueAfterProbe(probe, enumerator);
            }
            catch
            {
                enumerator.Dispose();
                throw;
            }
        }

        private static void DetectSharedStringRatio<T>(IList<T> rows, TypedRowPlan<T> plan, ref bool enableSst, int? rowCount = null)
        {
            var getters = plan.TypedGetters;
            var distinct = new HashSet<string>(StringComparer.Ordinal);
            int total = 0;
            int count = rowCount ?? rows.Count;
            for (int i = 0; i < count; i++)
            {
                var item = rows[i];
                if (item is null) continue;
                foreach (var g in getters)
                {
                    var cv = g(item);
                    if (cv.Type == CellType.String && cv.StringValue is { } s)
                    {
                        distinct.Add(s);
                        total++;
                    }
                }
            }
            if (total > 0)
                enableSst = (double)distinct.Count / total < 0.7;
        }

        private static IEnumerable<T> ContinueAfterProbe<T>(IReadOnlyList<T> probe, IEnumerator<T> enumerator)
        {
            using (enumerator)
            {
                for (int i = 0; i < probe.Count; i++)
                    yield return probe[i];
                while (enumerator.MoveNext())
                    yield return enumerator.Current;
            }
        }

        private static async Task<(IAsyncEnumerable<T> Data, bool Enable)> AutoSstBufferAndDetectAsync<T>(
            IAsyncEnumerable<T> data, TypedRowPlan<T> plan, CancellationToken cancellationToken)
        {
            var enumerator = data.GetAsyncEnumerator(cancellationToken);
            var probe = new List<T>(64);
            try
            {
                while (probe.Count < 64 && await enumerator.MoveNextAsync().ConfigureAwait(false))
                    probe.Add(enumerator.Current);
            }
            catch
            {
                await enumerator.DisposeAsync().ConfigureAwait(false);
                throw;
            }

            try
            {
                bool enable = false;
                if (probe.Count >= 16)
                    DetectSharedStringRatio(probe, plan, ref enable);
                return (ContinueAfterProbeAsync(probe, enumerator, cancellationToken), enable);
            }
            catch
            {
                await enumerator.DisposeAsync().ConfigureAwait(false);
                throw;
            }
        }

        private static async IAsyncEnumerable<T> ContinueAfterProbeAsync<T>(
            IReadOnlyList<T> probe, IAsyncEnumerator<T> enumerator,
            [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            try
            {
                for (int i = 0; i < probe.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    yield return probe[i];
                }

                while (await enumerator.MoveNextAsync().ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    yield return enumerator.Current;
                }
            }
            finally
            {
                await enumerator.DisposeAsync().ConfigureAwait(false);
            }
        }
    }
}