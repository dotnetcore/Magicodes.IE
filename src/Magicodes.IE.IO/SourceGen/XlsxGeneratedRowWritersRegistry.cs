
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// Source-generated registry mapping types to their strongly typed row writers.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class XlsxGeneratedRowWritersRegistry
    {
        private static readonly ConcurrentDictionary<Type, object> _writers = new();

        /// <summary>
        /// Registers a row-writer factory for the specified type.
        /// </summary>
        public static void Register<T>(Func<IReadOnlyDictionary<string, Action<XlsxWriter.XlsxRowWriter, T, int>>> factory)
        {
            if (factory is null) return;
            var dict = factory();
            if (dict is null || dict.Count == 0) return;
            _writers[typeof(T)] = dict;
        }

        /// <summary>
        /// Retrieves the registered row writer for the given type, or <see langword="null"/> when no generated code exists.
        /// </summary>
        public static IReadOnlyDictionary<string, Action<XlsxWriter.XlsxRowWriter, T, int>>? TryGet<T>()
        {
            if (_writers.TryGetValue(typeof(T), out var o) && o is IReadOnlyDictionary<string, Action<XlsxWriter.XlsxRowWriter, T, int>> d)
                return d;
            return null;
        }
    }
}
