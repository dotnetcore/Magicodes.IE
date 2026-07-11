
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// Source-generated registry mapping types to their object-based getter delegates.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class XlsxGeneratedGettersRegistry
    {
        private static readonly ConcurrentDictionary<Type, IReadOnlyDictionary<string, Func<object?, CellValue>>> _getters = new();

        /// <summary>
        /// Registers a getter factory for the specified type.
        /// </summary>
        public static void Register<T>(Func<IReadOnlyDictionary<string, Func<object?, CellValue>>> factory)
        {
            if (factory is null) return;
            var dict = factory();
            if (dict is null || dict.Count == 0) return;
            _getters[typeof(T)] = dict;
        }

        /// <summary>
        /// Retrieves the registered getters for the given type, or <see langword="null"/> when no generated code exists.
        /// </summary>
        public static IReadOnlyDictionary<string, Func<object?, CellValue>>? TryGet(Type t)
        {
            return _getters.TryGetValue(t, out var g) ? g : null;
        }
    }
}
