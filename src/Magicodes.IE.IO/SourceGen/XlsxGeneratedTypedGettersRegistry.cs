
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// Source-generated registry mapping types to their strongly typed getter delegates.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class XlsxGeneratedTypedGettersRegistry
    {
        private static readonly ConcurrentDictionary<Type, IReadOnlyDictionary<string, Delegate>> _getters = new();

        /// <summary>
        /// Registers a strongly typed getter factory for the specified type.
        /// </summary>
        public static void Register<T>(Func<IReadOnlyDictionary<string, Func<T, CellValue>>> factory)
        {
            if (factory is null) return;
            var dict = factory();
            if (dict is null || dict.Count == 0) return;
            var d = new System.Collections.Generic.Dictionary<string, Delegate>(dict.Count);
            foreach (var kv in dict) d[kv.Key] = kv.Value;
            _getters[typeof(T)] = d;
        }

        /// <summary>
        /// Retrieves the registered strongly typed getters for the given type, or <see langword="null"/> when no generated code exists.
        /// </summary>
        public static IReadOnlyDictionary<string, Delegate>? TryGet(Type t)
        {
            return _getters.TryGetValue(t, out var g) ? g : null;
        }
    }
}
