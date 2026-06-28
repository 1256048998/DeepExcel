using System.Collections.Generic;

namespace DeepExcel.AddIn
{
    internal static class DictionaryExtensions
    {
        public static TValue GetValueOrDefault<TKey, TValue>(
            this Dictionary<TKey, TValue> dict, TKey key, TValue defaultValue = default(TValue))
        {
            if (dict == null) return defaultValue;
            TValue value;
            if (dict.TryGetValue(key, out value))
                return value;
            return defaultValue;
        }
    }
}
