using System;
using System.Collections.Generic;

namespace NPOI
{
    internal static class Polyfills
    {
#if NETFRAMEWORK || NETSTANDARD2_0
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    internal static bool Contains(this string source, char c) => source.IndexOf(c) != -1;

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    internal static bool StartsWith(this string source, char c) => source.Length > 0 && source[0] == c;

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    internal static bool EndsWith(this string source, char c) => source.Length > 0 && source[source.Length - 1] == c;

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    internal static string[] Split(this string source, char c, StringSplitOptions options) => source.Split([c], options);

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    internal static bool TryAdd<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue value)
    {
        if(dictionary.ContainsKey(key))
        {
            return false;
        }

        dictionary[key] = value;
        return true;
    }
#endif

#if NETFRAMEWORK
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    internal static bool Contains(this ReadOnlySpan<string> source, string c) => source.IndexOf(c) != -1;
#endif
        internal static string ReplaceFirst(this string text, string search, string replace)
        {
            if(string.IsNullOrEmpty(text) || string.IsNullOrEmpty(search))
                return text;

            int pos = text.IndexOf(search);
            if(pos < 0)
                return text;

            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }
    }
}
