using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Utils
{
    public static class UtilExtentsions
    {
        //public static bool ChangeKey<TKey, TValue>(this IDictionary<TKey, TValue> dict,
        //								   TKey oldKey, TKey newKey)
        //{
        //	TValue value;
        //	if (!dict.Remove(oldKey, out value))
        //		return false;

        //	dict[newKey] = value;  // or dict.Add(newKey, value) depending on ur comfort
        //	return true;
        //}

        /// <summary>
        /// Attempts to change the key of a value in the dictionary.
        /// </summary>
        public static bool TryChangeKey<TKey, TValue>(this Dictionary<TKey, TValue> source,
                                                TKey oldKey, TKey newKey, out TValue value)
        {
            if (source.ContainsKey(newKey))
            {
                value = default;
                return false;
            }
            if (!source.TryGetValue(oldKey, out value))
                return false;
            if (!source.Remove(oldKey))
                return false;
            source.Add(newKey, value);
            return true;
        }
    }
}