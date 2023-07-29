using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Util
{
    /// <summary>
    /// This class provides helper methods that reduce dynamic code / reflection use, for better AOT performance.
    /// </summary>
    public static class AotExtensions
    {
        public static T[] ToArray<T>(this ArrayList arrayList)
        {
            if (arrayList.Count == 0)
                return Array.Empty<T>();

            var array = new T[arrayList.Count];
            arrayList.CopyTo(array);

            return array;
        }

#if !NET6_0_OR_GREATER
        public static T[] GetEnumValues<T>() where T : struct, Enum
        {
            return (T[])Enum.GetValues(typeof(T));
        }

        public static string GetEnumName<T>(T val) where T : struct, Enum
        {
            return Enum.GetName(typeof(T), val);
        }

        public static string[] GetEnumNames<T>() where T : struct, Enum
        {
            return Enum.GetNames(typeof(T));
        }
#else
        // AOT-friendly
        public static T[] GetEnumValues<T>() where T : struct, Enum
        {
            return Enum.GetValues<T>();
        }

        public static string GetEnumName<T>(T val) where T : struct, Enum
        {
            return Enum.GetName<T>(val);
        }

        public static string[] GetEnumNames<T>() where T : struct, Enum
        {
            return Enum.GetNames<T>();
        }
#endif
    }
}
