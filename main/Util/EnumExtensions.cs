using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Util
{
    /// <summary>
    /// This class provides helper methods that reduce dynamic code / reflection use on Enums
    /// </summary>
    public static class EnumExtensions
    {
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

        public static T Parse<T>(string name, bool ignoreCase = false) where T : struct, Enum
        {
            return (T)Enum.Parse(typeof(T), name, ignoreCase);
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

        public static T Parse<T>(string name, bool ignoreCase = false) where T : struct, Enum
        {
            return Enum.Parse<T>(name, ignoreCase);
        }
#endif
    }
}
