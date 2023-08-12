using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.OOXML.Util
{
    /// <summary>
    /// This class provides helper methods that reduce dynamic code / reflection use, for better AOT performance.
    /// </summary>
    internal static class CollectionExtensions
    {
        public static T[] ToArray<T>(this ArrayList arrayList)
        {
            if (arrayList.Count == 0)
                return Array.Empty<T>();

            var array = new T[arrayList.Count];
            arrayList.CopyTo(array);

            return array;
        }
    }
}
