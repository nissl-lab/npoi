using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.SS.UserModel
{
    public static class NCellRangeExtensions
    {
        public static double Sum(this NCellRange range, Func<ICell, double> selector)
        {
            return range.Cells.Sum(selector);
        }

        public static double Min(this NCellRange range, Func<ICell, double> selector)
        {
            return range.Cells.Min(selector);
        }

        public static double Max(this NCellRange range, Func<ICell, double> selector)
        {
            return range.Cells.Max(selector);
        }

        public static double Average(this NCellRange range, Func<ICell, double> selector)
        {
            return range.Cells.Average(selector);
        }

        public static T Max<T>(this NCellRange range, Func<ICell, T> selector) where T : IComparable<T>
        {
            return range.Cells.Max(selector);
        }

        public static T Min<T>(this NCellRange range, Func<ICell, T> selector) where T : IComparable<T>
        {
            return range.Cells.Min(selector);
        }

        public static IEnumerable<ICell> Where(this NCellRange range, Func<ICell, bool> predicate)
        {
            return range.Cells.Where(predicate);
        }

        public static IEnumerable<TResult> Select<TResult>(this NCellRange range, Func<ICell, TResult> selector)
        {
            return range.Cells.Select(selector);
        }

        public static int Count(this NCellRange range)
        {
            return range.Cells.Count;
        }
    }
}
