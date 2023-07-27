using NPOI.SS.Formula.UDF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    internal static class DoublePrecisionHelper
    {
        public static double GetFractionPart(double number)
        {
            return Math.Abs(number - Math.Truncate(number));
        }
        public static double DropDigitsAfterSignificantOnes(double number, int digits)
        {
            if (number == 0.0) return 0.0;

            var isNegative = number < 0;
            var positiveNumber = isNegative ? -number : number;

            var mostSignificantDigit = Math.Floor(Math.Log10(positiveNumber));
            var multiplier = Math.Pow(10, digits - mostSignificantDigit - 1);

            var newNumber = positiveNumber * multiplier;
            newNumber = GetFractionPart(newNumber) >= 0.5 ? Math.Truncate(newNumber) + 1 : Math.Truncate(newNumber);

            newNumber /= multiplier;
            return isNegative ? -newNumber : newNumber;
        }

        public static bool IsIntegerWithDigitsDropped(double number, int significantDigits)
        {
            return Math.Abs(GetFractionPart(DropDigitsAfterSignificantOnes(number, significantDigits))) == 0.0;
        }
    }
}
