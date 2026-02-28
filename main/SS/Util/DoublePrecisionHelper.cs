/*
 *  ====================================================================
 *    Licensed to the collaborators of the NPOI project under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The collaborators licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
using ExtendedNumerics;
using NPOI.SS.Formula.UDF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Util
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

        public static bool IsIntegerWithDigitsDropped(this double number, int significantDigits)
        {
            return Math.Abs(GetFractionPart(DropDigitsAfterSignificantOnes(number, significantDigits))) == 0.0;
        }

        public static bool IsIntegerWithDigitsDropped(this BigDecimal number, int significantDigits)
        {
            if (number.IsZero())
                return true;

            if (number.IsNegative())
                number = -number;

            int decimalPlaces = number.DecimalPlaces;
            int realSigDigits = number.SignifigantDigits;
            int integerPlaces = realSigDigits - decimalPlaces;

            if (integerPlaces >= significantDigits)
            {
                return true;
            }

            BigDecimal fracPart = number.GetFractionalPart();

            if (fracPart.IsZero())
            {
                return true;
            }

            decimalPlaces = Math.Min(decimalPlaces, significantDigits - integerPlaces);

            BigDecimal exp = BigDecimal.Pow(10, decimalPlaces);
            fracPart *= exp;
            fracPart = BigDecimal.Round(fracPart);

            return fracPart.IsZero() || fracPart == exp;
        }
    }
}
