using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Util
{
    public class IEEEDouble
    {
        private static long EXPONENT_MASK = 0x7FF0000000000000L;
        private static int EXPONENT_SHIFT = 52;
        public static long FRAC_MASK = 0x000FFFFFFFFFFFFFL;
        public static int EXPONENT_BIAS = 1023;
        public static long FRAC_ASSUMED_HIGH_BIT = (1L << EXPONENT_SHIFT);
        /**
         * The value the exponent field Gets for all <i>NaN</i> and <i>InfInity</i> values
         */
        public static int BIASED_EXPONENT_SPECIAL_VALUE = 0x07FF;

        /**
         * @param rawBits the 64 bit binary representation of the double value
         * @return the top 12 bits (sign and biased exponent value)
         */
        public static int GetBiasedExponent(long rawBits)
        {
            return (int)((rawBits & EXPONENT_MASK) >> EXPONENT_SHIFT);
        }
    }
}
