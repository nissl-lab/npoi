/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System;
namespace NPOI.SS.Util
{


    /**
     * Represents a 64 bit IEEE double quantity expressed with both decimal and binary exponents
     * Does not handle negative numbers or zero
     * <p/>
     * The value of a {@link ExpandedDouble} is given by<br/>
     * <tt> a &times; 2<sup>b</sup></tt>
     * <br/>
     * where:<br/>
     *
     * <tt>a</tt> = <i>significand</i><br/>
     * <tt>b</tt> = <i>binaryExponent</i> - bitLength(significand) + 1<br/>
     *
     * @author Josh Micich
     */
    class ExpandedDouble
    {
        private static BigInteger BI_FRAC_MASK = BigInt32.ValueOf(FRAC_MASK);
        private static BigInteger BI_IMPLIED_FRAC_MSB = BigInt32.ValueOf(FRAC_ASSUMED_HIGH_BIT);

        private static BigInteger GetFrac(long rawBits)
        {
            return BigInt32.ValueOf(rawBits).and(BI_FRAC_MASK).or(BI_IMPLIED_FRAC_MSB).ShiftLeft(11);
        }


        public static ExpandedDouble fromRawBitsAndExponent(long rawBits, int exp)
        {
            return new ExpandedDouble(GetFrac(rawBits), exp);
        }

        /**
         * Always 64 bits long (MSB, bit-63 is '1')
         */
        private BigInteger _significand;
        private int _binaryExponent;

        public ExpandedDouble(long rawBits)
        {
            int biasedExp = (int)(rawBits >> 52);
            if (biasedExp == 0)
            {
                // sub-normal numbers
                BigInteger frac = BigInt32.ValueOf(rawBits).and(BI_FRAC_MASK);
                int expAdj = 64 - frac.bitLength();
                _significand = frac.ShiftLeft(expAdj);
                _binaryExponent = (biasedExp & 0x07FF) - 1023 - expAdj;
            }
            else
            {
                BigInteger frac = GetFrac(rawBits);
                _significand = frac;
                _binaryExponent = (biasedExp & 0x07FF) - 1023;
            }
        }

        ExpandedDouble(BigInteger frac, int binaryExp)
        {
            if (frac.bitLength() != 64)
            {
                throw new ArgumentException("bad bit length");
            }
            _significand = frac;
            _binaryExponent = binaryExp;
        }


        /**
         * Convert to an equivalent {@link NormalisedDecimal} representation having 15 decimal digits of precision in the
         * non-fractional bits of the significand.
         */
        public NormalisedDecimal NormaliseBaseTen()
        {
            return NormalisedDecimal.Create(_significand, _binaryExponent);
        }

        /**
         * @return the number of non-fractional bits after the MSB of the significand
         */
        public int GetBinaryExponent()
        {
            return _binaryExponent;
        }

        public BigInteger GetSignificand()
        {
            return _significand;
        }
    }
}

