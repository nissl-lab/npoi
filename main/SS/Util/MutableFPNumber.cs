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

namespace NPOI.SS.Util
{
    using System;
    using NPOI.Util;
    public class MutableFPNumber
    {


        // TODO - what about values between (10<sup>14</sup>-0.5) and (10<sup>14</sup>-0.05) ?
        /**
         * The minimum value in 'Base-10 normalised form'.<br/>
         * When {@link #_binaryExponent} == 46 this is the the minimum {@link #_frac} value
         *  (10<sup>14</sup>-0.05) * 2^17
         *  <br/>
         *  Values between (10<sup>14</sup>-0.05) and 10<sup>14</sup> will be represented as '1'
         *  followed by 14 zeros.
         *  Values less than (10<sup>14</sup>-0.05) will get Shifted by one more power of 10
         *
         *  This frac value rounds to '1' followed by fourteen zeros with an incremented decimal exponent
         */
        //private static BigInteger BI_MIN_BASE = new BigInteger("0B5E620F47FFFE666", 16);
        private static readonly BigInteger BI_MIN_BASE = new BigInteger(new int[] { -1243209484, 2147477094 }, 1);
        /**
         * For 'Base-10 normalised form'<br/>
         * The maximum {@link #_frac} value when {@link #_binaryExponent} == 49
         * (10^15-0.5) * 2^14
         */
        //private static BigInteger BI_MAX_BASE = new BigInteger("0E35FA9319FFFE000", 16);
        private static readonly BigInteger BI_MAX_BASE = new BigInteger(new int[] { -480270031, -1610620928 }, 1);

        /**
         * Width of a long
         */
        private const int C_64 = 64;

        /**
         * Minimum precision after discarding whole 32-bit words from the significand
         */
        private const int MIN_PRECISION = 72;
        private BigInteger _significand;
        private int _binaryExponent;
        public MutableFPNumber(BigInteger frac, int binaryExponent)
        {
            _significand = frac;
            _binaryExponent = binaryExponent;
        }


        public MutableFPNumber Copy()
        {
            return new MutableFPNumber(_significand, _binaryExponent);
        }
        public void Normalise64bit()
        {
            int oldBitLen = _significand.BitLength();
            int sc = oldBitLen - C_64;
            if (sc == 0)
            {
                return;
            }
            if (sc < 0)
            {
                throw new InvalidOperationException("Not enough precision");
            }
            _binaryExponent += sc;
            if (sc > 32)
            {
                int highShift = (sc - 1) & 0xFFFFE0;
                _significand = _significand>>(highShift);
                sc -= highShift;
                oldBitLen -= highShift;
            }
            if (sc < 1)
            {
                throw new InvalidOperationException();
            }
            _significand = Rounder.Round(_significand, sc);
            if (_significand.BitLength() > oldBitLen)
            {
                sc++;
                _binaryExponent++;
            }
            _significand = _significand>>(sc);
        }
        public int Get64BitNormalisedExponent()
        {
            //return _binaryExponent + _significand.BitCount() - C_64;
            return _binaryExponent + _significand.BitLength() - C_64;
        }

        public bool IsBelowMaxRep()
        {
            int sc = _significand.BitLength() - C_64;
            //return _significand<(BI_MAX_BASE<<(sc));
            return _significand.CompareTo(BI_MAX_BASE.ShiftLeft(sc)) < 0;
        }
        public bool IsAboveMinRep()
        {
            int sc = _significand.BitLength() - C_64;
            return _significand.CompareTo(BI_MIN_BASE.ShiftLeft(sc)) > 0;
            //return _significand>(BI_MIN_BASE<<(sc));
        }
        public NormalisedDecimal CreateNormalisedDecimal(int pow10)
        {
            // missingUnderBits is (0..3)
            int missingUnderBits = _binaryExponent - 39;
            int fracPart = (_significand.IntValue() << missingUnderBits) & 0xFFFF80;
            long wholePart = (_significand>>(C_64 - _binaryExponent - 1)).LongValue();
            return new NormalisedDecimal(wholePart, fracPart, pow10);
        }
        public void multiplyByPowerOfTen(int pow10)
        {
            TenPower tp = TenPower.GetInstance(Math.Abs(pow10));
            if (pow10 < 0)
            {
                mulShift(tp._divisor, tp._divisorShift);
            }
            else
            {
                mulShift(tp._multiplicand, tp._multiplierShift);
            }
        }
        private void mulShift(BigInteger multiplicand, int multiplierShift)
        {
            _significand = _significand*multiplicand;
            _binaryExponent += multiplierShift;
            // check for too much precision
            int sc = (_significand.BitLength() - MIN_PRECISION) & unchecked((int)0xFFFFFFE0);
            // mask Makes multiples of 32 which optimises BigInt32.ShiftRight
            if (sc > 0)
            {
                // no need to round because we have at least 8 bits of extra precision
                _significand = _significand>>(sc);
                _binaryExponent += sc;
            }
        }

        private class Rounder
        {
            private static BigInteger[] HALF_BITS;

            static Rounder()
            {
                BigInteger[] bis = new BigInteger[33];
                long acc = 1;
                for (int i = 1; i < bis.Length; i++)
                {
                    bis[i] = new BigInteger(acc);
                    acc <<= 1;
                }
                HALF_BITS = bis;
            }
            /**
             * @param nBits number of bits to shift right
             */
            public static BigInteger Round(BigInteger bi, int nBits)
            {
                if (nBits < 1)
                {
                    return bi;
                }
                return bi+(HALF_BITS[nBits]);
            }
        }

        /**
         * Holds values for quick multiplication and division by 10
         */
        private class TenPower
        {
            private static readonly BigInteger FIVE = new BigInteger(5L);// new BigInteger("5",10);
            private static TenPower[] _cache = new TenPower[350];

            public BigInteger _multiplicand;
            public BigInteger _divisor;
            public int _divisorShift;
            public int _multiplierShift;

            private TenPower(int index)
            {
                //BigInteger fivePowIndex = FIVE.ModPow(new BigInteger(index),FIVE);
                BigInteger fivePowIndex = FIVE.Pow(index);
                int bitsDueToFiveFactors = fivePowIndex.BitLength();
                int px = 80 + bitsDueToFiveFactors;
                BigInteger fx = (BigInteger.One << px) / (fivePowIndex);
                int adj = fx.BitLength() - 80;
                _divisor = fx>>(adj);
                bitsDueToFiveFactors -= adj;

                _divisorShift = -(bitsDueToFiveFactors + index + 80);
                int sc = fivePowIndex.BitLength() - 68;
                if (sc > 0)
                {
                    _multiplierShift = index + sc;
                    _multiplicand = fivePowIndex>>(sc);
                }
                else
                {
                    _multiplierShift = index;
                    _multiplicand = fivePowIndex;
                }
            }

            public static TenPower GetInstance(int index)
            {
                TenPower result = _cache[index];
                if (result == null)
                {
                    result = new TenPower(index);
                    _cache[index] = result;
                }
                return result;
            }
        }

        public ExpandedDouble CreateExpandedDouble()
        {
            return new ExpandedDouble(_significand, _binaryExponent);
        }
    }
}

