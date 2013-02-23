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
    using System.Text;
    using NPOI.Util;
    using System.Globalization;

    /*
     * Represents a transformation of a 64 bit IEEE double quantity having a decimal exponent and a
     * fixed point (15 decimal digit) significand.  Some quirks of Excel's calculation behaviour are
     * simpler to reproduce with numeric quantities in this format.  This class is currently used to
     * help:
     * <ol>
     * <li>Comparison operations</li>
     * <li>Conversions to text</li>
     * </ol>
     *
     * <p/>
     * This class does not handle negative numbers or zero.
     * <p/>
     * The value of a {@link NormalisedDecimal} is given by<br/>
     * <c> significand &times; 10<sup>decimalExponent</sup></c>
     * <br/>
     * where:<br/>
     *
     * <c>significand</c> = wholePart + fractionalPart / 2<sup>24</sup><br/>
     *
     * @author Josh Micich
     */
    public class NormalisedDecimal
    {
        /**
         * Number of powers of ten Contained in the significand
         */
        private const int EXPONENT_OFFSET = 14;

        private static readonly decimal BD_2_POW_24 = new decimal((BigInteger.One << 24).LongValue());

        /*
         * log<sub>10</sub>(2)&times;2<sup>20</sup>
         */
        private const int LOG_BASE_10_OF_2_TIMES_2_POW_20 = 315653; // 315652.8287

        /**
         * 2<sup>19</sup>
         */
        private const int C_2_POW_19 = 1 << 19;


        /**
         * the value of {@link #_fractionalPart} that represents 0.5
         */
        private const int FRAC_HALF = 0x800000;

        /**
         * 10<sup>15</sup>
         */
        private const long MAX_REP_WHOLE_PART = 0x38D7EA4C68000L;



        public static NormalisedDecimal Create(BigInteger frac, int binaryExponent)
        {
            // estimate pow2&pow10 first, perform optional mulShift, then normalize
            int pow10;
            if (binaryExponent > 49 || binaryExponent < 46)
            {

                // working with ints (left Shifted 20) instead of doubles
                // x = 14.5 - binaryExponent * log10(2);
                int x = (29 << 19) - binaryExponent * LOG_BASE_10_OF_2_TIMES_2_POW_20;
                x += C_2_POW_19; // round
                pow10 = -(x >> 20);
            }
            else
            {
                pow10 = 0;
            }
            MutableFPNumber cc = new MutableFPNumber(frac, binaryExponent);
            if (pow10 != 0)
            {
                cc.multiplyByPowerOfTen(-pow10);
            }

            switch (cc.Get64BitNormalisedExponent())
            {
                case 46:
                    if (cc.IsAboveMinRep())
                    {
                        break;
                    }
                    goto case 44;
                case 44:
                case 45:
                    cc.multiplyByPowerOfTen(1);
                    pow10--;
                    break;
                case 47:
                case 48:
                    break;
                case 49:
                    if (cc.IsBelowMaxRep())
                    {
                        break;
                    }
                    goto case 50;
                case 50:
                    cc.multiplyByPowerOfTen(-1);
                    pow10++;
                    break;

                default:
                    throw new InvalidOperationException("Bad binary exp " + cc.Get64BitNormalisedExponent() + ".");
            }
            cc.Normalise64bit();

            return cc.CreateNormalisedDecimal(pow10);
        }

        /**
         * Rounds at the digit with value 10<sup>decimalExponent</sup>
         */
        public NormalisedDecimal RoundUnits()
        {
            long wholePart = _wholePart;
            if (_fractionalPart >= FRAC_HALF)
            {
                wholePart++;
            }

            int de = _relativeDecimalExponent;

            if (wholePart < MAX_REP_WHOLE_PART)
            {
                return new NormalisedDecimal(wholePart, 0, de);
            }
            return new NormalisedDecimal(wholePart / 10, 0, de + 1);
        }

        /**
         * The decimal exponent increased by one less than the digit count of {@link #_wholePart}
         */
        private int _relativeDecimalExponent;
        /**
         * The whole part of the significand (typically 15 digits).
         *
         * 47-50 bits long (MSB may be anywhere from bit 46 to 49)
         * LSB is units bit.
         */
        private long _wholePart;
        /**
         * The fractional part of the significand.
         * 24 bits (only top 14-17 bits significant): a value between 0x000000 and 0xFFFF80
         */
        private int _fractionalPart;


        public NormalisedDecimal(long wholePart, int fracPart, int decimalExponent)
        {
            _wholePart = wholePart;
            _fractionalPart = fracPart;
            _relativeDecimalExponent = decimalExponent;
        }


        /**
         * Convert to an equivalent {@link ExpandedDouble} representation (binary frac and exponent).
         * The resulting transformed object is easily Converted to a 64 bit IEEE double:
         * <ul>
         * <li>bits 2-53 of the {@link #GetSignificand()} become the 52 bit 'fraction'.</li>
         * <li>{@link #GetBinaryExponent()} is biased by 1023 to give the 'exponent'.</li>
         * </ul>
         * The sign bit must be obtained from somewhere else.
         * @return a new {@link NormalisedDecimal} normalised to base 2 representation.
         */
        public ExpandedDouble NormaliseBaseTwo()
        {
            MutableFPNumber cc = new MutableFPNumber(ComposeFrac(), 39);
            cc.multiplyByPowerOfTen(_relativeDecimalExponent);
            cc.Normalise64bit();
            return cc.CreateExpandedDouble();
        }

        /**
         * @return the significand as a fixed point number (with 24 fraction bits and 47-50 whole bits)
         */
        public BigInteger ComposeFrac()
        {
            long wp = _wholePart;
            int fp = _fractionalPart;
            return new BigInteger(new byte[] {
				(byte) (wp >> 56), // N.B. assuming sign bit is zero
				(byte) (wp >> 48),
				(byte) (wp >> 40),
				(byte) (wp >> 32),
				(byte) (wp >> 24),
				(byte) (wp >> 16),
				(byte) (wp >>  8),
				(byte) (wp >>  0),
				(byte) (fp >> 16),
				(byte) (fp >> 8),
				(byte) (fp >> 0),
		});
        }

        public String GetSignificantDecimalDigits()
        {
            return _wholePart.ToString(CultureInfo.InvariantCulture);
        }
        /**
         * Rounds the first whole digit position (considers only units digit, not frational part).
         * Caller should check total digit count of result to see whether the rounding operation caused
         * a carry out of the most significant digit
         */
        public String GetSignificantDecimalDigitsLastDigitRounded()
        {
            long wp = _wholePart + 5; // rounds last digit
            StringBuilder sb = new StringBuilder(24);
            sb.Append(wp);
            sb[sb.Length - 1]= '0';
            return sb.ToString();
        }

        /**
         * @return the number of powers of 10 which have been extracted from the significand and binary exponent.
         */
        public int GetDecimalExponent()
        {
            return _relativeDecimalExponent + EXPONENT_OFFSET;
        }

        /**
         * assumes both this and other are normalised
         */
        public int CompareNormalised(NormalisedDecimal other)
        {
            int cmp = _relativeDecimalExponent - other._relativeDecimalExponent;
            if (cmp != 0)
            {
                return cmp;
            }
            if (_wholePart > other._wholePart)
            {
                return 1;
            }
            if (_wholePart < other._wholePart)
            {
                return -1;
            }
            return _fractionalPart - other._fractionalPart;
        }
        public decimal GetFractionalPart()
        {
            return new decimal(_fractionalPart)/(BD_2_POW_24);
        }

        private String GetFractionalDigits()
        {
            if (_fractionalPart == 0)
            {
                return "0";
            }
            return GetFractionalPart().ToString(CultureInfo.InvariantCulture).Substring(2);
        }


        public override String ToString()
        {

            StringBuilder sb = new StringBuilder();
            sb.Append(this.GetType().Name);
            sb.Append(" [");
            String ws = _wholePart.ToString(CultureInfo.InvariantCulture);
            sb.Append(ws[0]);
            sb.Append('.');
            sb.Append(ws.Substring(1));
            sb.Append(' ');
            sb.Append(GetFractionalDigits());
            sb.Append("E");
            sb.Append(GetDecimalExponent());
            sb.Append("]");
            return sb.ToString();
        }
    }


}
