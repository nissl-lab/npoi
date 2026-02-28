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
using System.Text;
namespace NPOI.SS.Util
{


    /*
     * Excel Converts numbers to text with different rules to those of java, so
     *  <c>Double.ToString(value)</c> won't do.
     * <ul>
     * <li>No more than 15 significant figures are output (java does 18).</li>
     * <li>The sign char for the exponent is included even if positive</li>
     * <li>Special values (<c>NaN</c> and <c>InfInity</c>) Get rendered like the ordinary
     * number that the bit pattern represents.</li>
     * <li>Denormalised values (between &plusmn;2<sup>-1074</sup> and &plusmn;2<sup>-1022</sup>
     *  are displayed as "0"</sup>
     * </ul>
     * IEEE 64-bit Double Rendering Comparison
     *
     * <table border="1" cellpAdding="2" cellspacing="0" summary="IEEE 64-bit Double Rendering Comparison">
     * <tr><th>Raw bits</th><th>Java</th><th>Excel</th></tr>
     *
     * <tr><td>0x0000000000000000L</td><td>0.0</td><td>0</td></tr>
     * <tr><td>0x3FF0000000000000L</td><td>1.0</td><td>1</td></tr>
     * <tr><td>0x3FF00068DB8BAC71L</td><td>1.0001</td><td>1.0001</td></tr>
     * <tr><td>0x4087A00000000000L</td><td>756.0</td><td>756</td></tr>
     * <tr><td>0x401E3D70A3D70A3DL</td><td>7.56</td><td>7.56</td></tr>
     * <tr><td>0x405EDD3C07FB4C99L</td><td>123.45678901234568</td><td>123.456789012346</td></tr>
     * <tr><td>0x4132D687E3DF2180L</td><td>1234567.8901234567</td><td>1234567.89012346</td></tr>
     * <tr><td>0x3EE9E409302678BAL</td><td>1.2345678901234568E-5</td><td>1.23456789012346E-05</td></tr>
     * <tr><td>0x3F202E85BE180B74L</td><td>1.2345678901234567E-4</td><td>0.000123456789012346</td></tr>
     * <tr><td>0x3F543A272D9E0E51L</td><td>0.0012345678901234567</td><td>0.00123456789012346</td></tr>
     * <tr><td>0x3F8948B0F90591E6L</td><td>0.012345678901234568</td><td>0.0123456789012346</td></tr>
     * <tr><td>0x3EE9E409301B5A02L</td><td>1.23456789E-5</td><td>0.0000123456789</td></tr>
     * <tr><td>0x3E6E7D05BDABDE50L</td><td>5.6789012345E-8</td><td>0.000000056789012345</td></tr>
     * <tr><td>0x3E6E7D05BDAD407EL</td><td>5.67890123456E-8</td><td>5.67890123456E-08</td></tr>
     * <tr><td>0x3E6E7D06029F18BEL</td><td>5.678902E-8</td><td>0.00000005678902</td></tr>
     * <tr><td>0x2BCB5733CB32AE6EL</td><td>9.999999999999123E-98</td><td>9.99999999999912E-98</td></tr>
     * <tr><td>0x2B617F7D4ED8C59EL</td><td>1.0000000000001235E-99</td><td>1.0000000000001E-99</td></tr>
     * <tr><td>0x0036319916D67853L</td><td>1.2345678901234578E-307</td><td>1.2345678901235E-307</td></tr>
     * <tr><td>0x359DEE7A4AD4B81FL</td><td>2.0E-50</td><td>2E-50</td></tr>
     * <tr><td>0x41678C29DCD6E9E0L</td><td>1.2345678901234567E7</td><td>12345678.9012346</td></tr>
     * <tr><td>0x42A674E79C5FE523L</td><td>1.2345678901234568E13</td><td>12345678901234.6</td></tr>
     * <tr><td>0x42DC12218377DE6BL</td><td>1.2345678901234567E14</td><td>123456789012346</td></tr>
     * <tr><td>0x43118B54F22AEB03L</td><td>1.2345678901234568E15</td><td>1234567890123460</td></tr>
     * <tr><td>0x43E56A95319D63E1L</td><td>1.2345678901234567E19</td><td>12345678901234600000</td></tr>
     * <tr><td>0x441AC53A7E04BCDAL</td><td>1.2345678901234568E20</td><td>1.23456789012346E+20</td></tr>
     * <tr><td>0xC3E56A95319D63E1L</td><td>-1.2345678901234567E19</td><td>-12345678901234600000</td></tr>
     * <tr><td>0xC41AC53A7E04BCDAL</td><td>-1.2345678901234568E20</td><td>-1.23456789012346E+20</td></tr>
     * <tr><td>0x54820FE0BA17F46DL</td><td>1.2345678901234577E99</td><td>1.2345678901235E+99</td></tr>
     * <tr><td>0x54B693D8E89DF188L</td><td>1.2345678901234576E100</td><td>1.2345678901235E+100</td></tr>
     * <tr><td>0x4A611B0EC57E649AL</td><td>2.0E50</td><td>2E+50</td></tr>
     * <tr><td>0x7FEFFFFFFFFFFFFFL</td><td>1.7976931348623157E308</td><td>1.7976931348623E+308</td></tr>
     * <tr><td>0x0010000000000000L</td><td>2.2250738585072014E-308</td><td>2.2250738585072E-308</td></tr>
     * <tr><td>0x000FFFFFFFFFFFFFL</td><td>2.225073858507201E-308</td><td>0</td></tr>
     * <tr><td>0x0000000000000001L</td><td>4.9E-324</td><td>0</td></tr>
     * <tr><td>0x7FF0000000000000L</td><td>InfInity</td><td>1.7976931348623E+308</td></tr>
     * <tr><td>0xFFF0000000000000L</td><td>-InfInity</td><td>1.7976931348623E+308</td></tr>
     * <tr><td>0x441AC7A08EAD02F2L</td><td>1.234999999999999E20</td><td>1.235E+20</td></tr>
     * <tr><td>0x40FE26BFFFFFFFF9L</td><td>123499.9999999999</td><td>123500</td></tr>
     * <tr><td>0x3E4A857BFB2F2809L</td><td>1.234999999999999E-8</td><td>0.00000001235</td></tr>
     * <tr><td>0x3BCD291DEF868C89L</td><td>1.234999999999999E-20</td><td>1.235E-20</td></tr>
     * <tr><td>0x444B1AE4D6E2EF4FL</td><td>9.999999999999999E20</td><td>1E+21</td></tr>
     * <tr><td>0x412E847FFFFFFFFFL</td><td>999999.9999999999</td><td>1000000</td></tr>
     * <tr><td>0x3E45798EE2308C39L</td><td>9.999999999999999E-9</td><td>0.00000001</td></tr>
     * <tr><td>0x3C32725DD1D243ABL</td><td>9.999999999999999E-19</td><td>0.000000000000000001</td></tr>
     * <tr><td>0x3BFD83C94FB6D2ABL</td><td>9.999999999999999E-20</td><td>1E-19</td></tr>
     * <tr><td>0xC44B1AE4D6E2EF4FL</td><td>-9.999999999999999E20</td><td>-1E+21</td></tr>
     * <tr><td>0xC12E847FFFFFFFFFL</td><td>-999999.9999999999</td><td>-1000000</td></tr>
     * <tr><td>0xBE45798EE2308C39L</td><td>-9.999999999999999E-9</td><td>-0.00000001</td></tr>
     * <tr><td>0xBC32725DD1D243ABL</td><td>-9.999999999999999E-19</td><td>-0.000000000000000001</td></tr>
     * <tr><td>0xBBFD83C94FB6D2ABL</td><td>-9.999999999999999E-20</td><td>-1E-19</td></tr>
     * <tr><td>0xFFFF0420003C0000L</td><td>NaN</td><td>3.484840871308E+308</td></tr>
     * <tr><td>0x7FF8000000000000L</td><td>NaN</td><td>2.6965397022935E+308</td></tr>
     * <tr><td>0x7FFF0420003C0000L</td><td>NaN</td><td>3.484840871308E+308</td></tr>
     * <tr><td>0xFFF8000000000000L</td><td>NaN</td><td>2.6965397022935E+308</td></tr>
     * <tr><td>0xFFFF0AAAAAAAAAAAL</td><td>NaN</td><td>3.4877119413344E+308</td></tr>
     * <tr><td>0x7FF80AAAAAAAAAAAL</td><td>NaN</td><td>2.7012211948322E+308</td></tr>
     * <tr><td>0xFFFFFFFFFFFFFFFFL</td><td>NaN</td><td>3.5953862697246E+308</td></tr>
     * <tr><td>0x7FFFFFFFFFFFFFFFL</td><td>NaN</td><td>3.5953862697246E+308</td></tr>
     * <tr><td>0xFFF7FFFFFFFFFFFFL</td><td>NaN</td><td>2.6965397022935E+308</td></tr>
     * </table>
     *
     * <b>Note</b>:
     * Excel has inconsistent rules for the following numeric operations:
     * <ul>
     * <li>Conversion to string (as handled here)</li>
     * <li>Rendering numerical quantities in the cell grid.</li>
     * <li>Conversion from text</li>
     * <li>General arithmetic</li>
     * </ul>
     * Excel's text to number conversion is not a true <i>inverse</i> of this operation.  The
     * allowable ranges are different.  Some numbers that don't correctly convert to text actually
     * <b>do</b> Get handled properly when used in arithmetic Evaluations.
     *
     * @author Josh Micich
     */
    public class NumberToTextConverter
    {

        private const long EXCEL_NAN_BITS = unchecked((long)0xFFFF0420003C0000L);
        private const int MAX_TEXT_LEN = 20;

        private NumberToTextConverter()
        {
            // no instances of this class
        }

        /**
         * Converts the supplied <c>value</c> to the text representation that Excel would give if
         * the value were to appear in an unformatted cell, or as a literal number in a formula.<br/>
         * Note - the results from this method differ slightly from those of <c>Double.ToString()</c>
         * In some special cases Excel behaves quite differently.  This function attempts to reproduce
         * those results.
         */
        public static String ToText(double value)
        {
            return RawDoubleBitsToText(BitConverter.DoubleToInt64Bits(value));
        }
        /* namespace */
        public static String RawDoubleBitsToText(long pRawBits)
        {

            long rawBits = pRawBits;
            bool isNegative = rawBits < 0; // sign bit is in the same place for long and double
            if (isNegative)
            {
                rawBits &= 0x7FFFFFFFFFFFFFFFL;
            }
            if (rawBits == 0)
            {
                return isNegative ? "-0" : "0";
            }
            ExpandedDouble ed = new ExpandedDouble(rawBits);
            if (ed.GetBinaryExponent() < -1022)
            {
                // value is 'denormalised' which means it is less than 2^-1022
                // excel displays all these numbers as zero, even though calculations work OK
                return isNegative ? "-0" : "0";
            }
            if (ed.GetBinaryExponent() == 1024)
            {
                // Special number NaN /InfInity
                // Normally one would not create HybridDecimal objects from these values
                // except in these cases Excel really tries to render them as if they were normal numbers
                if (rawBits == EXCEL_NAN_BITS)
                {
                    return "3.484840871308E+308";
                }
                // This is where excel really Gets it wrong
                // Special numbers like InfInity and NaN are interpreted according to
                // the standard rules below.
                isNegative = false; // except that the sign bit is ignored
            }
            NormalisedDecimal nd = ed.NormaliseBaseTen();
            StringBuilder sb = new StringBuilder(MAX_TEXT_LEN + 1);
            if (isNegative)
            {
                sb.Append('-');
            }
            ConvertToText(sb, nd);
            return sb.ToString();
        }
        private static void ConvertToText(StringBuilder sb, NormalisedDecimal pnd)
        {
            NormalisedDecimal rnd = pnd.RoundUnits();
            int decExponent = rnd.GetDecimalExponent();
            String decimalDigits;
            if (Math.Abs(decExponent) > 98)
            {
                decimalDigits = rnd.GetSignificantDecimalDigitsLastDigitRounded();
                if (decimalDigits.Length == 16)
                {
                    // rounding caused carry
                    decExponent++;
                }
            }
            else
            {
                decimalDigits = rnd.GetSignificantDecimalDigits();
            }
            int countSigDigits = CountSignifantDigits(decimalDigits);
            if (decExponent < 0)
            {
                FormatLessThanOne(sb, decimalDigits, decExponent, countSigDigits);
            }
            else
            {
                FormatGreaterThanOne(sb, decimalDigits, decExponent, countSigDigits);
            }
        }

        private static void FormatLessThanOne(StringBuilder sb, String decimalDigits, int decExponent,
                int countSigDigits)
        {
            int nLeadingZeros = -decExponent - 1;
            int normalLength = 2 + nLeadingZeros + countSigDigits; // 2 == "0.".Length

            if (NeedsScientificNotation(normalLength))
            {
                sb.Append(decimalDigits[0]);
                if (countSigDigits > 1)
                {
                    sb.Append('.');
                    sb.Append(decimalDigits.Substring(1, countSigDigits-1));
                }
                sb.Append("E-");
                AppendExp(sb, -decExponent);
                return;
            }
            sb.Append("0.");
            for (int i = nLeadingZeros; i > 0; i--)
            {
                sb.Append('0');
            }
            sb.Append(decimalDigits.Substring(0, countSigDigits));
        }

        private static void FormatGreaterThanOne(StringBuilder sb, String decimalDigits, int decExponent, int countSigDigits)
        {

            if (decExponent > 19)
            {
                // scientific notation
                sb.Append(decimalDigits[0]);
                if (countSigDigits > 1)
                {
                    sb.Append('.');
                    sb.Append(decimalDigits.Substring(1, countSigDigits-1));
                }
                sb.Append("E+");
                AppendExp(sb, decExponent);
                return;
            }
            int nFractionalDigits = countSigDigits - decExponent - 1;
            if (nFractionalDigits > 0)
            {
                sb.Append(decimalDigits.Substring(0, decExponent + 1));
                sb.Append('.');
                sb.Append(decimalDigits.Substring(decExponent + 1, nFractionalDigits));
                return;
            }
            sb.Append(decimalDigits.Substring(0, countSigDigits));
            for (int i = -nFractionalDigits; i > 0; i--)
            {
                sb.Append('0');
            }
        }

        private static bool NeedsScientificNotation(int nDigits)
        {
            return nDigits > MAX_TEXT_LEN;
        }

        private static int CountSignifantDigits(String sb)
        {
            int result = sb.Length - 1;
            while (sb[result] == '0')
            {
                result--;
                if (result < 0)
                {
                    throw new Exception("No non-zero digits found");
                }
            }
            return result + 1;
        }

        private static void AppendExp(StringBuilder sb, int val)
        {
            if (val < 10)
            {
                sb.Append('0');
                sb.Append((char)('0' + val));
                return;
            }
            sb.Append(val);
        }
    }
}
