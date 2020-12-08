/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for Additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

using NPOI.SS.Util;
using System.Text.RegularExpressions;
using System;
using System.Text;
using System.Globalization;
using NPOI.SS.Format;
namespace NPOI.SS.UserModel
{

    /**
     * <p>Format class that handles Excel style fractions, such as "# #/#" and "#/###"</p>
     * 
     * <p>As of this writing, this is still not 100% accurate, but it does a reasonable job
     * of trying to mimic Excel's fraction calculations.  It does not currently
     * maintain Excel's spacing.</p>
     * 
     * <p>This class relies on a method lifted nearly verbatim from org.apache.math.fraction.
     *  If further uses for Commons Math are found, we will consider Adding it as a dependency.
     *  For now, we have in-lined the one method to keep things simple.</p>
     */
    /* One question remains...Is the value of epsilon in calcFractionMaxDenom reasonable? */

    public class FractionFormat : FormatBase
    {
        private static Regex DENOM_FORMAT_PATTERN = new Regex("(?:(#+)|(\\d+))");

        //this was chosen to match the earlier limitation of max denom power
        //it can be expanded to Get closer to Excel's calculations
        //with custom formats # #/#########
        //but as of this writing, the numerators and denominators
        //with formats of that nature on very small values were quite
        //far from Excel's calculations
        private static int MAX_DENOM_POW = 4;

        //there are two options:
        //a) an exact denominator is specified in the formatString
        //b) the maximum denominator can be calculated from the formatString
        private int exactDenom;
        private int maxDenom;

        private String wholePartFormatString;
        /**
         * Single parameter ctor
         * @param denomFormatString The format string for the denominator
         */
        public FractionFormat(String wholePartFormatString, String denomFormatString)
        {
            this.wholePartFormatString = wholePartFormatString;
            //init exactDenom and maxDenom
            Match m = DENOM_FORMAT_PATTERN.Match(denomFormatString);
            int tmpExact = -1;
            int tmpMax = -1;
            if (m.Success)
            {
                if (m.Groups[2] != null && m.Groups[2].Success)
                {
                    try
                    {
                        tmpExact = Int32.Parse(m.Groups[2].Value);
                        //if the denom is 0, fall back to the default: tmpExact=100

                        if (tmpExact == 0)
                        {
                            tmpExact = -1;
                        }
                    }
                    catch (FormatException)
                    {
                        //should never happen
                    }
                }
                else if (m.Groups[1] != null && m.Groups[1].Success)
                {
                    int len = m.Groups[1].Value.Length;
                    len = len > MAX_DENOM_POW ? MAX_DENOM_POW : len;
                    tmpMax = (int)Math.Pow(10, len);
                }
                else
                {
                    tmpExact = 100;
                }
            }
            if (tmpExact <= 0 && tmpMax <= 0)
            {
                //use 100 as the default denom if something went horribly wrong
                tmpExact = 100;
            }
            exactDenom = tmpExact;
            maxDenom = tmpMax;
        }

        public String Format(string num)
        {

            double doubleValue = 0;
            double.TryParse(num, out doubleValue);

            bool isNeg = (doubleValue < 0.0f) ? true : false;
            double absDoubleValue = Math.Abs(doubleValue);

            double wholePart = Math.Floor(absDoubleValue);
            double decPart = absDoubleValue - wholePart;
            if (wholePart + decPart == 0)
            {
                return "0";
            }

            //if the absolute value is smaller than 1 over the exact or maxDenom
            //you can stop here and return "0"
            if (absDoubleValue < (1 / Math.Max(exactDenom, maxDenom)))
            {
                return "0";
            }

            //this is necessary to prevent overflow in the maxDenom calculation
            //stink1
            if (wholePart + (int)decPart == wholePart + decPart)
            {

                StringBuilder sb = new StringBuilder();
                if (isNeg)
                {
                    sb.Append("-");
                }
                sb.Append((int)wholePart);
                return sb.ToString();
            }

            SimpleFraction fract = null;
            try
            {
                //this should be the case because of the constructor
                if (exactDenom > 0)
                {
                    fract = SimpleFraction.BuildFractionExactDenominator(decPart, exactDenom);
                }
                else
                {
                    fract = SimpleFraction.BuildFractionMaxDenominator(decPart, maxDenom);
                }
            }
            catch (SimpleFractionException)
            {
                return doubleValue.ToString();
            }

            StringBuilder sb1 = new StringBuilder();

            //now format the results
            if (isNeg)
            {
                sb1.Append("-");
            }

            //if whole part has to go into the numerator
            if ("".Equals(wholePartFormatString))
            {
                int trueNum = (fract.Denominator * (int)wholePart) + fract.Numerator;
                sb1.Append(trueNum).Append("/").Append(fract.Denominator);
                return sb1.ToString();
            }


            //short circuit if fraction is 0 or 1
            if (fract.Numerator == 0)
            {
                sb1.Append((int)wholePart);
                return sb1.ToString();
            }
            else if (fract.Numerator == fract.Denominator)
            {
                sb1.Append((int)wholePart + 1);
                return sb1.ToString();
            }
            //as mentioned above, this ignores the exact space formatting in Excel
            if (wholePart > 0)
            {
                sb1.Append((int)wholePart).Append(" ");
            }
            sb1.Append(fract.Numerator).Append("/").Append(fract.Denominator);
            return sb1.ToString();
        }

        protected override StringBuilder Format(object obj, StringBuilder toAppendTo, int pos)
        {
            return toAppendTo.Append(Format(obj.ToString()));
        }

        public override string Format(object obj, CultureInfo culture)
        {
            return this.Format(obj.ToString());
        }

        public override object ParseObject(string source, int pos)
        {
            throw new NotImplementedException("Reverse parsing not supported");
        }
       
        private class SimpleFractionException : Exception
        {
            public SimpleFractionException(String message) :
                base(message)
            {
            }
        }

        public override StringBuilder Format(object obj, StringBuilder toAppendTo, CultureInfo culture)
        {
            return toAppendTo.Append(Format(obj, culture));
        }
    }

}