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
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.SS.Util;

namespace NPOI.SS.UserModel
{
    /**
     * A wrapper around a {@link SimpleDateFormat} instance,
     * which handles a few Excel-style extensions that
     * are not supported by {@link SimpleDateFormat}.
     * Currently, the extensions are around the handling
     * of elapsed time, eg rendering 1 day 2 hours
     * as 26 hours.
     */
    public class ExcelStyleDateFormatter : SimpleDateFormat
    {
        public const char MMMMM_START_SYMBOL = '\ue001';
        public const char MMMMM_TRUNCATE_SYMBOL = '\ue002';
        public const char H_BRACKET_SYMBOL = '\ue010';
        public const char HH_BRACKET_SYMBOL = '\ue011';
        public const char M_BRACKET_SYMBOL = '\ue012';
        public const char MM_BRACKET_SYMBOL = '\ue013';
        public const char S_BRACKET_SYMBOL = '\ue014';
        public const char SS_BRACKET_SYMBOL = '\ue015';
        public const char L_BRACKET_SYMBOL = '\ue016';
        public const char LL_BRACKET_SYMBOL = '\ue017';
        public const char QUOTE_SYMBOL = '\ue009'; //add for C# DateTime format

        private static DecimalFormat format1digit = new DecimalFormat("0");
        private static DecimalFormat format2digits = new DecimalFormat("00");

        private static DecimalFormat format3digit = new DecimalFormat("0");
        private static DecimalFormat format4digits = new DecimalFormat("00");

        static ExcelStyleDateFormatter()
        {
            //DataFormatter.SetExcelStyleRoundingMode(format1digit, RoundingMode.DOWN);
            //DataFormatter.SetExcelStyleRoundingMode(format2digits, RoundingMode.DOWN);
            //DataFormatter.SetExcelStyleRoundingMode(format3digit);
            //DataFormatter.SetExcelStyleRoundingMode(format4digits);
        }

        private double dateToBeFormatted = 0.0;

        public ExcelStyleDateFormatter()
            : base()
        {
        }

        public ExcelStyleDateFormatter(String pattern)
            : base(ProcessFormatPattern(pattern))
        {

        }


        //public ExcelStyleDateFormatter(String pattern,
        //                           DateFormatSymbols formatSymbols)
        //{
        //    super(processFormatPattern(pattern), formatSymbols);
        //}

        //public ExcelStyleDateFormatter(String pattern, Locale locale)
        //{
        //    super(processFormatPattern(pattern), locale);
        //}
        private static string DateTimeMatchEvaluator(Match match)
        {
            return match.Groups[1].Value;
        }
        /**
         * Takes a format String, and Replaces Excel specific bits
         * with our detection sequences
         */
        private static String ProcessFormatPattern(String f)
        {
            String t = f.Replace("MMMMM", MMMMM_START_SYMBOL + "MMM" + MMMMM_TRUNCATE_SYMBOL);
            t = Regex.Replace(t, "\\[H\\]", (H_BRACKET_SYMBOL).ToString(), RegexOptions.IgnoreCase);
            t = Regex.Replace(t, "\\[HH\\]", (HH_BRACKET_SYMBOL).ToString(), RegexOptions.IgnoreCase);
            t = Regex.Replace(t, "\\[m\\]", (M_BRACKET_SYMBOL).ToString(), RegexOptions.IgnoreCase);
            t = Regex.Replace(t, "\\[mm\\]", (MM_BRACKET_SYMBOL).ToString(), RegexOptions.IgnoreCase);
            t = Regex.Replace(t, "\\[s\\]", (S_BRACKET_SYMBOL).ToString(), RegexOptions.IgnoreCase);
            t = Regex.Replace(t, "\\[ss\\]", (SS_BRACKET_SYMBOL).ToString(), RegexOptions.IgnoreCase);
            t = t.Replace("s.000", "s.fff");
            t = t.Replace("s.00", "s." + LL_BRACKET_SYMBOL);
            t = t.Replace("s.0", "s." + L_BRACKET_SYMBOL);
            t = t.Replace("\"", QUOTE_SYMBOL.ToString());
            //only one char 'M'
            //see http://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx#UsingSingleSpecifiers
            t = Regex.Replace(t, "(?<![M%])M(?!M)", "%M");
            return t;
        }

        /**
         * Used to let us know what the date being
         * formatted is, in Excel terms, which we
         * may wish to use when handling elapsed
         * times.
         */
        public void SetDateToBeFormatted(double date)
        {
            this.dateToBeFormatted = date;
        }
        public override string Format(object obj, CultureInfo culture)
        {
            return this.Format((DateTime)obj, new StringBuilder(), culture).ToString();
        }

        public StringBuilder Format(DateTime date, StringBuilder paramStringBuilder, CultureInfo culture)
        {
            // Do the normal format
            string s = string.Empty;
            if (Regex.IsMatch(Pattern, "[yYmMdDhHsS\\-/,. :\"\\\\]+0?[ampAMP/]*"))
            {
                s = date.ToString(Pattern, culture);
            }
            else
                s = Pattern;
            if (s.IndexOf(QUOTE_SYMBOL) != -1)
            {
                s = s.Replace(QUOTE_SYMBOL, '"');
            }
            // Now handle our special cases
            if (s.IndexOf(MMMMM_START_SYMBOL) != -1)
            {
                Regex reg = new Regex(MMMMM_START_SYMBOL + "(\\w)\\w+" + MMMMM_TRUNCATE_SYMBOL, RegexOptions.IgnoreCase);
                Match m = reg.Match(s);
                if (m.Success)
                {
                    s = reg.Replace(s, m.Groups[1].Value);
                }
            }

            if (s.IndexOf(H_BRACKET_SYMBOL) != -1 ||
                    s.IndexOf(HH_BRACKET_SYMBOL) != -1)
            {
                double hours = dateToBeFormatted * 24 + 0.01;
                //get the hour part of the time
                hours = Math.Floor(hours);
                s = s.Replace(
                        (H_BRACKET_SYMBOL).ToString(),
                        format1digit.Format(hours, culture)
                );
                s = s.Replace(
                        (HH_BRACKET_SYMBOL).ToString(),
                        format2digits.Format(hours, culture)
                );
            }

            if (s.IndexOf(M_BRACKET_SYMBOL) != -1 ||
                    s.IndexOf(MM_BRACKET_SYMBOL) != -1)
            {
                double minutes = dateToBeFormatted * 24 * 60 + 0.01;
                minutes = Math.Floor(minutes);
                s = s.Replace(
                        (M_BRACKET_SYMBOL).ToString(),
                        format1digit.Format(minutes, culture)
                );
                s = s.Replace(
                        (MM_BRACKET_SYMBOL).ToString(),
                        format2digits.Format(minutes, culture)
                );
            }
            if (s.IndexOf(S_BRACKET_SYMBOL) != -1 ||
                    s.IndexOf(SS_BRACKET_SYMBOL) != -1)
            {
                double seconds = (dateToBeFormatted * 24.0 * 60.0 * 60.0) + 0.01;
                s = s.Replace(
                        (S_BRACKET_SYMBOL).ToString(),
                        format1digit.Format(seconds, culture)
                );
                s = s.Replace(
                        (SS_BRACKET_SYMBOL).ToString(),
                        format2digits.Format(seconds, culture)
                );
            }

            if (s.IndexOf(L_BRACKET_SYMBOL) != -1 ||
                    s.IndexOf(LL_BRACKET_SYMBOL) != -1)
            {
                float millisTemp = (float)((dateToBeFormatted - Math.Floor(dateToBeFormatted)) * 24.0 * 60.0 * 60.0);
                float millis = (millisTemp - (int)millisTemp);
                s = s.Replace(
                        (L_BRACKET_SYMBOL).ToString(),
                        format3digit.Format(millis * 10, culture)
                );
                s = s.Replace(
                        (LL_BRACKET_SYMBOL).ToString(),
                        format4digits.Format(millis * 100, culture)
                );
            }

            return new StringBuilder(s);
        }

        public override bool Equals(Object o)
        {
            if (!(o is ExcelStyleDateFormatter)) {
                return false;
            }

            ExcelStyleDateFormatter other = (ExcelStyleDateFormatter)o;
            return dateToBeFormatted == other.dateToBeFormatted;
        }

        public override int GetHashCode()
        {
            return dateToBeFormatted.GetHashCode();
        }
    }

}