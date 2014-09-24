/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using NPOI.Util;

namespace NPOI.SS.Util
{
    /**
     *  Convert DateFormat patterns into Excel custom number formats.
     *  For example, to format a date in excel using the "dd MMMM, yyyy" pattern and Japanese
     *  locale, use the following code:
     *
     *  <pre><code>
     *      // returns "[$-0411]dd MMMM, yyyy;@" where the [$-0411] prefix tells Excel to use the Japanese locale
     *      String excelFormatPattern = DateFormatConverter.convert(Locale.JAPANESE, "dd MMMM, yyyy");
     *
     *      CellStyle cellStyle = workbook.createCellStyle();
     *
     *      DataFormat poiFormat = workbook.createDataFormat();
     *      cellStyle.setDataFormat(poiFormat.getFormat(excelFormatPattern));
     *      cell.setCellValue(new Date());
     *      cell.setCellStyle(cellStyle);  // formats date as '2012\u5e743\u670817\u65e5'
     *
     *  </code></pre>
     *
     *
     */
    public class DateFormatConverter
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(DateFormatConverter));

        public class DateFormatTokenizer
        {
            string format;
            int pos;

            public DateFormatTokenizer(string format)
            {
                this.format = format;
            }

            public string GetNextToken()
            {
                if (pos >= format.Length)
                {
                    return null;
                }
                int subStart = pos;
                char curChar = format[pos];
                ++pos;
                if (curChar == '\'')
                {
                    while ((pos < format.Length) && ((curChar = format[pos]) != '\''))
                    {
                        ++pos;
                    }
                    if (pos < format.Length)
                    {
                        ++pos;
                    }
                }
                else
                {
                    char activeChar = curChar;
                    while ((pos < format.Length) && ((curChar = format[pos])) == activeChar)
                    {
                        ++pos;
                    }
                }
                return format.Substring(subStart, pos - subStart);
            }

            public static string[] Tokenize(string format)
            {
                List<string> result = new List<string>();

                DateFormatTokenizer tokenizer = new DateFormatTokenizer(format);
                string token;
                while ((token = tokenizer.GetNextToken()) != null)
                {
                    result.Add(token);
                }

                return result.ToArray();
            }

            public override string ToString()
            {
                StringBuilder result = new StringBuilder();

                DateFormatTokenizer tokenizer = new DateFormatTokenizer(format);
                string token;
                while ((token = tokenizer.GetNextToken()) != null)
                {
                    if (result.Length > 0)
                    {
                        result.Append(", ");
                    }
                    result.Append("[").Append(token).Append("]");
                }

                return result.ToString();
            }
        }

        private static Dictionary<string, string> tokenConversions = PrepareTokenConversions();
        private static Dictionary<string, string> localePrefixes = PrepareLocalePrefixes();

        private static Dictionary<string, string> PrepareTokenConversions()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            result.Add("EEEE", "dddd");
            result.Add("EEE", "ddd");
            result.Add("EE", "ddd");
            result.Add("E", "d");
            result.Add("Z", "");
            result.Add("z", "");
            result.Add("a", "am/pm");
            result.Add("A", "AM/PM");
            result.Add("K", "H");
            result.Add("KK", "HH");
            result.Add("k", "h");
            result.Add("kk", "hh");
            result.Add("S", "0");
            result.Add("SS", "00");
            result.Add("SSS", "000");

            return result;
        }

        private static Dictionary<string, string> PrepareLocalePrefixes()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            result.Add("af", "[$-0436]");
            result.Add("am", "[$-45E]");
            result.Add("ar-ae", "[$-3801]");
            result.Add("ar-bh", "[$-3C01]");
            result.Add("ar-dz", "[$-1401]");
            result.Add("ar-eg", "[$-C01]");
            result.Add("ar-iq", "[$-0801]");
            result.Add("ar-jo", "[$-2C01]");
            result.Add("ar-kw", "[$-3401]");
            result.Add("ar-lb", "[$-3001]");
            result.Add("ar-ly", "[$-1001]");
            result.Add("ar-ma", "[$-1801]");
            result.Add("ar-om", "[$-2001]");
            result.Add("ar-qa", "[$-4001]");
            result.Add("ar-sa", "[$-0401]");
            result.Add("ar-sy", "[$-2801]");
            result.Add("ar-tn", "[$-1C01]");
            result.Add("ar-ye", "[$-2401]");
            result.Add("as", "[$-44D]");
            result.Add("az-az", "[$-82C]");
            //result.Add("az-az", "[$-42C]");
            result.Add("be", "[$-0423]");
            result.Add("bg", "[$-0402]");
            result.Add("bn", "[$-0845]");
            //result.Add("bn", "[$-0445]");
            result.Add("bo", "[$-0451]");
            result.Add("bs", "[$-141A]");
            result.Add("ca", "[$-0403]");
            result.Add("cs", "[$-0405]");
            result.Add("cy", "[$-0452]");
            result.Add("da", "[$-0406]");
            result.Add("de-at", "[$-C07]");
            result.Add("de-ch", "[$-0807]");
            result.Add("de-de", "[$-0407]");
            result.Add("de-li", "[$-1407]");
            result.Add("de-lu", "[$-1007]");
            result.Add("dv", "[$-0465]");
            result.Add("el", "[$-0408]");
            result.Add("en-au", "[$-C09]");
            result.Add("en-bz", "[$-2809]");
            result.Add("en-ca", "[$-1009]");
            result.Add("en-cb", "[$-2409]");
            result.Add("en-gb", "[$-0809]");
            result.Add("en-ie", "[$-1809]");
            result.Add("en-in", "[$-4009]");
            result.Add("en-jm", "[$-2009]");
            result.Add("en-nz", "[$-1409]");
            result.Add("en-ph", "[$-3409]");
            result.Add("en-tt", "[$-2C09]");
            result.Add("en-us", "[$-0409]");
            result.Add("en-za", "[$-1C09]");
            result.Add("es-ar", "[$-2C0A]");
            result.Add("es-bo", "[$-400A]");
            result.Add("es-cl", "[$-340A]");
            result.Add("es-co", "[$-240A]");
            result.Add("es-cr", "[$-140A]");
            result.Add("es-do", "[$-1C0A]");
            result.Add("es-ec", "[$-300A]");
            result.Add("es-es", "[$-40A]");
            result.Add("es-gt", "[$-100A]");
            result.Add("es-hn", "[$-480A]");
            result.Add("es-mx", "[$-80A]");
            result.Add("es-ni", "[$-4C0A]");
            result.Add("es-pa", "[$-180A]");
            result.Add("es-pe", "[$-280A]");
            result.Add("es-pr", "[$-500A]");
            result.Add("es-py", "[$-3C0A]");
            result.Add("es-sv", "[$-440A]");
            result.Add("es-uy", "[$-380A]");
            result.Add("es-ve", "[$-200A]");
            result.Add("et", "[$-0425]");
            result.Add("eu", "[$-42D]");
            result.Add("fa", "[$-0429]");
            result.Add("fi", "[$-40B]");
            result.Add("fo", "[$-0438]");
            result.Add("fr-be", "[$-80C]");
            result.Add("fr-ca", "[$-C0C]");
            result.Add("fr-ch", "[$-100C]");
            result.Add("fr-fr", "[$-40C]");
            result.Add("fr-lu", "[$-140C]");
            result.Add("gd", "[$-43C]");
            result.Add("gd-ie", "[$-83C]");
            result.Add("gn", "[$-0474]");
            result.Add("gu", "[$-0447]");
            result.Add("he", "[$-40D]");
            result.Add("hi", "[$-0439]");
            result.Add("hr", "[$-41A]");
            result.Add("hu", "[$-40E]");
            result.Add("hy", "[$-42B]");
            result.Add("id", "[$-0421]");
            result.Add("is", "[$-40F]");
            result.Add("it-ch", "[$-0810]");
            result.Add("it-it", "[$-0410]");
            result.Add("ja", "[$-0411]");
            result.Add("kk", "[$-43F]");
            result.Add("km", "[$-0453]");
            result.Add("kn", "[$-44B]");
            result.Add("ko", "[$-0412]");
            result.Add("ks", "[$-0460]");
            result.Add("la", "[$-0476]");
            result.Add("lo", "[$-0454]");
            result.Add("lt", "[$-0427]");
            result.Add("lv", "[$-0426]");
            result.Add("mi", "[$-0481]");
            result.Add("mk", "[$-42F]");
            result.Add("ml", "[$-44C]");
            result.Add("mn", "[$-0850]");
            //result.Add("mn", "[$-0450]");
            result.Add("mr", "[$-44E]");
            result.Add("ms-bn", "[$-83E]");
            result.Add("ms-my", "[$-43E]");
            result.Add("mt", "[$-43A]");
            result.Add("my", "[$-0455]");
            result.Add("ne", "[$-0461]");
            result.Add("nl-be", "[$-0813]");
            result.Add("nl-nl", "[$-0413]");
            result.Add("no-no", "[$-0814]");
            result.Add("or", "[$-0448]");
            result.Add("pa", "[$-0446]");
            result.Add("pl", "[$-0415]");
            result.Add("pt-br", "[$-0416]");
            result.Add("pt-pt", "[$-0816]");
            result.Add("rm", "[$-0417]");
            result.Add("ro", "[$-0418]");
            result.Add("ro-mo", "[$-0818]");
            result.Add("ru", "[$-0419]");
            result.Add("ru-mo", "[$-0819]");
            result.Add("sa", "[$-44F]");
            result.Add("sb", "[$-42E]");
            result.Add("sd", "[$-0459]");
            result.Add("si", "[$-45B]");
            result.Add("sk", "[$-41B]");
            result.Add("sl", "[$-0424]");
            result.Add("so", "[$-0477]");
            result.Add("sq", "[$-41C]");
            result.Add("sr-sp", "[$-C1A]");
            //result.Add("sr-sp", "[$-81A]");
            result.Add("sv-fi", "[$-81D]");
            result.Add("sv-se", "[$-41D]");
            result.Add("sw", "[$-0441]");
            result.Add("ta", "[$-0449]");
            result.Add("te", "[$-44A]");
            result.Add("tg", "[$-0428]");
            result.Add("th", "[$-41E]");
            result.Add("tk", "[$-0442]");
            result.Add("tn", "[$-0432]");
            result.Add("tr", "[$-41F]");
            result.Add("ts", "[$-0431]");
            result.Add("tt", "[$-0444]");
            result.Add("uk", "[$-0422]");
            result.Add("ur", "[$-0420]");
            result.Add("UTF-8", "[$-0000]");
            result.Add("uz-uz", "[$-0843]");
            //result.Add("uz-uz", "[$-0443]");
            result.Add("vi", "[$-42A]");
            result.Add("xh", "[$-0434]");
            result.Add("yi", "[$-43D]");
            result.Add("zh-cn", "[$-0804]");
            result.Add("zh-hk", "[$-C04]");
            result.Add("zh-mo", "[$-1404]");
            result.Add("zh-sg", "[$-1004]");
            result.Add("zh-tw", "[$-0404]");
            result.Add("zu", "[$-0435]");

            result.Add("ar", "[$-0401]");
            //result.Add("bn", "[$-0845]");
            result.Add("de", "[$-0407]");
            result.Add("en", "[$-0409]");
            result.Add("es", "[$-40A]");
            result.Add("fr", "[$-40C]");
            result.Add("it", "[$-0410]");
            result.Add("ms", "[$-43E]");
            result.Add("nl", "[$-0413]");
            result.Add("nn", "[$-0814]");
            result.Add("no", "[$-0414]");
            result.Add("pt", "[$-0816]");
            result.Add("sr", "[$-C1A]");
            result.Add("sv", "[$-41D]");
            result.Add("uz", "[$-0843]");
            result.Add("zh", "[$-0804]");

            result.Add("ga", "[$-43C]");
            result.Add("ga-ie", "[$-83C]");
            result.Add("in", "[$-0421]");
            result.Add("iw", "[$-40D]");
            return result;
        }

        public static string GetPrefixForLocale(CultureInfo locale)
        {
            string localeString = locale.ToString().ToLower();
            string result = localePrefixes[localeString];
            if (result == null)
            {
                result = localePrefixes[localeString.Substring(0, 2)];
                if (result == null)
                {
                    CultureInfo parentLocale = CultureInfo.GetCultureInfo(localeString.Substring(0, 2));
                    logger.Log(POILogger.ERROR, "Unable to find prefix for " + locale + "(" + locale.DisplayName + ") or "
                            + localeString.Substring(0, 2) + "(" + parentLocale.DisplayName + ")");
                    return "";
                }
            }
            return result;
        }

        //public static string Convert(CultureInfo locale, DateFormat df)
        //{
        //    string ptrn = ((SimpleDateFormat)df).ToPattern();
        //    return convert(locale, ptrn);
        //}

        public static string Convert(CultureInfo locale, string format)
        {
            StringBuilder result = new StringBuilder();

            result.Append(GetPrefixForLocale(locale));
            DateFormatTokenizer tokenizer = new DateFormatTokenizer(format);
            string token;
            while ((token = tokenizer.GetNextToken()) != null)
            {
                if (token.StartsWith("'"))
                {
                    result.Append(token.Replace("'", "\""));
                }
                else if (!char.IsLetter(token[0]))
                {
                    result.Append(token);
                }
                else
                {
                    // It's a code, translate it if necessary
                    string mappedToken = tokenConversions[(token)];
                    result.Append(mappedToken == null ? token : mappedToken);
                }
            }
            result.Append(";@");
            return result.ToString().Trim();
        }

        public static string GetDatePattern(int style, CultureInfo locale)
        {
            string pattern = DateFormat.GetDatePattern(style, locale);
            return pattern;
        }

        public static string GetTimePattern(int style, CultureInfo locale)
        {
            string pattern = DateFormat.GetTimePattern(style, locale);
            return pattern;
        }

        public static string GetDateTimePattern(int style, CultureInfo locale)
        {
            string pattern= DateFormat.GetDateTimePattern(style, style, locale);
            return pattern;
        }

    }

    public class DateFormat
    {
        public const int FULL = 0;
        public const int LONG = 1;
        public const int MEDIUM = 2;
        public const int SHORT = 3;
        public const int DEFAULT = MEDIUM;

        public static string GetDateTimePattern(int dateStyle, int timeStyle, CultureInfo locale)
        {
            DateTimeFormatInfo dfi = locale.DateTimeFormat;
            string datePattern = GetDatePattern(dateStyle,locale);
            string timePattern = GetTimePattern(timeStyle, locale);

            if (locale.TextInfo.IsRightToLeft)
                return timePattern + " " + datePattern;//Is this right???
            else
                return datePattern + " " + timePattern;
        }
        public static string GetDatePattern(int dateStyle, CultureInfo locale)
        {
            DateTimeFormatInfo dfi = locale.DateTimeFormat;
            switch (dateStyle)
            {
                case DateFormat.SHORT:
                    return dfi.ShortDatePattern.Replace("yyyy", "yy").Replace("YYYY", "YY");
                case DateFormat.MEDIUM:
                    return dfi.ShortDatePattern;
                case DateFormat.LONG:
                    return dfi.LongDatePattern.Replace("dddd,", "").Trim();
                case DateFormat.FULL:
                    return dfi.LongDatePattern;
                default:
                    return dfi.ShortDatePattern;
            }
        }
        public static string GetTimePattern(int timeStyle, CultureInfo locale)
        {
            DateTimeFormatInfo dfi = locale.DateTimeFormat;
            switch (timeStyle)
            {
                case DateFormat.SHORT:
                    return dfi.ShortTimePattern;
                case DateFormat.MEDIUM:
                case DateFormat.LONG:
                case DateFormat.FULL:
                default:
                    return dfi.LongTimePattern;
            }
        }
    }
}