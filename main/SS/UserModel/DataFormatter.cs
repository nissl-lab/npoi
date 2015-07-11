/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional inFormation regarding copyright ownership.
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

namespace NPOI.SS.UserModel
{
    using System;
    using System.Collections;
    using System.Text;
    using System.Text.RegularExpressions;

    using NPOI.SS.Util;
    using System.Globalization;



    /**
     * HSSFDataFormatter contains methods for Formatting the value stored in an
     * Cell. This can be useful for reports and GUI presentations when you
     * need to display data exactly as it appears in Excel. Supported Formats
     * include currency, SSN, percentages, decimals, dates, phone numbers, zip
     * codes, etc.
     * 
     * Internally, Formats will be implemented using subclasses of <see cref="NPOI.SS.Util.FormatBase"/>
     * such as <see cref="NPOI.SS.Util.DecimalFormat"/> and <see cref="NPOI.SS.Util.SimpleDateFormat"/>. Therefore the
     * Formats used by this class must obey the same pattern rules as these FormatBase
     * subclasses. This means that only legal number pattern characters ("0", "#",
     * ".", "," etc.) may appear in number formats. Other characters can be
     * inserted <em>before</em> or <em> after</em> the number pattern to form a
     * prefix or suffix.
     * 
     * 
     * For example the Excel pattern <c>"$#,##0.00 "USD"_);($#,##0.00 "USD")"
     * </c> will be correctly Formatted as "$1,000.00 USD" or "($1,000.00 USD)".
     * However the pattern <c>"00-00-00"</c> is incorrectly Formatted by
     * DecimalFormat as "000000--". For Excel Formats that are not compatible with
     * DecimalFormat, you can provide your own custom {@link FormatBase} implementation
     * via <c>HSSFDataFormatter.AddFormat(String,FormatBase)</c>. The following
     * custom Formats are already provided by this class:
     * 
     * <pre>
     * <ul><li>SSN "000-00-0000"</li>
     *     <li>Phone Number "(###) ###-####"</li>
     *     <li>Zip plus 4 "00000-0000"</li>
     * </ul>
     * </pre>
     * 
     * If the Excel FormatBase pattern cannot be Parsed successfully, then a default
     * FormatBase will be used. The default number FormatBase will mimic the Excel General
     * FormatBase: "#" for whole numbers and "#.##########" for decimal numbers. You
     * can override the default FormatBase pattern with <c>
     * HSSFDataFormatter.setDefaultNumberFormat(FormatBase)</c>. <b>Note:</b> the
     * default FormatBase will only be used when a FormatBase cannot be Created from the
     * cell's data FormatBase string.
     *
     * @author James May (james dot may at fmr dot com)
     *
     */
    public class DataFormatter
    {
        private static String defaultFractionWholePartFormat = "#";
        private static String defaultFractionFractionPartFormat = "#/##";
        /** Pattern to find a number FormatBase: "0" or  "#" */
        private static string numPattern = "[0#]+";

        /** Pattern to find days of week as text "ddd...." */
        private static string daysAsText = "([d]{3,})";

        /** Pattern to find "AM/PM" marker */
        private static string amPmPattern = "((A|P)[M/P]*)";

        /** A regex to find patterns like [$$-1009] and [$�-452]. 
         *  Note that we don't currently process these into locales 
         */
        private static string localePatternGroup = "(\\[\\$[^-\\]]*-[0-9A-Z]+\\])";

        /*
     * A regex to match the colour formattings rules.
     * Allowed colours are: Black, Blue, Cyan, Green,
     *  Magenta, Red, White, Yellow, "Color n" (1<=n<=56)
     */
        private static Regex colorPattern = new Regex("(\\[BLACK\\])|(\\[BLUE\\])|(\\[CYAN\\])|(\\[GREEN\\])|" +
            "(\\[MAGENTA\\])|(\\[RED\\])|(\\[WHITE\\])|(\\[YELLOW\\])|" +
            "(\\[COLOR\\s*\\d\\])|(\\[COLOR\\s*[0-5]\\d\\])", RegexOptions.IgnoreCase);

        /**
         * A regex to identify a fraction pattern.
         * This requires that replaceAll("\\?", "#") has already been called 
         */
        private static Regex fractionPattern = new Regex("(?:([#\\d]+)\\s+)?(#+)\\s*\\/\\s*([#\\d]+)");

        /**
         * A regex to strip junk out of fraction formats
         */
        private static Regex fractionStripper = new Regex("(\"[^\"]*\")|([^ \\?#\\d\\/]+)");
        /**
      * Cells formatted with a date or time format and which contain invalid date or time values
     *  show 255 pound signs ("#").
      */
        private static String invalidDateTimeString;
        static DataFormatter()
        {
            StringBuilder buf = new StringBuilder();
            for (int i = 0; i < 255; i++) buf.Append('#');
            invalidDateTimeString = buf.ToString();
        }
        /** <em>General</em> FormatBase for whole numbers. */
        private static DecimalFormat generalWholeNumFormat = new DecimalFormat("0");

        /** <em>General</em> FormatBase for decimal numbers. */
        private static DecimalFormat generalDecimalNumFormat = new DecimalFormat("#.##########");

        /** A default FormatBase to use when a number pattern cannot be Parsed. */
        private FormatBase defaultNumFormat;
        private CultureInfo currentCulture;
        /*
         * A map to cache formats.
         *  Map<String,FormatBase> Formats
         */
        private Hashtable formats;
        private bool emulateCsv = false;
        /**
     * Creates a formatter using the {@link Locale#getDefault() default locale}.
     */
        public DataFormatter()
            : this(false)
        {
        }
        /**
         * Constructor
         */
        public DataFormatter(CultureInfo culture)
        {
            this.currentCulture = culture;
            formats = new Hashtable();

            // init built-in Formats

            FormatBase zipFormat = ZipPlusFourFormat.instance;
            AddFormat("00000\\-0000", zipFormat);
            AddFormat("00000-0000", zipFormat);

            FormatBase phoneFormat = PhoneFormat.instance;
            // allow for FormatBase string variations
            AddFormat("[<=9999999]###\\-####;\\(###\\)\\ ###\\-####", phoneFormat);
            AddFormat("[<=9999999]###-####;(###) ###-####", phoneFormat);
            AddFormat("###\\-####;\\(###\\)\\ ###\\-####", phoneFormat);
            AddFormat("###-####;(###) ###-####", phoneFormat);

            FormatBase ssnFormat = SSNFormat.instance;
            AddFormat("000\\-00\\-0000", ssnFormat);
            AddFormat("000-00-0000", ssnFormat);
        }
        public DataFormatter(bool emulateCsv)
            : this(CultureInfo.CurrentCulture)
        {
            this.emulateCsv = emulateCsv;
        }
        /**
         * Creates a formatter using the given locale.
         *
         * @param  emulateCsv whether to emulate CSV output.
         */
        public DataFormatter(CultureInfo locale, bool emulateCsv)
            : this(locale)
        {

            this.emulateCsv = emulateCsv;
        }
        /**
         * Return a FormatBase for the given cell if one exists, otherwise try to
         * Create one. This method will return <c>null</c> if the any of the
         * following is true:
         * <ul>
         * <li>the cell's style is null</li>
         * <li>the style's data FormatBase string is null or empty</li>
         * <li>the FormatBase string cannot be recognized as either a number or date</li>
         * </ul>
         *
         * @param cell The cell to retrieve a FormatBase for
         * @return A FormatBase for the FormatBase String
         */
        private FormatBase GetFormat(ICell cell)
        {
            if (cell.CellStyle == null)
            {
                return null;
            }

            int formatIndex = cell.CellStyle.DataFormat;
            String formatStr = cell.CellStyle.GetDataFormatString();
            if (formatStr == null || formatStr.Trim().Length == 0)
            {
                return null;
            }
            return GetFormat(cell.NumericCellValue, formatIndex, formatStr);
        }

        private FormatBase GetFormat(double cellValue, int formatIndex, String formatStrIn)
        {
            //      // Might be better to separate out the n p and z formats, falling back to p when n and z are not set.
            //      // That however would require other code to be re factored.
            //      String[] formatBits = formatStrIn.split(";");
            //      int i = cellValue > 0.0 ? 0 : cellValue < 0.0 ? 1 : 2; 
            //      String formatStr = (i < formatBits.length) ? formatBits[i] : formatBits[0];

            String formatStr = formatStrIn;
            // Excel supports positive/negative/zero, but java
            // doesn't, so we need to do it specially
            int firstAt = formatStr.IndexOf(';');
            int lastAt = formatStr.LastIndexOf(';');
            // p and p;n are ok by default. p;n;z and p;n;z;s need to be fixed.
            if (firstAt != -1 && firstAt != lastAt)
            {
                int secondAt = formatStr.IndexOf(';', firstAt + 1);
                if (secondAt == lastAt)
                { // p;n;z
                    if (cellValue == 0.0)
                    {
                        formatStr = formatStr.Substring(lastAt + 1);
                    }
                    else
                    {
                        formatStr = formatStr.Substring(0, lastAt);
                    }
                }
                else
                {
                    if (cellValue == 0.0)
                    { // p;n;z;s
                        formatStr = formatStr.Substring(secondAt + 1, lastAt - (secondAt + 1));
                    }
                    else
                    {
                        formatStr = formatStr.Substring(0, secondAt);
                    }
                }
            }

            // Excel's # with value 0 will output empty where Java will output 0. This hack removes the # from the format.
            if (emulateCsv && cellValue == 0.0 && formatStr.Contains("#") && !formatStr.Contains("0"))
            {
                formatStr = formatStr.Replace("#", "");
            }
            FormatBase format = (FormatBase)formats[formatStr];
            if (format != null)
            {
                return format;
            }
            // Is it one of the special built in types, General or @?
            if (formatStr.Equals("General", StringComparison.CurrentCultureIgnoreCase) || "@".Equals(formatStr))
            {
                if (DataFormatter.IsWholeNumber(cellValue))
                {
                    return generalWholeNumFormat;
                }
                return generalDecimalNumFormat;
            }

            // Build a formatter, and cache it
            format = CreateFormat(cellValue, formatIndex, formatStr);
            formats[formatStr] = format;
            return format;
        }

        /**
         * Create and return a FormatBase based on the FormatBase string from a  cell's
         * style. If the pattern cannot be Parsed, return a default pattern.
         *
         * @param cell The Excel cell
         * @return A FormatBase representing the excel FormatBase. May return null.
         */
        public FormatBase CreateFormat(ICell cell)
        {

            int formatIndex = cell.CellStyle.DataFormat;
            String formatStr = cell.CellStyle.GetDataFormatString();
            return CreateFormat(cell.NumericCellValue, formatIndex, formatStr);
        }

        private static readonly Regex RegexDoubleBackslashAny = new Regex("\\\\.");
        private static readonly Regex RegexContinueWs = new Regex("\\s");
        private static readonly Regex RegexAnyInDoubleQuote = new Regex("\"[^\"]*\"");

        private FormatBase CreateFormat(double cellValue, int formatIndex, String sFormat)
        {
            // remove color Formatting if present
            String formatStr = colorPattern.Replace(sFormat, "");

            // Strip off the locale information, we use an instance-wide locale for everything
            MatchCollection matches = Regex.Matches(formatStr, localePatternGroup);
            foreach (Match match in matches)
            {
                string matchedstring = match.Value;
                int beginpos = matchedstring.IndexOf('$') + 1;
                int endpos = matchedstring.IndexOf('-');
                String symbol = matchedstring.Substring(beginpos, endpos - beginpos);

                if (symbol.IndexOf('$') > -1)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(symbol.Substring(0, symbol.IndexOf('$')));
                    sb.Append('\\');
                    sb.Append(symbol.Substring(symbol.IndexOf('$'), symbol.Length));
                    symbol = sb.ToString();
                }
                matchedstring = Regex.Replace(matchedstring, localePatternGroup, symbol);

                formatStr = formatStr.Remove(match.Index, match.Length);
                formatStr = formatStr.Insert(match.Index, matchedstring);
            }

            // Check for special cases
            if (formatStr == null || formatStr.Trim().Length == 0)
            {
                return GetDefaultFormat(cellValue);
            }

            if ("General".Equals(formatStr, StringComparison.CurrentCultureIgnoreCase) || "@".Equals(formatStr))
            {
                if (IsWholeNumber(cellValue))
                {
                    return generalWholeNumFormat;
                }
                return generalDecimalNumFormat;
            }


            if (DateUtil.IsADateFormat(formatIndex, formatStr) &&
                    DateUtil.IsValidExcelDate(cellValue))
            {
                return CreateDateFormat(formatStr, cellValue);
            }

            // Excel supports fractions in format strings, which Java doesn't
            if (formatStr.IndexOf("#/") >= 0 || formatStr.IndexOf("?/") >= 0)
            {
                String[] chunks = formatStr.Split(";".ToCharArray());
                for (int i = 0; i < chunks.Length; i++)
                {
                    String chunk = chunks[i].Replace("?", "#");
                    //Match matcher = fractionStripper.Match(chunk);
                    //chunk = matcher.Replace(" ");
                    chunk = fractionStripper.Replace(chunk, " ");
                    chunk = chunk.Replace(" +", " ");
                    Match fractionMatcher = fractionPattern.Match(chunk);
                    //take the first match
                    if (fractionMatcher.Success)
                    {
                        String wholePart = (fractionMatcher.Groups[1] == null||!fractionMatcher.Groups[1].Success) ? "" : defaultFractionWholePartFormat;
                        return new FractionFormat(wholePart, fractionMatcher.Groups[3].Value);
                    }
                }

                // Strip custom text in quotes and escaped characters for now as it can cause performance problems in fractions.
                //String strippedFormatStr = formatStr.replaceAll("\\\\ ", " ").replaceAll("\\\\.", "").replaceAll("\"[^\"]*\"", " ").replaceAll("\\?", "#");
                //System.out.println("formatStr: "+strippedFormatStr);
                return new FractionFormat(defaultFractionWholePartFormat, defaultFractionFractionPartFormat);
            }


            if (Regex.IsMatch(formatStr, numPattern))
            {
                return CreateNumberFormat(formatStr, cellValue);
            }
            if (emulateCsv)
            {
                return new ConstantStringFormat(cleanFormatForNumber(formatStr));
            }
            // TODO - when does this occur?
            return null;
        }
        private int IndexOfFraction(String format)
        {
            int i = format.IndexOf("#/#");
            int j = format.IndexOf("?/?");
            return i == -1 ? j : j == -1 ? i : Math.Min(i, j);
        }

        private int LastIndexOfFraction(String format)
        {
            int i = format.LastIndexOf("#/#");
            int j = format.LastIndexOf("?/?");
            return i == -1 ? j : j == -1 ? i : Math.Max(i, j);
        }
        private FormatBase CreateDateFormat(String pformatStr, double cellValue)
        {
            String formatStr = pformatStr;
            formatStr = formatStr.Replace("\\-", "-");
            formatStr = formatStr.Replace("\\,", ",");
            formatStr = formatStr.Replace("\\.", ".");
            formatStr = formatStr.Replace("\\ ", " ");
            formatStr = formatStr.Replace("\\/", "/");
            formatStr = formatStr.Replace(";@", "");
            formatStr = formatStr.Replace("\"/\"", "/"); // "/" is escaped for no reason in: mm"/"dd"/"yyyy
            formatStr = formatStr.Replace("\"\"", "'");	// replace Excel quoting with Java style quoting
            formatStr = formatStr.Replace("\\\\T", "'T'"); // Quote the T is iso8601 style dates

            bool hasAmPm = Regex.IsMatch(formatStr, amPmPattern);
            if (hasAmPm)
            {
                formatStr = Regex.Replace(formatStr, amPmPattern, "@");
            }
            formatStr = formatStr.Replace("@", "tt");


            MatchCollection match = Regex.Matches(formatStr, daysAsText);
            if (match.Count > 0)
            {
                string replacement = match[0].Groups[0].Value.ToUpper().Replace("D", "E");
                formatStr = Regex.Replace(formatStr, daysAsText, replacement);
            }


            // Convert excel date FormatBase to SimpleDateFormat.
            // Excel uses lower case 'm' for both minutes and months.
            // From Excel help:
            /*
              The "m" or "mm" code must appear immediately after the "h" or"hh"
              code or immediately before the "ss" code; otherwise, Microsoft
              Excel displays the month instead of minutes."
            */

            StringBuilder sb = new StringBuilder();
            char[] chars = formatStr.ToCharArray();
            bool mIsMonth = true;
            bool isElapsed = false;
            ArrayList ms = new ArrayList();
            for (int j = 0; j < chars.Length; j++)
            {
                char c = chars[j];
                if (c == '\'')
                {
                    sb.Append(c);
                    j++;

                    // skip until the next quote
                    while (j < chars.Length)
                    {
                        c = chars[j];
                        sb.Append(c);
                        if (c == '\'')
                        {
                            break;
                        }
                        j++;
                    }
                }
                else if (c == '[' && !isElapsed)
                {
                    isElapsed = true;
                    mIsMonth = false;
                    sb.Append(c);
                }
                else if (c == ']' && isElapsed)
                {
                    isElapsed = false;
                    sb.Append(c);
                }
                else if (isElapsed)
                {
                    if (c == 'h' || c == 'H')
                    {
                        sb.Append('H');
                    }
                    else if (c == 'm' || c == 'M')
                    {
                        sb.Append('m');
                    }
                    else if (c == 's' || c == 'S')
                    {
                        sb.Append('s');
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                else if (c == 'h' || c == 'H')
                {
                    mIsMonth = false;
                    if (hasAmPm)
                    {
                        sb.Append('h');
                    }
                    else
                    {
                        sb.Append('H');
                    }
                }
                else if (c == 'm' || c == 'M')
                {
                    if (mIsMonth)
                    {
                        sb.Append('M');
                        ms.Add(
                                sb.Length - 1
                        );
                    }
                    else
                    {
                        sb.Append('m');
                    }
                }
                else if (c == 's' || c == 'S')
                {
                    sb.Append('s');
                    // if 'M' precedes 's' it should be minutes ('m')
                    for (int i = 0; i < ms.Count; i++)
                    {
                        int index = (int)ms[i];
                        if (sb[index] == 'M')
                        {
                            sb[index] = 'm';
                        }
                    }
                    mIsMonth = true;
                    ms.Clear();
                }
                else if (Char.IsLetter(c))
                {
                    mIsMonth = true;
                    ms.Clear();
                    if (c == 'y' || c == 'Y')
                    {
                        sb.Append('y');
                    }
                    else if (c == 'd' || c == 'D')
                    {
                        sb.Append('d');
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }
            formatStr = sb.ToString();

            try
            {
                //return new SimpleDateFormat(formatStr);
                return new ExcelStyleDateFormatter(formatStr);
            }
            catch (ArgumentException)
            {

                // the pattern could not be Parsed correctly,
                // so fall back to the default number FormatBase
                return GetDefaultFormat(cellValue);
            }

        }

        private String cleanFormatForNumber(String formatStr)
        {
            StringBuilder sb = new StringBuilder(formatStr);

            if (emulateCsv)
            {
                // Requested spacers with "_" are replaced by a single space.
                // Full-column-width padding "*" are removed.
                // Not processing fractions at this time. Replace ? with space.
                // This matches CSV output.
                for (int i = 0; i < sb.Length; i++)
                {
                    char c = sb[i];
                    if (c == '_' || c == '*' || c == '?')
                    {
                        if (i > 0 && sb[i - 1] == '\\')
                        {
                            // It's escaped, don't worry
                            continue;
                        }
                        if (c == '?')
                        {
                            sb[i] = ' ';
                        }
                        else if (i < sb.Length - 1)
                        {
                            // Remove the character we're supposed
                            //  to match the space of / pad to the
                            //  column width with
                            if (c == '_')
                            {
                                sb[i + 1] = ' ';
                            }
                            else
                            {
                                sb.Remove(i + 1, 1);
                            }
                            // Remove the character too
                            sb.Remove(i, 1);
                            i--;
                        }
                    }
                }
            }
            else
            {
                // If they requested spacers, with "_",
                //  remove those as we don't do spacing
                // If they requested full-column-width
                //  padding, with "*", remove those too
                for (int i = 0; i < sb.Length; i++)
                {
                    char c = sb[i];
                    if (c == '_' || c == '*')
                    {
                        if (i > 0 && sb[i - 1] == '\\')
                        {
                            // It's escaped, don't worry
                            continue;
                        }
                        if (i < sb.Length - 1)
                        {
                            // Remove the character we're supposed
                            //  to match the space of / pad to the
                            //  column width with
                            sb.Remove(i + 1, 1);
                        }
                        // Remove the _ too
                        sb.Remove(i, 1);
                        i--;
                    }
                }
            }

            // Now, handle the other aspects like 
            //  quoting and scientific notation
            for (int i = 0; i < sb.Length; i++)
            {
                char c = sb[i];
                // remove quotes and back slashes
                if (c == '\\' || c == '"')
                {
                    sb.Remove(i, 1);
                    i--;

                    // for scientific/engineering notation
                }
                else if (c == '+' && i > 0 && sb[i - 1] == 'E')
                {
                    sb.Remove(i, 1);
                    i--;
                }
            }

            return sb.ToString();
        }
        private FormatBase CreateNumberFormat(String formatStr, double cellValue)
        {
            String format = cleanFormatForNumber(formatStr);
            try
            {
                return new DecimalFormat(format);
            }
            catch (ArgumentException)
            {

                // the pattern could not be Parsed correctly,
                // so fall back to the default number FormatBase
                return GetDefaultFormat(cellValue);
            }
        }

        /**
         * Return true if the double value represents a whole number
         * @param d the double value to check
         * @return <c>true</c> if d is a whole number
         */
        private static bool IsWholeNumber(double d)
        {
            return d == Math.Floor(d);
        }

        /**
         * Returns a default FormatBase for a cell.
         * @param cell The cell
         * @return a default FormatBase
         */
        public FormatBase GetDefaultFormat(ICell cell)
        {
            return GetDefaultFormat(cell.NumericCellValue);
        }
        private FormatBase GetDefaultFormat(double cellValue)
        {
            // for numeric cells try user supplied default
            if (defaultNumFormat != null)
            {
                return defaultNumFormat;

                // otherwise use general FormatBase
            }
            if (IsWholeNumber(cellValue))
            {
                return generalWholeNumFormat;
            }
            return generalDecimalNumFormat;
        }

        /**
         * Returns the Formatted value of an Excel date as a <c>String</c> based
         * on the cell's <c>DataFormat</c>. i.e. "Thursday, January 02, 2003"
         * , "01/02/2003" , "02-Jan" , etc.
         *
         * @param cell The cell
         * @return a Formatted date string
         */
        private String GetFormattedDateString(ICell cell)
        {
            FormatBase dateFormat = GetFormat(cell);
            DateTime d = cell.DateCellValue;
            if (dateFormat != null)
            {
                return dateFormat.Format(d, currentCulture);
            }
            return d.ToString();
        }

        /**
         * Returns the Formatted value of an Excel number as a <c>String</c>
         * based on the cell's <c>DataFormat</c>. Supported Formats include
         * currency, percents, decimals, phone number, SSN, etc.:
         * "61.54%", "$100.00", "(800) 555-1234".
         *
         * @param cell The cell
         * @return a Formatted number string
         */
        private String GetFormattedNumberString(ICell cell)
        {

            FormatBase numberFormat = GetFormat(cell);
            double d = cell.NumericCellValue;
            if (numberFormat == null)
            {
                return d.ToString(currentCulture);
            }
            return numberFormat.Format(d, currentCulture);
        }

        /**
         * Formats the given raw cell value, based on the supplied
         *  FormatBase index and string, according to excel style rules.
         * @see #FormatCellValue(Cell)
         */
        public String FormatRawCellContents(double value, int formatIndex, String formatString)
        {
            return FormatRawCellContents(value, formatIndex, formatString, false);
        }
        /**
     * Performs Excel-style date formatting, using the
     *  supplied Date and format
     */
        private String PerformDateFormatting(DateTime d, FormatBase dateFormat)
        {
            if (dateFormat != null)
            {
                return dateFormat.Format(d, currentCulture);
            }
            return d.ToString();
        }
        /**
     * Formats the given raw cell value, based on the supplied
     *  format index and string, according to excel style rules.
     * @see #formatCellValue(Cell)
     */
        public String FormatRawCellContents(double value, int formatIndex, String formatString, bool use1904Windowing)
        {
            // Is it a date?
            if (DateUtil.IsADateFormat(formatIndex, formatString))
            {
                if (DateUtil.IsValidExcelDate(value))
                {
                    FormatBase dateFormat = GetFormat(value, formatIndex, formatString);

                    if (dateFormat is ExcelStyleDateFormatter)
                    {
                        // Hint about the raw excel value
                        ((ExcelStyleDateFormatter)dateFormat).SetDateToBeFormatted(value);
                    }

                    DateTime d = DateUtil.GetJavaDate(value, use1904Windowing);
                    return PerformDateFormatting(d, dateFormat);
                }

                // RK: Invalid dates are 255 #s.
                if (emulateCsv)
                {
                    return invalidDateTimeString;
                }
            }
            // else Number
            FormatBase numberFormat = GetFormat(value, formatIndex, formatString);
            if (numberFormat == null)
            {
                return value.ToString(currentCulture);
            }
            // When formatting 'value', double to text to BigDecimal produces more
            // accurate results than double to Double in JDK8 (as compared to
            // previous versions). However, if the value contains E notation, this
            // would expand the values, which we do not want, so revert to
            // original method.
            String result;
            String textValue = NumberToTextConverter.ToText(value);
            if (textValue.IndexOf('E') > -1) {
                result = numberFormat.Format(value);
            }
            else {
                result = numberFormat.Format(textValue);
            }
            // Complete scientific notation by adding the missing +.
            if (result.Contains("E") && !result.Contains("E-"))
            {
                result = result.Replace("E", "E+");
            }
            return result;
        }
        /**
         * 
         * Returns the Formatted value of a cell as a <c>String</c> regardless
         * of the cell type. If the Excel FormatBase pattern cannot be Parsed then the
         * cell value will be Formatted using a default FormatBase.
         * 
         * When passed a null or blank cell, this method will return an empty
         * String (""). Formulas in formula type cells will not be evaluated.
         * 
         *
         * @param cell The cell
         * @return the Formatted cell value as a String
         */
        public String FormatCellValue(ICell cell)
        {
            return FormatCellValue(cell, null);
        }

        /**
         * 
         * Returns the Formatted value of a cell as a <c>String</c> regardless
         * of the cell type. If the Excel FormatBase pattern cannot be Parsed then the
         * cell value will be Formatted using a default FormatBase.
         * 
         * When passed a null or blank cell, this method will return an empty
         * String (""). Formula cells will be evaluated using the given
         * {@link HSSFFormulaEvaluator} if the evaluator is non-null. If the
         * evaluator is null, then the formula String will be returned. The caller
         * is responsible for setting the currentRow on the evaluator
         *
         *
         * @param cell The cell (can be null)
         * @param evaluator The HSSFFormulaEvaluator (can be null)
         * @return a string value of the cell
         */
        public String FormatCellValue(ICell cell, IFormulaEvaluator evaluator)
        {

            if (cell == null)
            {
                return "";
            }

            CellType cellType = cell.CellType;
            if (evaluator != null && cellType == CellType.Formula)
            {
                if (evaluator == null)
                {
                    return cell.CellFormula;
                }
                cellType = evaluator.EvaluateFormulaCell(cell);
            }
            switch (cellType)
            {
                case CellType.Formula:
                    // should only occur if evaluator is null
                    return cell.CellFormula;

                case CellType.Numeric:

                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return GetFormattedDateString(cell);
                    }
                    return GetFormattedNumberString(cell);

                case CellType.String:
                    return cell.RichStringCellValue.String;

                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString().ToUpper();
                case CellType.Blank:
                    return "";
                case CellType.Error:
                    return FormulaError.ForInt(cell.ErrorCellValue).String;
            }
            throw new Exception("Unexpected celltype (" + cellType + ")");
        }


        /**
         * 
         * Sets a default number FormatBase to be used when the Excel FormatBase cannot be
         * Parsed successfully. <b>Note:</b> This is a fall back for when an error
         * occurs while parsing an Excel number FormatBase pattern. This will not
         * affect cells with the <em>General</em> FormatBase.
         * 
         * 
         * The value that will be passed to the FormatBase's FormatBase method (specified
         * by <c>java.text.FormatBase#FormatBase</c>) will be a double value from a
         * numeric cell. Therefore the code in the FormatBase method should expect a
         * <c>Number</c> value.
         * 
         *
         * @param FormatBase A FormatBase instance to be used as a default
         * @see java.text.FormatBase#FormatBase
         */
        public void SetDefaultNumberFormat(FormatBase format)
        {
            IEnumerator itr = formats.Keys.GetEnumerator();
            while (itr.MoveNext())
            {
                string key = (string)itr.Current;
                if (formats[key] == generalDecimalNumFormat
                        || formats[key] == generalWholeNumFormat)
                {
                    formats[key] = format;
                }
            }
            defaultNumFormat = format;
        }

        /**
         * Adds a new FormatBase to the available formats.
         * 
         * The value that will be passed to the FormatBase's FormatBase method (specified
         * by <c>java.text.FormatBase#FormatBase</c>) will be a double value from a
         * numeric cell. Therefore the code in the FormatBase method should expect a
         * <c>Number</c> value.
         * 
         * @param excelformatStr The data FormatBase string
         * @param FormatBase A FormatBase instance
         */
        public void AddFormat(String excelformatStr, FormatBase format)
        {
            formats[excelformatStr] = format;
        }

        // Some custom Formats
    }
}