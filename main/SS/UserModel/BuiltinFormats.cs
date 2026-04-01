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
=================================================================== */
namespace NPOI.SS.UserModel
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;




    /**
     * Utility to identify built-in formats.  The following is a list of the formats as
     * returned by this class.<p/>
     *<p/>
     *       0, "General"<br/>
     *       1, "0"<br/>
     *       2, "0.00"<br/>
     *       3, "#,##0"<br/>
     *       4, "#,##0.00"<br/>
     *       5, "$#,##0_);($#,##0)"<br/>
     *       6, "$#,##0_);[Red]($#,##0)"<br/>
     *       7, "$#,##0.00);($#,##0.00)"<br/>
     *       8, "$#,##0.00_);[Red]($#,##0.00)"<br/>
     *       9, "0%"<br/>
     *       0xa, "0.00%"<br/>
     *       0xb, "0.00E+00"<br/>
     *       0xc, "# ?/?"<br/>
     *       0xd, "# ??/??"<br/>
     *       0xe, "m/d/yy"<br/>
     *       0xf, "d-mmm-yy"<br/>
     *       0x10, "d-mmm"<br/>
     *       0x11, "mmm-yy"<br/>
     *       0x12, "h:mm AM/PM"<br/>
     *       0x13, "h:mm:ss AM/PM"<br/>
     *       0x14, "h:mm"<br/>
     *       0x15, "h:mm:ss"<br/>
     *       0x16, "m/d/yy h:mm"<br/>
     *<p/>
     *       // 0x17 - 0x24 reserved for international and undocumented
     *       0x25, "#,##0_);(#,##0)"<br/>
     *       0x26, "#,##0_);[Red](#,##0)"<br/>
     *       0x27, "#,##0.00_);(#,##0.00)"<br/>
     *       0x28, "#,##0.00_);[Red](#,##0.00)"<br/>
     *       0x29, "_(*#,##0_);_(*(#,##0);_(* \"-\"_);_(@_)"<br/>
     *       0x2a, "_($*#,##0_);_($*(#,##0);_($* \"-\"_);_(@_)"<br/>
     *       0x2b, "_(*#,##0.00_);_(*(#,##0.00);_(*\"-\"??_);_(@_)"<br/>
     *       0x2c, "_($*#,##0.00_);_($*(#,##0.00);_($*\"-\"??_);_(@_)"<br/>
     *       0x2d, "mm:ss"<br/>
     *       0x2e, "[h]:mm:ss"<br/>
     *       0x2f, "mm:ss.0"<br/>
     *       0x30, "##0.0E+0"<br/>
     *       0x31, "@" - This is text format.<br/>
     *       0x31  "text" - Alias for "@"<br/>
     * <p/>
     *
     * @author Yegor Kozlov
     *
     * Modified 6/17/09 by Stanislav Shor - positive formats don't need starting '('
     *
     */
    public class BuiltinFormats
    {
        /**
         * The first user-defined number format starts at 164.
         */
        public const int FIRST_USER_DEFINED_FORMAT_INDEX = 164;

        private static readonly String[] _formats = new string[]{
            "General",
            "0",
            "0.00",
            "#,##0",
            "#,##0.00",
            "\"$\"#,##0_);(\"$\"#,##0)",
            "\"$\"#,##0_);[Red](\"$\"#,##0)",
            "\"$\"#,##0.00_);(\"$\"#,##0.00)",
            "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
            "0%",
            "0.00%",
            "0.00E+00",
            "# ?/?",
            "# ??/??",
            "m/d/yy",
            "d-mmm-yy",
            "d-mmm",
            "mmm-yy",
            "h:mm AM/PM",
            "h:mm:ss AM/PM",
            "h:mm",
            "h:mm:ss",
            "m/d/yy h:mm",

            // 0x17 - 0x24 reserved for international and undocumented
            // TODO - one junit relies on these values which seems incorrect
            "reserved-0x17",
            "reserved-0x18",
            "reserved-0x19",
            "reserved-0x1A",
            "reserved-0x1B",
            "reserved-0x1C",
            "reserved-0x1D",
            "reserved-0x1E",
            "reserved-0x1F",
            "reserved-0x20",
            "reserved-0x21",
            "reserved-0x22",
            "reserved-0x23",
            "reserved-0x24",

            "#,##0_);(#,##0)",
            "#,##0_);[Red](#,##0)",
            "#,##0.00_);(#,##0.00)",
            "#,##0.00_);[Red](#,##0.00)",
            "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
            "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
            "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
            "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
            "mm:ss",
            "[h]:mm:ss",
            "mm:ss.0",
            "##0.0E+0",
            "@"
        };

        public static String[] GetAll()
        {
            return (String[])_formats.Clone();
        }

        /**
         * Get the format string that matches the given format index
         *
         * @param index of a built in format
         * @return string represented at index of format or <code>null</code> if there is not a built-in format at that index
         */
        public static String GetBuiltinFormat(int index)
        {
            if (index < 0 || index >= _formats.Length)
            {
                return null;
            }
            return _formats[index];
        }

        /**
         * Get the format index that matches the given format string.
         * 
         * <p>
         * Automatically converts "text" to excel's format string to represent text.
         * </p>
         * @param pFmt string matching a built-in format
         * @return index of format or -1 if undefined.
         */
        public static int GetBuiltinFormat(String pFmt)
        {
            String fmt = "TEXT".Equals(pFmt, StringComparison.OrdinalIgnoreCase) ? "@" : pFmt;

            int i = -1;
            foreach (String f in _formats)
            {
                i++;
                if (f.Equals(fmt))
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>General format (index 0) - "General"</summary>
        public static string General => _formats[0];
        /// <summary>Integer with no decimal places (index 1) - "0"</summary>
        public static string Integer => _formats[1];
        /// <summary>Decimal with two decimal places (index 2) - "0.00"</summary>
        public static string Decimal2 => _formats[2];
        /// <summary>Integer with thousands separator (index 3) - "#,##0"</summary>
        public static string Thousands => _formats[3];
        /// <summary>Decimal with thousands separator (index 4) - "#,##0.00"</summary>
        public static string Currency => _formats[4];
        /// <summary>Currency with parentheses for negatives (index 5) - "$#,##0_);($#,##0)"</summary>
        public static string Accounting => _formats[5];
        /// <summary>Currency red (index 6) - "$#,##0_);[Red]($#,##0)"</summary>
        public static string AccountingRed => _formats[6];
        /// <summary>Currency decimal (index 7) - "$#,##0.00_);($#,##0.00)"</summary>
        public static string AccountingDecimal => _formats[7];
        /// <summary>Currency decimal red (index 8) - "$#,##0.00_);[Red]($#,##0.00)"</summary>
        public static string AccountingDecimalRed => _formats[8];
        /// <summary>Percentage no decimals (index 9) - "0%"</summary>
        public static string Percent => _formats[9];
        /// <summary>Percentage with two decimals (index 10) - "0.00%"</summary>
        public static string Percent2 => _formats[10];
        /// <summary>Scientific notation (index 11) - "0.00E+00"</summary>
        public static string Scientific => _formats[11];
        /// <summary>Fraction one digit (index 12) - "# ?/?"</summary>
        public static string Fraction1 => _formats[12];
        /// <summary>Fraction two digits (index 13) - "# ??/??"</summary>
        public static string Fraction2 => _formats[13];
        /// <summary>Short date (index 14) - "m/d/yy"</summary>
        public static string ShortDate => _formats[14];
        /// <summary>Long date (index 15) - "d-mmm-yy"</summary>
        public static string LongDate => _formats[15];
        /// <summary>Date (index 16) - "d-mmm"</summary>
        public static string Date2 => _formats[16];
        /// <summary>Month year (index 17) - "mmm-yy"</summary>
        public static string MonthYear => _formats[17];
        /// <summary>Time 12-hour (index 18) - "h:mm AM/PM"</summary>
        public static string Time12 => _formats[18];
        /// <summary>Time 12-hour with seconds (index 19) - "h:mm:ss AM/PM"</summary>
        public static string Time12Sec => _formats[19];
        /// <summary>Time 24-hour (index 20) - "h:mm"</summary>
        public static string Time24 => _formats[20];
        /// <summary>Time 24-hour with seconds (index 21) - "h:mm:ss"</summary>
        public static string Time24Sec => _formats[21];
        /// <summary>Date and time (index 22) - "m/d/yy h:mm"</summary>
        public static string DateTime => _formats[22];
        /// <summary>Thousands with parentheses (index 37) - "#,##0_);(#,##0)"</summary>
        public static string ThousandsParens => _formats[37];
        /// <summary>Thousands red (index 38) - "#,##0_);[Red](#,##0)"</summary>
        public static string ThousandsParensRed => _formats[38];
        /// <summary>Decimal thousands parentheses (index 39) - "#,##0.00_);(#,##0.00)"</summary>
        public static string DecimalThousandsParens => _formats[39];
        /// <summary>Decimal thousands red (index 40) - "#,##0.00_);[Red](#,##0.00)"</summary>
        public static string DecimalThousandsParensRed => _formats[40];
        /// <summary>Minutes seconds (index 41) - "mm:ss"</summary>
        public static string MinutesSeconds => _formats[41];
        /// <summary>Hours (index 42) - "[h]:mm:ss"</summary>
        public static string Hours => _formats[42];
        /// <summary>Minutes seconds tenths (index 43) - "mm:ss.0"</summary>
        public static string MinutesSecondsTenths => _formats[43];
        /// <summary>Scientific (index 44) - "##0.0E+0"</summary>
        public static string Scientific1 => _formats[44];
        /// <summary>Text format (index 45) - "@"</summary>
        public static string Text => _formats[45];
    }
}
