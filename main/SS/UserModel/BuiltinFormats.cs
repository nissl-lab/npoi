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
         * The first user-defined format starts at 164.
         */
        public const int FIRST_USER_DEFINED_FORMAT_INDEX = 164;

        private static String[] _formats;

        /*
        0 General General 18 Time h:mm AM/PM
        1 Decimal 0 19 Time h:mm:ss AM/PM
        2 Decimal 0.00 20 Time h:mm
        3 Decimal #,##0 21 Time h:mm:ss
        4 Decimal #,##0.00 2232 Date/Time M/D/YY h:mm
        531 Currency "$"#,##0_);("$"#,##0) 37 Account. _(#,##0_);(#,##0)
        631 Currency "$"#,##0_);[Red]("$"#,##0) 38 Account. _(#,##0_);[Red](#,##0)
        731 Currency "$"#,##0.00_);("$"#,##0.00) 39 Account. _(#,##0.00_);(#,##0.00)
        831 Currency "$"#,##0.00_);[Red]("$"#,##0.00) 40 Account. _(#,##0.00_);[Red](#,##0.00)
        9 Percent 0% 4131 Currency _("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)
        10 Percent 0.00% 4231 33 Currency _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
        11 Scientific 0.00E+00 4331 Currency _("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)
        12 Fraction # ?/? 4431 33 Currency _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
        13 Fraction # ??/?? 45 Time mm:ss
        1432 Date M/D/YY 46 Time [h]:mm:ss
        15 Date D-MMM-YY 47 Time mm:ss.0
        16 Date D-MMM 48 Scientific ##0.0E+0
        17 Date MMM-YY 49 Text @
        * */
        static BuiltinFormats()
        {
            List<String> m = new List<String>();
            PutFormat(m, 0, "General");
            PutFormat(m, 1, "0");
            PutFormat(m, 2, "0.00");
            PutFormat(m, 3, "#,##0");
            PutFormat(m, 4, "#,##0.00");
            PutFormat(m, 5, "\"$\"#,##0_);(\"$\"#,##0)");
            PutFormat(m, 6, "\"$\"#,##0_);[Red](\"$\"#,##0)");
            PutFormat(m, 7, "\"$\"#,##0.00_);(\"$\"#,##0.00)");
            PutFormat(m, 8, "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)");
            PutFormat(m, 9, "0%");
            PutFormat(m, 0xa, "0.00%");
            PutFormat(m, 0xb, "0.00E+00");
            PutFormat(m, 0xc, "# ?/?");
            PutFormat(m, 0xd, "# ??/??");
            PutFormat(m, 0xe, "m/d/yy");
            PutFormat(m, 0xf, "d-mmm-yy");
            PutFormat(m, 0x10, "d-mmm");
            PutFormat(m, 0x11, "mmm-yy");
            PutFormat(m, 0x12, "h:mm AM/PM");
            PutFormat(m, 0x13, "h:mm:ss AM/PM");
            PutFormat(m, 0x14, "h:mm");
            PutFormat(m, 0x15, "h:mm:ss");
            PutFormat(m, 0x16, "m/d/yy h:mm");

            // 0x17 - 0x24 reserved for international and undocumented
            for (int i = 0x17; i <= 0x24; i++)
            {
                // TODO - one junit relies on these values which seems incorrect
                PutFormat(m, i, "reserved-0x" + (i).ToString("X", CultureInfo.CurrentCulture));
            }

            PutFormat(m, 0x25, "#,##0_);(#,##0)");
            PutFormat(m, 0x26, "#,##0_);[Red](#,##0)");
            PutFormat(m, 0x27, "#,##0.00_);(#,##0.00)");
            PutFormat(m, 0x28, "#,##0.00_);[Red](#,##0.00)");
            PutFormat(m, 0x29, "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)");
            PutFormat(m, 0x2a, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)");
            PutFormat(m, 0x2b, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)");
            PutFormat(m, 0x2c, "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)");
            PutFormat(m, 0x2d, "mm:ss");
            PutFormat(m, 0x2e, "[h]:mm:ss");
            PutFormat(m, 0x2f, "mm:ss.0");
            PutFormat(m, 0x30, "##0.0E+0");
            PutFormat(m, 0x31, "@");
            //String[] ss = new String[m.Count];
            String[] ss = m.ToArray();
            _formats = ss;
        }
        private static void PutFormat(List<String> m, int index, String value)
        {
            if (m.Count != index)
            {
                throw new InvalidOperationException("index " + index + " is wrong");
            }
            m.Add(value);
        }


        /**
         * @deprecated (May 2009) use {@link #getAll()}
         */
        [Obsolete]
        public static Dictionary<int, String> GetBuiltinFormats()
        {
            Dictionary<int, String> result = new Dictionary<int, String>();
            for (int i = 0; i < _formats.Length; i++)
            {
                result.Add(i, _formats[i]);
            }
            return result;
        }

        /**
         * @return array of built-in data formats
         */
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
            String fmt;
            if (string.Compare(pFmt, ("TEXT"), StringComparison.OrdinalIgnoreCase) == 0)
            {
                fmt = "@";
            }
            else
            {
                fmt = pFmt;
            }

            for (int i = 0; i < _formats.Length; i++)
            {
                if (fmt.Equals(_formats[i]))
                {
                    return i;
                }
            }
            return -1;
        }
    }

}