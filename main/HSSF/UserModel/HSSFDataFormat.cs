/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Collections;

    using NPOI.Util;
    using NPOI.DDF;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using System.Collections.Generic;
    using NPOI.SS.UserModel;

    /// <summary>
    /// Utility to identify builtin formats.  Now can handle user defined data formats also.  The following Is a list of the formats as
    /// returned by this class.
    /// </summary>
    /// <remark>
    ///  @author  Andrew C. Oliver (acoliver at apache dot org)
    ///  @author  Shawn M. Laubach (slaubach at apache dot org)
    ///  </remark>
    ///  <example>
    ///       0, "General"
    ///       1, "0"
    ///       2, "0.00"
    ///       3, "#,##0"
    ///       4, "#,##0.00"
    ///       5, "($#,##0_);($#,##0)"
    ///       6, "($#,##0_);[Red]($#,##0)"
    ///       7, "($#,##0.00);($#,##0.00)"
    ///       8, "($#,##0.00_);[Red]($#,##0.00)"
    ///       9, "0%"
    ///       0xa, "0.00%"
    ///       0xb, "0.00E+00"
    ///       0xc, "# ?/?"
    ///       0xd, "# ??/??"
    ///       0xe, "m/d/yy"
    ///       0xf, "d-mmm-yy"
    ///       0x10, "d-mmm"
    ///       0x11, "mmm-yy"
    ///       0x12, "h:mm AM/PM"
    ///       0x13, "h:mm:ss AM/PM"
    ///       0x14, "h:mm"
    ///       0x15, "h:mm:ss"
    ///       0x16, "m/d/yy h:mm"
    ///
    ///       // 0x17 - 0x24 reserved for international and Undocumented
    ///       0x25, "(#,##0_);(#,##0)"
    ///       0x26, "(#,##0_);[Red](#,##0)"
    ///       0x27, "(#,##0.00_);(#,##0.00)"
    ///       0x28, "(#,##0.00_);[Red](#,##0.00)"
    ///       0x29, "_(///#,##0_);_(///(#,##0);_(/// \"-\"_);_(@_)"
    ///       0x2a, "_($///#,##0_);_($///(#,##0);_($/// \"-\"_);_(@_)"
    ///       0x2b, "_(///#,##0.00_);_(///(#,##0.00);_(///\"-\"??_);_(@_)"
    ///       0x2c, "_($///#,##0.00_);_($///(#,##0.00);_($///\"-\"??_);_(@_)"
    ///       0x2d, "mm:ss"
    ///       0x2e, "[h]:mm:ss"
    ///       0x2f, "mm:ss.0"
    ///       0x30, "##0.0E+0"
    ///       0x31, "@" - This Is text format.
    ///       0x31  "text" - Alias for "@"
    ///  </example>
    [Serializable]
    public class HSSFDataFormat: IDataFormat
    {
        	/**
	 * The first user-defined format starts at 164.
	 */
	public const int FIRST_USER_DEFINED_FORMAT_INDEX = 164;

        private static List<string> builtinFormats = CreateBuiltinFormats();

        private List<string> formats = new List<string>();
        private InternalWorkbook workbook;
        private bool movedBuiltins = false;  // Flag to see if need to
        // Check the built in list
        // or if the regular list
        // has all entries.

        /// <summary>
        /// Construncts a new data formatter.  It takes a workbook to have
        /// access to the workbooks format records.
        /// </summary>
        /// <param name="workbook">the workbook the formats are tied to..</param>
        public HSSFDataFormat(InternalWorkbook workbook)
        {
            this.workbook = workbook;
            IEnumerator i = workbook.Formats.GetEnumerator();
            while (i.MoveNext())
            {
                FormatRecord r = (FormatRecord)i.Current;
                for (int j = formats.Count; formats.Count <= r.GetIndexCode(); j++)
                {
                    formats.Add(null);
                }
                formats[r.GetIndexCode()] = r.GetFormatString();
            }
        }

        //synchonized
        private static List<string> CreateBuiltinFormats()
        {
            List<string> builtinFormats = new List<string>();
            builtinFormats.Insert(0, "General");
            builtinFormats.Insert(1, "0");
            builtinFormats.Insert(2, "0.00");
            builtinFormats.Insert(3, "#,##0");
            builtinFormats.Insert(4, "#,##0.00");
            builtinFormats.Insert(5, "($#,##0_);($#,##0)");
            builtinFormats.Insert(6, "($#,##0_);[Red]($#,##0)");
            builtinFormats.Insert(7, "($#,##0.00);($#,##0.00)");
            builtinFormats.Insert(8, "($#,##0.00_);[Red]($#,##0.00)");
            builtinFormats.Insert(9, "0%");
            builtinFormats.Insert(0xa, "0.00%");
            builtinFormats.Insert(0xb, "0.00E+00");
            builtinFormats.Insert(0xc, "# ?/?");
            builtinFormats.Insert(0xd, "# ??/??");
            builtinFormats.Insert(0xe, "m/d/yy");
            builtinFormats.Insert(0xf, "d-mmm-yy");
            builtinFormats.Insert(0x10, "d-mmm");
            builtinFormats.Insert(0x11, "mmm-yy");
            builtinFormats.Insert(0x12, "h:mm AM/PM");
            builtinFormats.Insert(0x13, "h:mm:ss AM/PM");
            builtinFormats.Insert(0x14, "h:mm");
            builtinFormats.Insert(0x15, "h:mm:ss");
            builtinFormats.Insert(0x16, "m/d/yy h:mm");

            // 0x17 - 0x24 reserved for international and Undocumented
            builtinFormats.Insert(0x17, "0x17");
            builtinFormats.Insert(0x18, "0x18");
            builtinFormats.Insert(0x19, "0x19");
            builtinFormats.Insert(0x1a, "0x1a");
            builtinFormats.Insert(0x1b, "0x1b");
            builtinFormats.Insert(0x1c, "0x1c");
            builtinFormats.Insert(0x1d, "0x1d");
            builtinFormats.Insert(0x1e, "0x1e");
            builtinFormats.Insert(0x1f, "0x1f");
            builtinFormats.Insert(0x20, "0x20");
            builtinFormats.Insert(0x21, "0x21");
            builtinFormats.Insert(0x22, "0x22");
            builtinFormats.Insert(0x23, "0x23");
            builtinFormats.Insert(0x24, "0x24");

            // 0x17 - 0x24 reserved for international and Undocumented
            builtinFormats.Insert(0x25, "(#,##0_);(#,##0)");
            builtinFormats.Insert(0x26, "(#,##0_);[Red](#,##0)");
            builtinFormats.Insert(0x27, "(#,##0.00_);(#,##0.00)");
            builtinFormats.Insert(0x28, "(#,##0.00_);[Red](#,##0.00)");
            builtinFormats.Insert(0x29, "_(*#,##0_);_(*(#,##0);_(* \"-\"_);_(@_)");
            builtinFormats.Insert(0x2a, "_($*#,##0_);_($*(#,##0);_($* \"-\"_);_(@_)");
            builtinFormats.Insert(0x2b, "_(*#,##0.00_);_(*(#,##0.00);_(*\"-\"??_);_(@_)");
            builtinFormats.Insert(0x2c,
                    "_($*#,##0.00_);_($*(#,##0.00);_($*\"-\"??_);_(@_)");
            builtinFormats.Insert(0x2d, "mm:ss");
            builtinFormats.Insert(0x2e, "[h]:mm:ss");
            builtinFormats.Insert(0x2f, "mm:ss.0");
            builtinFormats.Insert(0x30, "##0.0E+0");
            builtinFormats.Insert(0x31, "@");
            return builtinFormats;
        }

        public static List<string> GetBuiltinFormats()
        {
            return builtinFormats;
        }

        /// <summary>
        /// Get the format index that matches the given format string
        /// Automatically Converts "text" to excel's format string to represent text.
        /// </summary>
        /// <param name="format">The format string matching a built in format.</param>
        /// <returns>index of format or -1 if Undefined.</returns>
        public static short GetBuiltinFormat(String format)
        {
            if (format.ToUpper().Equals("TEXT"))
                format = "@";

            short retval = -1;

            for (short k = 0; k <= 0x31; k++)
            {
                String nformat = (String)builtinFormats[k];

                if ((nformat != null) && nformat.Equals(format))
                {
                    retval = k;
                    break;
                }
            }
            return retval;
        }

        /// <summary>
        /// Get the format index that matches the given format
        /// string, creating a new format entry if required.
        /// Aliases text to the proper format as required.
        /// </summary>
        /// <param name="format">The format string matching a built in format.</param>
        /// <returns>index of format.</returns>
        public short GetFormat(String format)
        {
            IEnumerator i;
            int ind;

            if (format.ToUpper().Equals("TEXT"))
                format = "@";

            if (!movedBuiltins)
            {
                i = builtinFormats.GetEnumerator();
                ind = 0;
                while (i.MoveNext())
                {
                    ind ++;

                    for (int j = formats.Count; formats.Count < ind + 1; j++)
                    {
                        formats.Add(null);
                    }
                    formats[ind] = i.Current as string;
                }
                movedBuiltins = true;
            }
            i = formats.GetEnumerator();
            ind = 0;
            while (i.MoveNext())
            {
                if (format.Equals(i.Current))
                    return (short)ind;

                ind++;
            }

            ind = workbook.GetFormat(format, true);
            for (int j = formats.Count; formats.Count < ind + 1; j++)
            {
                formats.Add(null);
            }
            formats[ind] = format;

            return (short)ind;
        }

        /// <summary>
        /// Get the format string that matches the given format index
        /// </summary>
        /// <param name="index">The index of a format.</param>
        /// <returns>string represented at index of format or null if there Is not a  format at that index</returns>
        public String GetFormat(short index)
        {
            if (movedBuiltins)
                return (String)formats[index];
            else
                return (String)(builtinFormats.Count > index
                        && builtinFormats[index] != null
                        ? builtinFormats[index] : formats[index]);
        }

        /// <summary>
        /// Get the format string that matches the given format index
        /// </summary>
        /// <param name="index">The index of a built in format.</param>
        /// <returns>string represented at index of format or null if there Is not a builtin format at that index</returns>
        public static String GetBuiltinFormat(short index)
        {
            return (String)builtinFormats[index];
        }

        /// <summary>
        /// Get the number of builtin and reserved builtinFormats
        /// </summary>
        /// <returns>number of builtin and reserved builtinFormats</returns>
        public static int NumberOfBuiltinBuiltinFormats
        {
            get
            {
                return builtinFormats.Count;
            }
        }
    }
}