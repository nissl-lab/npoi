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

    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using System.Collections.Generic;
    using NPOI.SS.UserModel;

    /*
     * Identifies both built-in and user defined formats within a workbook.<p/>
     * See {@link BuiltinFormats} for a list of supported built-in formats.<p/>
     *
     * <b>International Formats</b><br/>
     * Since version 2003 Excel has supported international formats.  These are denoted
     * with a prefix "[$-xxx]" (where xxx is a 1-7 digit hexadecimal number).
     * See the Microsoft article
     * <a href="http://office.microsoft.com/assistance/hfws.aspx?AssetID=HA010346351033&CTT=6&Origin=EC010272491033">
     *   Creating international number formats
     * </a> for more details on these codes.
     *
     * @author  Andrew C. Oliver (acoliver at apache dot org)
     * @author  Shawn M. Laubach (slaubach at apache dot org)
     */
    [Serializable]
    public class HSSFDataFormat : IDataFormat
    {
        /**
 * The first user-defined format starts at 164.
 */
        public const int FIRST_USER_DEFINED_FORMAT_INDEX = 164;

        private static List<string> builtinFormats = new List<string>(BuiltinFormats.GetAll());

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
        /// <param name="workbook">the workbook the formats are tied to.</param>
        public HSSFDataFormat(InternalWorkbook workbook)
        {
            this.workbook = workbook;
            IEnumerator i = workbook.Formats.GetEnumerator();
            while (i.MoveNext())
            {
                FormatRecord r = (FormatRecord)i.Current;
                for (int j = formats.Count; formats.Count <= r.IndexCode; j++)
                {
                    formats.Add(null);
                }
                formats[r.IndexCode] = r.FormatString;
            }
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
        /// <param name="pFormat">The format string matching a built in format.</param>
        /// <returns>index of format.</returns>
        public short GetFormat(String pFormat)
        {
            IEnumerator i;
            int ind;
            String format;
            if (pFormat.ToUpper().Equals("TEXT"))
            {
                format = "@";
            }
            else
            {
                format = pFormat;
            }

            if (!movedBuiltins)
            {
                i = builtinFormats.GetEnumerator();
                ind = 0;
                while (i.MoveNext())
                {
                    for (int j = formats.Count; formats.Count < ind + 1; j++)
                    {
                        formats.Add(null);
                    }
                    formats[ind] = i.Current as string;
                    ind++;
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

            if (index == -1)
                return null;

            String fmt = formats.Count > index ? formats[index] : null;

            if (builtinFormats.Count > index
                        && builtinFormats[index] != null)
            {
                // It's in the built in range
                if (fmt != null)
                {
                    // It's been overriden, use that value
                    return fmt;
                }
                else
                {
                    // Standard built in format
                    return builtinFormats[index];
                }
            }
            return fmt;
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
        /**
	     * Ensures that the formats list can hold entries
	     *  up to and including the entry with this index
	     */
        private void EnsureFormatsSize(int index)
        {
            if (formats.Count <= index)
            {
                formats.Capacity = (index + 1);
            }
        }
    }
}