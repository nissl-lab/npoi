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
using NPOI.SS.UserModel;
using System;
using NPOI.XSSF.Model;
namespace NPOI.XSSF.UserModel
{

    /**
     * Handles data formats for XSSF.
     * Per Microsoft Excel 2007+ format limitations:
     * Workbooks support between 200 and 250 "number formats"
     * (POI calls them "data formats") So short or even byte
     * would be acceptable data types to use for referring to
     * data format indices.
     * https://support.office.com/en-us/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
     * 
     */
    public class XSSFDataFormat : IDataFormat
    {
        private StylesTable stylesSource;

        public XSSFDataFormat(StylesTable stylesSource)
        {
            this.stylesSource = stylesSource;
        }

        /**
         * Get the format index that matches the given format
         *  string, creating a new format entry if required.
         * Aliases text to the proper format as required.
         *
         * @param format string matching a built in format
         * @return index of format.
         */
        public short GetFormat(String format)
        {
            int idx = BuiltinFormats.GetBuiltinFormat(format);
            if (idx == -1) idx = stylesSource.PutNumberFormat(format);
            return (short)idx;
        }

        /**
         * Get the format string that matches the given format index
         * @param index of a format
         * @return string represented at index of format or null if there is not a  format at that index
         */
        public String GetFormat(short index)
        {
            // Indices used for built-in formats may be overridden with
            // custom formats, such as locale-specific currency.
            // See org.apache.poi.xssf.usermodel.TestXSSFDataFormat#test49928() 
            // or bug 49928 for an example.
            // This is why we need to check stylesSource first and only fall back to
            // BuiltinFormats if the format hasn't been overridden.
            String fmt = stylesSource.GetNumberFormatAt(index);
            if (fmt == null) fmt = BuiltinFormats.GetBuiltinFormat(index);
            return fmt;
        }

        /**
         * get the format string that matches the given format index
         * @param index of a format
         * @return string represented at index of format or <code>null</code> if there is not a  format at that index
         * 
         * @deprecated POI 3.16 beta 1 - use {@link #getFormat(short)} instead
         */

        [Obsolete("use GetFormat(short) instead, schedule to remove NPOI 2.8")]
        public String GetFormat(int index)
        {
            return GetFormat((short)index);
        }
        /**
         * Add a number format with a specific ID into the number format style table.
         * If a format with the same ID already exists, overwrite the format code
         * with <code>fmt</code>
         * This may be used to override built-in number formats.
         *
         * @param index the number format ID
         * @param format the number format code
         */
        public void PutFormat(short index, String format)
        {
            stylesSource.PutNumberFormat(index, format);
        }
    }
}


