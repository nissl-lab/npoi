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

namespace NPOI.HSSF.EventUserModel
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using System.Collections.Generic;
    using NPOI.SS.UserModel;
    using System.Globalization;
    using NPOI.Util;

    /**
     * A proxy HSSFListener that keeps track of the document
     *  formatting records, and provides an easy way to look
     *  up the format strings used by cells from their ids.
     */
    public class FormatTrackingHSSFListener : IHSSFListener
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(FormatTrackingHSSFListener));
        private IHSSFListener childListener;
        private Dictionary<int, FormatRecord> customFormatRecords = new Dictionary<int, FormatRecord>();
        private DataFormatter formatter = new DataFormatter();
        private List<ExtendedFormatRecord> xfRecords = new List<ExtendedFormatRecord>();

        public FormatTrackingHSSFListener(IHSSFListener childListener)
        {
            this.childListener = childListener;
        }

        public int NumberOfCustomFormats
        {
            get
            {
                return customFormatRecords.Count;
            }
        }
        public int NumberOfExtendedFormats
        {
            get
            {
                return xfRecords.Count;
            }
        }
        /**
         * Process this record ourselves, and then
         *  pass it on to our child listener
         */
        public void ProcessRecord(Record record)
        {
            // Handle it ourselves
            ProcessRecordInternally(record);

            // Now pass on to our child
            childListener.ProcessRecord(record);
        }

        /**
         * Process the record ourselves, but do not
         *  pass it on to the child Listener.
         * @param record
         */
        public void ProcessRecordInternally(Record record)
        {
            if (record is FormatRecord)
            {
                FormatRecord fr = (FormatRecord)record;
                customFormatRecords[fr.IndexCode] = fr;
            }
            else if (record is ExtendedFormatRecord)
            {
                ExtendedFormatRecord xr = (ExtendedFormatRecord)record;
                xfRecords.Add(xr);
            }
        }
        /**
         * Formats the given numeric of date Cell's contents
         *  as a String, in as close as we can to the way 
         *  that Excel would do so.
         * Uses the various format records to manage this.
         * 
         * TODO - move this to a central class in such a
         *  way that hssf.usermodel can make use of it too
         */
        public String FormatNumberDateCell(CellValueRecordInterface cell)
        {
            double value;
            if (cell is NumberRecord)
            {
                value = ((NumberRecord)cell).Value;
            }
            else if (cell is FormulaRecord)
            {
                value = ((FormulaRecord)cell).Value;
            }
            else
            {
                throw new ArgumentException("Unsupported CellValue Record passed in " + cell);
            }

            // Get the built in format, if there is one
            int formatIndex = GetFormatIndex(cell);
            String formatString = GetFormatString(cell);

            if (formatString == null)
            {
                return value.ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                // Format, using the nice new
                //  HSSFDataFormatter to do the work for us 
                return formatter.FormatRawCellContents(value, formatIndex, formatString);
            }
        }
        /**
         * Returns the format string, eg $##.##, for the
         *  given number format index.
         */
        public String GetFormatString(int formatIndex)
        {
            String format = null;
            if (formatIndex >= HSSFDataFormat.NumberOfBuiltinBuiltinFormats)
            {
                FormatRecord tfr = (FormatRecord)customFormatRecords[formatIndex];
                if (tfr == null)
                {
                    logger.Log(POILogger.ERROR, "Requested format at index " + formatIndex + ", but it wasn't found");
                }
                else
                {
                    format = tfr.FormatString;
                }
            }
            else
            {
                format = HSSFDataFormat.GetBuiltinFormat((short)formatIndex);
            }
            return format;
        }

        /**
         * Returns the format string, eg $##.##, used
         *  by your cell 
         */
        public String GetFormatString(CellValueRecordInterface cell)
        {
            int formatIndex = GetFormatIndex(cell);
            if (formatIndex == -1)
            {
                // Not found
                return null;
            }
            return GetFormatString(formatIndex);
        }

        /**
         * Returns the index of the format string, used by your cell,
         *  or -1 if none found
         */
        public int GetFormatIndex(CellValueRecordInterface cell)
        {
            ExtendedFormatRecord xfr = (ExtendedFormatRecord)
                xfRecords[cell.XFIndex];
            if (xfr == null)
            {
                logger.Log(POILogger.ERROR, "Cell " + cell.Row + "," + cell.Column + " uses XF with index " + cell.XFIndex + ", but we don't have that");
                return -1;
            }
            return xfr.FormatIndex;
        }
    }
}