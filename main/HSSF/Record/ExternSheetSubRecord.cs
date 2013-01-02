
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


namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using NPOI.Util;

    /**
     * Title:        A sub Record for Extern Sheet 
     * Description:  Defines a named range within a workbook. 
     * REFERENCE:  
     * @author Libin Roman (Vista Portal LDT. Developer)
     * @version 1.0-pre
     */

    public class ExternSheetSubRecord : Record
    {
        public const short sid = 0xFFF; // only here for conformance, doesn't really have an sid
        private short field_1_index_to_supbook;
        private short field_2_index_to_first_supbook_sheet;
        private short field_3_index_to_last_supbook_sheet;


        /** a Constractor for making new sub record
         */
        public ExternSheetSubRecord()
        {
        }

        /**
         * Constructs a Extern Sheet Sub Record record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */
        public ExternSheetSubRecord(RecordInputStream in1)
        {
            field_1_index_to_supbook = in1.ReadShort();
            field_2_index_to_first_supbook_sheet = in1.ReadShort();
            field_3_index_to_last_supbook_sheet = in1.ReadShort();
        }


        /** Sets the Index to the sup book
         * @param index sup book index
         */
        public void SetIndexToSupBook(short index)
        {
            field_1_index_to_supbook = index;
        }

        /** Gets the index to sup book
         * @return sup book index
         */
        public short GetIndexToSupBook()
        {
            return field_1_index_to_supbook;
        }

        /** Sets the index to first sheet in supbook
         * @param index index to first sheet
         */
        public void SetIndexToFirstSupBook(short index)
        {
            field_2_index_to_first_supbook_sheet = index;
        }

        /** Gets the index to first sheet from supbook
         * @return index to first supbook
         */
        public short GetIndexToFirstSupBook()
        {
            return field_2_index_to_first_supbook_sheet;
        }

        /** Sets the index to last sheet in supbook
         * @param index index to last sheet
         */
        public void SetIndexToLastSupBook(short index)
        {
            field_3_index_to_last_supbook_sheet = index;
        }

        /** Gets the index to last sheet in supbook
         * @return index to last supbook
         */
        public short GetIndexToLastSupBook()
        {
            return field_3_index_to_last_supbook_sheet;
        }



        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("   supbookindex =").Append(GetIndexToSupBook()).Append('\n');
            buffer.Append("   1stsbindex   =").Append(GetIndexToFirstSupBook()).Append('\n');
            buffer.Append("   lastsbindex  =").Append(GetIndexToLastSupBook()).Append('\n');
            return buffer.ToString();
        }

        /**
         * called by the class that Is responsible for writing this sucker.
         * Subclasses should implement this so that their data Is passed back in a
         * byte array.
         *
         * @param offset to begin writing at
         * @param data byte array containing instance data
         * @return number of bytes written
         */
        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, GetIndexToSupBook());
            LittleEndian.PutShort(data, 2 + offset, GetIndexToFirstSupBook());
            LittleEndian.PutShort(data, 4 + offset, GetIndexToLastSupBook());

            return RecordSize;
        }


        /** returns the record size
         */
        public override int RecordSize
        {
            get { return 6; }
        }

        /**
         * return the non static version of the id for this record.
         */
        public override short Sid
        {
            get { return sid; }
        }
    }

}