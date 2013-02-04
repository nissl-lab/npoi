
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
        

/*
 * BlankRecord.java
 *
 * Created on December 10, 2001, 12:07 PM
 */
namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * Title:        Blank cell record 
     * Description:  Represents a column in a row with no value but with styling.
     * REFERENCE:  PG 287 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class BlankRecord : StandardRecord, CellValueRecordInterface, IComparable
    {
        public const short sid = 0x201;
        //private short             field_1_row;
        private int field_1_row;
        private int field_2_col;
        private short field_3_xf;

        /** Creates a new instance of BlankRecord */

        public BlankRecord()
        {
        }

        /**
         * Constructs a BlankRecord and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */

        public BlankRecord(RecordInputStream in1)
        {
            field_1_row = in1.ReadUShort();
            field_2_col = in1.ReadShort();
            field_3_xf = in1.ReadShort();
        }

        /**
         * Get the row this cell occurs on
         *
         * @return the row
         */

        //public short Row
        public int Row
        {
            get { return field_1_row; }
            set { field_1_row = value; }
        }

        /**
         * Get the column this cell defines within the row
         *
         * @return the column
         */

        public int Column
        {
            get { return field_2_col; }
            set { field_2_col = value; }
        }

        /**
         * Set the index of the extended format record to style this cell with
         *
         * @param xf - the 0-based index of the extended format
         * @see org.apache.poi.hssf.record.ExtendedFormatRecord
         */

        public short XFIndex
        {
            set { field_3_xf = value; }
            get { return field_3_xf; }
        }

        /**
         * return the non static version of the id for this record.
         */

        public override short Sid
        {
            get { return sid; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[BLANK]\n");
            buffer.Append("row       = ").Append(HexDump.ShortToHex(Row)).Append("\n");
            buffer.Append("col       = ").Append(HexDump.ShortToHex(Column)).Append("\n");
            buffer.Append("xf        = ").Append(HexDump.ShortToHex(XFIndex)).Append("\n");
            buffer.Append("[/BLANK]\n");
            return buffer.ToString();
        }

        protected override int DataSize
        {
            get { return 6; }
        }

        /**
         * called by the class that is responsible for writing this sucker.
         * Subclasses should implement this so that their data is passed back in a
         * byte array.
         *
         * @return byte array containing instance data
         */

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Row);
            out1.WriteShort(Column);
            out1.WriteShort(XFIndex);
        }

        public int CompareTo(Object obj)
        {
            CellValueRecordInterface loc = (CellValueRecordInterface)obj;

            if ((this.Row == loc.Row)
                    && (this.Column == loc.Column))
            {
                return 0;
            }
            if (this.Row < loc.Row)
            {
                return -1;
            }
            if (this.Row > loc.Row)
            {
                return 1;
            }
            if (this.Column < loc.Column)
            {
                return -1;
            }
            if (this.Column > loc.Column)
            {
                return 1;
            }
            return -1;
        }

        public override Object Clone()
        {
            BlankRecord rec = new BlankRecord();
            rec.field_1_row = field_1_row;
            rec.field_2_col = field_2_col;
            rec.field_3_xf = field_3_xf;
            return rec;
        }
    }
}