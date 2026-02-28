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

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * Base class for all old (Biff 2 - Biff 4) cell value records 
     *  (implementors of {@link CellValueRecordInterface}).
     * Subclasses are expected to manage the cell data values (of various types).
     */
    public abstract class OldCellRecord
    {
        private short sid;
        private bool isBiff2;
        private int field_1_row;
        private short field_2_column;
        private int field_3_cell_attrs; // Biff 2
        private short field_3_xf_index;   // Biff 3+

        protected OldCellRecord(RecordInputStream in1, bool isBiff2)
        {
            this.sid = in1.Sid;
            this.isBiff2 = isBiff2;
            field_1_row = in1.ReadUShort();
            field_2_column = in1.ReadShort();

            if (isBiff2)
            {
                field_3_cell_attrs = in1.ReadUShort() << 8;
                field_3_cell_attrs += in1.ReadUByte();
            }
            else
            {
                field_3_xf_index = in1.ReadShort();
            }
        }

        public int Row
        {
            get
            {
                return field_1_row;
            }
        }

        public short Column
        {
            get
            {
                return field_2_column;
            }
        }

        /**
         * Get the index to the ExtendedFormat, for non-Biff2
         *
         * @see NPOI.HSSF.Record.ExtendedFormatRecord
         * @return index to the XF record
         */
        public short XFIndex
        {
            get
            {
                return field_3_xf_index;
            }
        }
        public int CellAttrs
        {
            get
            {
                return field_3_cell_attrs;
            }
        }

        /**
         * Is this a Biff2 record, or newer?
         */
        public virtual bool IsBiff2
        {
            get
            {
                return isBiff2;
            }
        }
        public virtual short Sid
        {
            get
            {
                return sid;
            }
        }


        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            String recordName = this.RecordName;

            sb.Append("[").Append(recordName).Append("]\n");
            sb.Append("    .row    = ").Append(HexDump.ShortToHex(Row)).Append("\n");
            sb.Append("    .col    = ").Append(HexDump.ShortToHex(Column)).Append("\n");
            if (IsBiff2)
            {
                sb.Append("    .cellattrs = ").Append(HexDump.ShortToHex(CellAttrs)).Append("\n");
            }
            else
            {
                sb.Append("    .xFindex   = ").Append(HexDump.ShortToHex(XFIndex)).Append("\n");
            }
            AppendValueText(sb);
            sb.Append("\n");
            sb.Append("[/").Append(recordName).Append("]\n");
            return sb.ToString();
        }

        /**
         * Append specific debug info (used by {@link #ToString()} for the value
         * Contained in this record. Trailing new-line should not be Appended
         * (superclass does that).
         */
        protected abstract void AppendValueText(StringBuilder sb);

        /**
         * Gets the debug info BIFF record type name (used by {@link #ToString()}.
         */
        protected abstract String RecordName { get; }
    }

}