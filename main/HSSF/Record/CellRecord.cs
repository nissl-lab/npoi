/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Text;
using NPOI.Util;


namespace NPOI.HSSF.Record
{
    public abstract class CellRecord : StandardRecord, CellValueRecordInterface,IComparable
    {
        private int _rowIndex;
        private int _columnIndex;
        private int _formatIndex;

        protected CellRecord()
        {
            // fields uninitialised
        }

        protected CellRecord(RecordInputStream in1)
        {
            _rowIndex = in1.ReadUShort();
            _columnIndex = in1.ReadUShort();
            _formatIndex = in1.ReadUShort();
        }

        public int Row
        {
            get
            {
                return _rowIndex;
            }
            set
            {
                _rowIndex = value;
            }
        }

        public int Column
        {
            get
            {
                return _columnIndex;
            }
            set
            {
                _columnIndex = value;
            }
        }


        /**
         * get the index to the ExtendedFormat
         *
         * @see org.apache.poi.hssf.record.ExtendedFormatRecord
         * @return index to the XF record
         */
        public short XFIndex
        {
            get
            {
                return (short)_formatIndex;
            }
            set
            {
                _formatIndex = value;
            }
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

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            String recordName = RecordName;

            sb.Append("[").Append(recordName).Append("]\n");
            sb.Append("    .row    = ").Append(HexDump.ShortToHex(Row)).Append("\n");
            sb.Append("    .col    = ").Append(HexDump.ShortToHex(Column)).Append("\n");
            sb.Append("    .xfindex= ").Append(HexDump.ShortToHex(XFIndex)).Append("\n");
            AppendValueText(sb);
            sb.Append("\n");
            sb.Append("[/").Append(recordName).Append("]\n");
            return sb.ToString();
        }

        /**
         * Append specific debug info (used by {@link #toString()} for the value
         * contained in this record. Trailing new-line should not be Appended
         * (superclass does that).
         */
        protected abstract void AppendValueText(StringBuilder sb);

        /**
         * Gets the debug info BIFF record type name (used by {@link #toString()}.
         */
        protected abstract String RecordName { get; }

        /**
         * writes out the value data for this cell record
         */
        protected abstract void SerializeValue(ILittleEndianOutput out1);

        /**
         * @return the size (in bytes) of the value data for this cell record
         */
        protected abstract int ValueDataSize { get; }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Row);
            out1.WriteShort(Column);
            out1.WriteShort(XFIndex);
            SerializeValue(out1);
        }

        protected override int DataSize
        {
            get
            {
                return 6 + ValueDataSize;
            }
        }

        protected void CopyBaseFields(CellRecord rec)
        {
            rec._rowIndex = _rowIndex;
            rec._columnIndex = _columnIndex;
            rec._formatIndex = _formatIndex;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is CellValueRecordInterface))
            {
                return false;
            }
            CellValueRecordInterface loc = (CellValueRecordInterface)obj;

            if ((this.Row == loc.Row)
                    && (this.Column == loc.Column))
            {
                return true;
            }
            return false;
        }

        public override int GetHashCode ()
        {
            return Row ^ Column;
        }
    }
}
