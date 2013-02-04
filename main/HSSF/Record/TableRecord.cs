/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License Is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;
    using NPOI.SS.Util;
    /**
     * DATATABLE (0x0236)<p/>
     *
     * TableRecord - The record specifies a data table.
     * This record Is preceded by a single Formula record that
     *  defines the first cell in the data table, which should
     *  only contain a single Ptg, {@link TblPtg}.
     *
     * See p536 of the June 08 binary docs
     */
    public class TableRecord : SharedValueRecordBase
    {
        public const short sid = 0x0236;

        private static BitField alwaysCalc = BitFieldFactory.GetInstance(0x0001);
        private static BitField calcOnOpen = BitFieldFactory.GetInstance(0x0002);
        private static BitField rowOrColInpCell = BitFieldFactory.GetInstance(0x0004);
        private static BitField oneOrTwoVar = BitFieldFactory.GetInstance(0x0008);
        private static BitField rowDeleted = BitFieldFactory.GetInstance(0x0010);
        private static BitField colDeleted = BitFieldFactory.GetInstance(0x0020);

        private int field_5_flags;
        private int field_6_res;
        private int field_7_rowInputRow;
        private int field_8_colInputRow;
        private int field_9_rowInputCol;
        private int field_10_colInputCol;

        public TableRecord(RecordInputStream in1)
            : base(in1)
        {

            field_5_flags = in1.ReadByte();
            field_6_res = in1.ReadByte();
            field_7_rowInputRow = in1.ReadShort();
            field_8_colInputRow = in1.ReadShort();
            field_9_rowInputCol = in1.ReadShort();
            field_10_colInputCol = in1.ReadShort();
        }

        public TableRecord(CellRangeAddress8Bit range)
            : base(range)
        {

            field_6_res = 0;
        }

        public int Flags
        {
            get
            {
                return field_5_flags;
            }
            set { field_5_flags = value; }
        }
        public int RowInputRow
        {
            get
            {
                return field_7_rowInputRow;
            }
            set { field_7_rowInputRow = value; }
        }

        public int ColInputRow
        {
            get
            {
                return field_8_colInputRow;
            }
            set { field_8_colInputRow = value; }
        }


        public int RowInputCol
        {
            get
            {
                return field_9_rowInputCol;
            }
            set { field_9_rowInputCol = value; }
        }
        public int ColInputCol
        {
            get
            {
                return field_10_colInputCol;
            }
            set { field_10_colInputCol = value; }
        }


        public bool IsAlwaysCalc
        {
            get
            {
                return alwaysCalc.IsSet(field_5_flags);
            }
            set { field_5_flags = alwaysCalc.SetBoolean(field_5_flags, value); }
        }


        public bool IsRowOrColInpCell
        {
            get
            {
                return rowOrColInpCell.IsSet(field_5_flags);
            }
            set { field_5_flags = rowOrColInpCell.SetBoolean(field_5_flags, value); }
        }

        public bool IsOneNotTwoVar
        {
            get
            {
                return oneOrTwoVar.IsSet(field_5_flags);
            }
            set { field_5_flags = oneOrTwoVar.SetBoolean(field_5_flags, value); }
        }

        public bool IsColDeleted
        {
            get
            {
                return colDeleted.IsSet(field_5_flags);
            }
            set { field_5_flags = colDeleted.SetBoolean(field_5_flags, value); }
        }

        public bool IsRowDeleted
        {
            get
            {
                return rowDeleted.IsSet(field_5_flags);
            }
            set { field_5_flags = rowDeleted.SetBoolean(field_5_flags, value); }
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }
        protected override int ExtraDataSize
        {
            get
            {
                return
                2 // 2 byte fields
                + 8; // 4 short fields
            }
        }
        protected override void SerializeExtraData(ILittleEndianOutput out1)
        {
            out1.WriteByte(field_5_flags);
            out1.WriteByte(field_6_res);
            out1.WriteShort(field_7_rowInputRow);
            out1.WriteShort(field_8_colInputRow);
            out1.WriteShort(field_9_rowInputCol);
            out1.WriteShort(field_10_colInputCol);
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[TABLE]\n");
            buffer.Append("    .range    = ").Append(Range.ToString()).Append("\n");
            buffer.Append("    .flags    = ").Append(HexDump.ByteToHex(field_5_flags)).Append("\n");
            buffer.Append("    .alwaysClc= ").Append(IsAlwaysCalc).Append("\n");
            buffer.Append("    .reserved = ").Append(HexDump.IntToHex(field_6_res)).Append("\n");
            CellReference crRowInput = cr(field_7_rowInputRow, field_8_colInputRow);
            CellReference crColInput = cr(field_9_rowInputCol, field_10_colInputCol);
            buffer.Append("    .rowInput = ").Append(crRowInput.FormatAsString()).Append("\n");
            buffer.Append("    .colInput = ").Append(crColInput.FormatAsString()).Append("\n");
            buffer.Append("[/TABLE]\n");
            return buffer.ToString();
        }

        private static CellReference cr(int rowIx, int colIxAndFlags)
        {
            int colIx = colIxAndFlags & 0x00FF;
            bool IsRowAbs = (colIxAndFlags & 0x8000) == 0;
            bool IsColAbs = (colIxAndFlags & 0x4000) == 0;
            return new CellReference(rowIx, colIx, IsRowAbs, IsColAbs);
        }
    }
}