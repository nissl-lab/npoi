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

    using NPOI.Util;
    using System;
    using System.Text;

    /**
     * Title:        DBCell Record
     * Description:  Used by Excel and other MS apps to quickly Find rows in the sheets.
     * REFERENCE:  PG 299/440 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height
     * @version 2.0-pre
     */
    public class DBCellRecord : StandardRecord
    {
        public const int BLOCK_SIZE = 32;
        public const short sid = 0xd7;


        public DBCellRecord()
        {
            field_2_cell_offsets = new short[0];
        }

        /**
         * Constructs a DBCellRecord and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */
        public DBCellRecord(RecordInputStream in1)
        {
            field_1_row_offset = in1.ReadUShort();
            int size = in1.Remaining;
            field_2_cell_offsets = new short[size / 2];

            for (int i = 0; i < field_2_cell_offsets.Length; i++)
            {
                field_2_cell_offsets[i] = in1.ReadShort();
            }
        }

        private int field_1_row_offset;
        private short[] field_2_cell_offsets;

           /**
         * offset from the start of this DBCellRecord to the start of the first cell in
         * the next DBCell block.
         */
        public DBCellRecord(int rowOffset, short[]cellOffsets) {
            field_1_row_offset = rowOffset;
            field_2_cell_offsets = cellOffsets;
        }

        // need short list impl.
        public void AddCellOffset(short offset)
        {
            if (field_2_cell_offsets == null)
            {
                field_2_cell_offsets = new short[1];
            }
            else
            {
                short[] temp = new short[field_2_cell_offsets.Length + 1];

                Array.Copy(field_2_cell_offsets, 0, temp, 0,
                                 field_2_cell_offsets.Length);
                field_2_cell_offsets = temp;
            }
            field_2_cell_offsets[field_2_cell_offsets.Length - 1] = offset;
        }

        /**
         * Gets offset from the start of this DBCellRecord to the start of the first cell in
         * the next DBCell block.
         *
         * @return rowoffset to the start of the first cell in the next DBCell block
         */
        public int RowOffset
        {
            get { return field_1_row_offset; }
            set { field_1_row_offset = value; }
        }

        /**
         * return the cell offset in the array
         *
         * @param index of the cell offset to retrieve
         * @return celloffset from the celloffset array
         */
        public short GetCellOffsetAt(int index)
        {
            return field_2_cell_offsets[index];
        }

        /**
         * Get the number of cell offsets in the celloffset array
         *
         * @return number of cell offsets
         */
        public int NumCellOffsets
        {
            get{return field_2_cell_offsets.Length;}
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DBCELL]\n");
            buffer.Append("    .rowoffset       = ")
                .Append(StringUtil.ToHexString(RowOffset)).Append("\n");
            for (int k = 0; k < field_2_cell_offsets.Length; k++)
            {
                buffer.Append("    .cell_").Append(k).Append(" = ")
                    .Append(HexDump.ShortToHex(field_2_cell_offsets[k])).Append("\n");
            }
            buffer.Append("[/DBCELL]\n");
            return buffer.ToString();
        }
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteInt(field_1_row_offset);
            for (int k = 0; k < field_2_cell_offsets.Length; k++)
            {
                out1.WriteShort(field_2_cell_offsets[k]);
            }
        }

        protected override int DataSize
        {
            get { return 4 + field_2_cell_offsets.Length * 2; }
        }

        /**
         *  @returns the size of the Group of <c>DBCellRecord</c>s needed to encode
         *  the specified number of blocks and rows
         */
        public static int CalculateSizeOfRecords(int nBlocks, int nRows)
        {
            // One DBCell per block.
            // 8 bytes per DBCell (non variable section)
            // 2 bytes per row reference
            return nBlocks * 8 + nRows * 2;
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            // TODO - make immutable.
            // this should be safe because only the instantiating code mutates these objects
            return this;
        }
    }
}