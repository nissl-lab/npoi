
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

    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * Title:        Selection Record
     * Description:  shows the user's selection on the sheet
     *               for Write Set num refs to 0
     *
     * TODO :  Fully implement reference subrecords.
     * REFERENCE:  PG 291 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @author Glen Stampoultzis (glens at apache.org)
     */

    public class SelectionRecord
       : StandardRecord
    {
        public const short sid = 0x1d;
        private byte field_1_pane;
        private int field_2_row_active_cell;
        private int field_3_col_active_cell;
        private int field_4_ref_active_cell;
        private CellRangeAddress8Bit[] field_6_refs;     

        public SelectionRecord(int activeCellRow, int activeCellCol)
        {
            field_1_pane = 3; // pane id 3 is always present.  see OOO sec 5.75 'PANE'
            field_2_row_active_cell = activeCellRow;
            field_3_col_active_cell = activeCellCol;
            field_4_ref_active_cell = 0;
            field_6_refs = new CellRangeAddress8Bit[] {
                new CellRangeAddress8Bit(activeCellRow, activeCellRow, activeCellCol, activeCellCol),
            };
        }

        /// <summary>
        /// Constructs a Selection record and Sets its fields appropriately.
        /// </summary>
        /// <param name="in1">the RecordInputstream to Read the record from</param>
        public SelectionRecord(RecordInputStream in1)
        {
            field_1_pane = (byte)in1.ReadByte();
            
            field_2_row_active_cell = in1.ReadUShort();
            field_3_col_active_cell = in1.ReadShort();
            field_4_ref_active_cell = in1.ReadShort();
            int field_5_num_refs = in1.ReadUShort();

            field_6_refs = new CellRangeAddress8Bit[field_5_num_refs];
            for (int i = 0; i < field_5_num_refs; i++)
            {
                field_6_refs[i] = new CellRangeAddress8Bit(in1);
            }
        }

        /// <summary>
        /// Gets or sets the pane this is for.
        /// </summary>
        /// <value>The pane.</value>
        public byte Pane
        {
            get
            {
                return field_1_pane;
            }
            set { field_1_pane = value; }
        }

        /// <summary>
        /// Gets or sets the active cell row.
        /// </summary>
        /// <value>row number of active cell</value>
        public int ActiveCellRow
        {
            get
            {
                return field_2_row_active_cell;
            }
            set { field_2_row_active_cell = value; }
        }

        /// <summary>
        /// Gets or sets the active cell's col
        /// </summary>
        /// <value>number of active cell</value>
        public int ActiveCellCol
        {
            get
            {
                return field_3_col_active_cell;
            }
            set { field_3_col_active_cell = value; }
        }

        /// <summary>
        /// Gets or sets the active cell's reference number
        /// </summary>
        /// <value>ref number of active cell</value>
        public int ActiveCellRef
        {
            get
            {
                return field_4_ref_active_cell;
            }
            set { field_4_ref_active_cell = value; }
        }

        public CellRangeAddress8Bit[] CellReferences
        {
            get { return field_6_refs; }
            set { field_6_refs = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SELECTION]\n");
            buffer.Append("    .pane            = ")
                .Append(StringUtil.ToHexString(Pane)).Append("\n");
            buffer.Append("    .activecellrow   = ")
                .Append(StringUtil.ToHexString(ActiveCellRow)).Append("\n");
            buffer.Append("    .activecellcol   = ")
                .Append(StringUtil.ToHexString(ActiveCellCol)).Append("\n");
            buffer.Append("    .activecellref   = ")
                .Append(StringUtil.ToHexString(ActiveCellRef)).Append("\n");
            buffer.Append("    .numrefs         = ")
                .Append(StringUtil.ToHexString(field_6_refs.Length)).Append("\n");
            buffer.Append("[/SELECTION]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(Pane);
            out1.WriteShort(ActiveCellRow);
            out1.WriteShort(ActiveCellCol);
            out1.WriteShort(ActiveCellRef);
            int nRefs = field_6_refs.Length;
            out1.WriteShort(nRefs);
            for (int i = 0; i < field_6_refs.Length; i++)
            {
                field_6_refs[i].Serialize(out1);
            }
        }
        protected override int DataSize
        {
            get
            {
                return 9 // 1 byte + 4 shorts 
                    + CellRangeAddress8Bit.GetEncodedSize(field_6_refs.Length);
            }
        }


        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            SelectionRecord rec = new SelectionRecord(field_2_row_active_cell, field_3_col_active_cell);
            rec.field_1_pane = field_1_pane;
            rec.field_4_ref_active_cell = field_4_ref_active_cell;
            rec.field_6_refs = field_6_refs;
            return rec;
        }
    }
}
