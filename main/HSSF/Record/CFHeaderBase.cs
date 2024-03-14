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
    using NPOI.HSSF.Record.CF;
    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * Parent of Conditional Formatting Header records,
     *  {@link CFHeaderRecord} and {@link CFHeader12Record}.
     */
    public abstract class CFHeaderBase : StandardRecord, ICloneable
    {
        private int field_1_numcf;
        private int field_2_need_recalculation_and_id;
        private CellRangeAddress field_3_enclosing_cell_range;
        private CellRangeAddressList field_4_cell_ranges;

        /** Creates new CFHeaderBase */
        protected CFHeaderBase()
        {
        }
        protected CFHeaderBase(CellRangeAddress[] regions, int nRules)
        {
            CellRangeAddress[] unmergedRanges = regions;
            CellRangeAddress[] mergeCellRanges = CellRangeUtil.MergeCellRanges(unmergedRanges);
            CellRanges = (mergeCellRanges);
            field_1_numcf = nRules;
        }

        protected void CreateEmpty()
        {
            field_3_enclosing_cell_range = new CellRangeAddress(0, 0, 0, 0);
            field_4_cell_ranges = new CellRangeAddressList();
        }
        protected void Read(RecordInputStream in1)
        {
            field_1_numcf = in1.ReadShort();
            field_2_need_recalculation_and_id = in1.ReadShort();
            field_3_enclosing_cell_range = new CellRangeAddress(in1);
            field_4_cell_ranges = new CellRangeAddressList(in1);
        }

        public int NumberOfConditionalFormats
        {
            get
            {
                return field_1_numcf;
            }
            set
            {
                field_1_numcf = value;
            }
        }

        public bool NeedRecalculation
        {
            get
            {
                // Held on the 1st bit
                return (field_2_need_recalculation_and_id & 1) == 1;
            }
            set
            {
                // held on the first bit
                if (value == NeedRecalculation)
                    return;
                if (value)
                    field_2_need_recalculation_and_id++;
                else
                    field_2_need_recalculation_and_id--;
            }
        }

        public int ID
        {
            get
            {
                // Remaining 15 bits of field 2
                return field_2_need_recalculation_and_id >> 1;
            }
            set
            {
                // Remaining 15 bits of field 2
                bool needsRecalc = NeedRecalculation;
                field_2_need_recalculation_and_id = (value << 1);
                if (needsRecalc)
                    field_2_need_recalculation_and_id++;
            }
        }

        public CellRangeAddress EnclosingCellRange
        {
            get
            {
                return field_3_enclosing_cell_range;
            }
            set
            {
                field_3_enclosing_cell_range = value;
            }
            
        }

        /**
         * Set cell ranges list to a single cell range and 
         * modify the enclosing cell range accordingly.
         * @param cellRanges - list of CellRange objects
         */
        public CellRangeAddress[] CellRanges
        {
            get
            {
                return field_4_cell_ranges.CellRangeAddresses;
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("cellRanges must not be null");
                }
                CellRangeAddressList cral = new CellRangeAddressList();
                CellRangeAddress enclosingRange = null;
                for (int i = 0; i < value.Length; i++)
                {
                    CellRangeAddress cr = value[i];
                    enclosingRange = CellRangeUtil.CreateEnclosingCellRange(cr, enclosingRange);
                    cral.AddCellRangeAddress(cr);
                }
                field_3_enclosing_cell_range = enclosingRange;
                field_4_cell_ranges = cral;
            }
        }

        protected abstract String RecordName { get; }
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[").Append(RecordName).Append("]\n");
            buffer.Append("\t.numCF             = ").Append(NumberOfConditionalFormats).Append("\n");
            buffer.Append("\t.needRecalc        = ").Append(NeedRecalculation).Append("\n");
            buffer.Append("\t.id                = ").Append(ID).Append("\n");
            buffer.Append("\t.enclosingCellRange= ").Append(EnclosingCellRange).Append("\n");
            buffer.Append("\t.CFranges=[");
            for (int i = 0; i < field_4_cell_ranges.CountRanges(); i++)
            {
                buffer.Append(i == 0 ? "" : ",").Append(field_4_cell_ranges.GetCellRangeAddress(i).ToString());
            }
            buffer.Append("]\n");
            buffer.Append("[/").Append(RecordName).Append("]\n");
            return buffer.ToString();
        }

        protected override int DataSize
        {
            get
            {
                return 4 // 2 short fields
                 + CellRangeAddress.ENCODED_SIZE
                 + field_4_cell_ranges.Size;
            }
            
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_numcf);
            out1.WriteShort(field_2_need_recalculation_and_id);
            field_3_enclosing_cell_range.Serialize(out1);
            field_4_cell_ranges.Serialize(out1);
        }

        protected void CopyTo(CFHeaderBase result)
        {
            result.field_1_numcf = field_1_numcf;
            result.field_2_need_recalculation_and_id = field_2_need_recalculation_and_id;
            result.field_3_enclosing_cell_range = field_3_enclosing_cell_range.Copy();
            result.field_4_cell_ranges = field_4_cell_ranges.Copy();
        }
    }

}