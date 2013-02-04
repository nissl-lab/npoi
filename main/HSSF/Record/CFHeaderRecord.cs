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
    using NPOI.HSSF.Record.CF;
    using NPOI.SS.Util;


    /**
     * Conditional Formatting Header record (CFHEADER)
     * 
     * @author Dmitriy Kumshayev
     */
    public class CFHeaderRecord : StandardRecord
    {
        public const short sid = 0x1B0;

        private int field_1_numcf;
        private int field_2_need_recalculation;
        private CellRangeAddress field_3_enclosing_cell_range;
        private CellRangeAddressList field_4_cell_ranges;

        /** Creates new CFHeaderRecord */
        public CFHeaderRecord()
        {
            field_4_cell_ranges = new CellRangeAddressList();
        }
        public CFHeaderRecord(CellRangeAddress[] regions, int nRules)
        {
            CellRangeAddress[] unmergedRanges = regions;
            CellRangeAddress[] mergeCellRanges = CellRangeUtil.MergeCellRanges(unmergedRanges);
            CellRanges= mergeCellRanges;
            field_1_numcf = nRules;
        }

        public CFHeaderRecord(RecordInputStream in1)
        {
            field_1_numcf = in1.ReadShort();
            field_2_need_recalculation = in1.ReadShort();
            field_3_enclosing_cell_range = new CellRangeAddress(in1);
            field_4_cell_ranges = new CellRangeAddressList(in1);

        }

        public int NumberOfConditionalFormats
        {
            get
            {
                return field_1_numcf;
            }
            set { field_1_numcf = value; }
        }

        public bool NeedRecalculation
        {
            get
            {
                return field_2_need_recalculation == 1 ? true : false;
            }
            set { field_2_need_recalculation = value? 1 : 0; }
        }

        public CellRangeAddress EnclosingCellRange
        {
            get
            {
                return field_3_enclosing_cell_range;
            }
            set { field_3_enclosing_cell_range = value; }
        }


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
                    throw new ArgumentException("cellRanges must not be null");
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

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CFHEADER]\n");
            buffer.Append("	.id		= ").Append(StringUtil.ToHexString(sid)).Append("\n");
            buffer.Append("	.numCF			= ").Append(NumberOfConditionalFormats).Append("\n");
            buffer.Append("	.needRecalc	   = ").Append(NeedRecalculation).Append("\n");
            buffer.Append("	.enclosingCellRange= ").Append(EnclosingCellRange).Append("\n");
            buffer.Append("	.cfranges=[");
            for (int i = 0; i < field_4_cell_ranges.CountRanges(); i++)
            {
                buffer.Append(i == 0 ? "" : ",").Append(field_4_cell_ranges.GetCellRangeAddress(i).ToString());
            }
            buffer.Append("]\n");
            buffer.Append("[/CFHEADER]\n");
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
            out1.WriteShort(field_2_need_recalculation);
            field_3_enclosing_cell_range.Serialize(out1);
            field_4_cell_ranges.Serialize(out1);
        }


        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            CFHeaderRecord result = new CFHeaderRecord();
            result.field_1_numcf = field_1_numcf;
            result.field_2_need_recalculation = field_2_need_recalculation;
            result.field_3_enclosing_cell_range = field_3_enclosing_cell_range;
            result.field_4_cell_ranges = field_4_cell_ranges.Copy();
            return result;
        }
    }
}