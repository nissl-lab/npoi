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
    using NPOI.HSSF.Record.Formula;
    using System.Collections;
    using NPOI.Util;

    /**
     * Title:        SharedFormulaRecord
     * Description:  Primarily used as an excel optimization so that multiple similar formulas
     * 				  are not written out too many times.  We should recognize this record and
     *               Serialize as Is since this Is used when Reading templates.
     * 
     * Note: the documentation says that the SID Is BC where biffviewer reports 4BC.  The hex dump shows
     * that the two byte sid representation to be 'BC 04' that Is consistent with the other high byte
     * record types.
     * @author Danny Mui at apache dot org
     */
    public class SharedFormulaRecord : SharedValueRecordBase
    {
        public const short sid = 0x4BC;


        private int field_5_reserved;
        private NPOI.SS.Formula.Formula field_7_parsed_expr;

        public SharedFormulaRecord()
        {
            field_7_parsed_expr = NPOI.SS.Formula.Formula.Create(Ptg.EMPTY_PTG_ARRAY);
        }

        /**
         * @param in the RecordInputstream to Read the record from
         */

        public SharedFormulaRecord(RecordInputStream in1):base(in1)
        {
            field_5_reserved        = in1.ReadShort();
            int field_6_expression_len = in1.ReadShort();
            int nAvailableBytes = in1.Available();
            field_7_parsed_expr = NPOI.SS.Formula.Formula.Read(field_6_expression_len, in1, nAvailableBytes);
        }
        protected override int ExtraDataSize
        {
            get
            {
                //Because this record is converted to individual Formula records, this method is not required.
                return 2 + field_7_parsed_expr.EncodedSize;
            }

        }

        /**
         * print a sort of string representation ([SHARED FORMULA RECORD] id = x [/SHARED FORMULA RECORD])
         */

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[SHARED FORMULA (").Append(HexDump.IntToHex(sid)).Append("]\n");
            buffer.Append("    .range      = ").Append(Range.ToString()).Append("\n");
            buffer.Append("    .reserved    = ").Append(HexDump.ShortToHex(field_5_reserved)).Append("\n");

            Ptg[] ptgs = field_7_parsed_expr.Tokens;
            for (int k = 0; k < ptgs.Length; k++)
            {
                buffer.Append("Formula[").Append(k).Append("]");
                Ptg ptg = ptgs[k];
                buffer.Append(ptg.ToString()).Append(ptg.RVAType).Append("\n");
            }

            buffer.Append("[/SHARED FORMULA]\n");
            return buffer.ToString();
        }

        public override short Sid
        {
            get { return sid; }
        }
        public override Object Clone()
        {
            //Because this record is converted to individual Formula records, this method is not required.
            throw new NotSupportedException("Cannot clone a SharedFormulaRecord");
        }
        /**
         * Creates a non shared formula from the shared formula counterpart<br/>
         *
         * Perhaps this functionality could be implemented in terms of the raw
         * byte array inside {@link Formula}.
         */
        public static Ptg[] ConvertSharedFormulas(Ptg[] ptgs, int formulaRow, int formulaColumn)
        {

            Ptg[] newPtgStack = new Ptg[ptgs.Length];

            for (int k = 0; k < ptgs.Length; k++)
            {
                Ptg ptg = ptgs[k];
                byte originalOperandClass = unchecked((byte)-1);
                if (!ptg.IsBaseToken)
                {
                    originalOperandClass = ptg.PtgClass;
                }
                if (ptg is RefPtgBase)
                {
                    RefPtgBase refNPtg = (RefPtgBase)ptg;
                    ptg = new RefPtg(FixupRelativeRow(formulaRow, refNPtg.Row, refNPtg.IsRowRelative),
                                         FixupRelativeColumn(formulaColumn, refNPtg.Column, refNPtg.IsColRelative),
                                         refNPtg.IsRowRelative,
                                         refNPtg.IsColRelative);
                    ptg.PtgClass = originalOperandClass;
                }
                else if (ptg is AreaPtgBase)
                {
                    AreaPtgBase areaNPtg = (AreaPtgBase)ptg;
                    ptg = new AreaPtg(FixupRelativeRow(formulaRow, areaNPtg.FirstRow, areaNPtg.IsFirstRowRelative),
                                    FixupRelativeRow(formulaRow, areaNPtg.LastRow, areaNPtg.IsLastRowRelative),
                                    FixupRelativeColumn(formulaColumn, areaNPtg.FirstColumn, areaNPtg.IsFirstColRelative),
                                    FixupRelativeColumn(formulaColumn, areaNPtg.LastColumn, areaNPtg.IsLastColRelative),
                                    areaNPtg.IsFirstRowRelative,
                                    areaNPtg.IsLastRowRelative,
                                    areaNPtg.IsFirstColRelative,
                                    areaNPtg.IsLastColRelative);
                    ptg.PtgClass = originalOperandClass;
                }
                else if (ptg is OperandPtg)
                {
                    // Any subclass of OperandPtg is mutable, so it's safest to not share these instances.
                    ptg = ((OperandPtg)ptg).Copy();
                }
                else
                {
                    // all other Ptgs are immutable and can be shared
                }
                newPtgStack[k] = ptg;
            }
            return newPtgStack;
        }
        protected override void SerializeExtraData(NPOI.Util.IO.LittleEndianOutput out1)
        {
            out1.WriteShort(field_5_reserved);
            field_7_parsed_expr.Serialize(out1);
        }
        /**
 * @return the equivalent {@link Ptg} array that the formula would have, were it not shared.
 */
        public Ptg[] GetFormulaTokens(FormulaRecord formula)
        {
            int formulaRow = formula.Row;
            int formulaColumn = formula.Column;
            //Sanity checks
            if (!IsInRange(formulaRow, formulaColumn))
            {
                throw new Exception("Shared Formula Conversion: Coding Error");
            }

            return ConvertSharedFormulas(field_7_parsed_expr.Tokens, formulaRow, formulaColumn);
        }

        private static int FixupRelativeColumn(int currentcolumn, int column, bool relative)
        {
            if (relative)
            {
                // mask out upper bits to produce 'wrapping' at column 256 ("IV")
                return (column + currentcolumn) & 0x00FF;
            }
            return column;
        }

        private static int FixupRelativeRow(int currentrow, int row, bool relative)
        {
            if (relative)
            {
                // mask out upper bits to produce 'wrapping' at row 65536
                return (row + currentrow) & 0x00FFFF;
            }
            return row;
        }
    }
}