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

namespace NPOI.SS.Formula.PTG
{
    using System;
    using System.Text;
    using NPOI.Util;
    
    using NPOI.SS.Util;
    using NPOI.SS.Formula;



    /**
     * Title:        Reference 3D Ptg 
     * Description:  Defined a cell in extern sheet. 
     * REFERENCE:  
     * @author Libin Roman (Vista Portal LDT. Developer)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 1.0-pre
     */
    public class Ref3DPtg : RefPtgBase, WorkbookDependentFormula, IExternSheetReferenceToken
    {
        public const byte sid = 0x3a;
        private const int SIZE = 7; // 6 + 1 for Ptg
        private int field_1_index_extern_sheet;
        /** Field 2 
         * - lower 8 bits is the zero based Unsigned byte column index 
         * - bit 16 - IsRowRelative
         * - bit 15 - IsColumnRelative 
         */
        private BitField rowRelative = BitFieldFactory.GetInstance(0x8000);
        private BitField colRelative = BitFieldFactory.GetInstance(0x4000);

        /** Creates new AreaPtg */
        public Ref3DPtg() { }

        public Ref3DPtg(ILittleEndianInput in1)
        {
            field_1_index_extern_sheet = in1.ReadShort();
            ReadCoordinates(in1);
        }

        public Ref3DPtg(String cellref, int externIdx)
        {
            CellReference c = new CellReference(cellref);
            Row=c.Row;
            Column=c.Col;
            IsColRelative=!c.IsColAbsolute;
            IsRowRelative=!c.IsRowAbsolute;
            ExternSheetIndex=externIdx;
        }

        public Ref3DPtg(CellReference cr, int externIdx):base(cr)
        {
            ExternSheetIndex = externIdx;
        }

        public override String ToString()
        {
            CellReference cr = new CellReference(Row, Column, !IsRowRelative, !IsColRelative);
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name);
            sb.Append(" [");
            sb.Append("sheetIx=").Append(ExternSheetIndex);
            sb.Append(" ! ");
            sb.Append(cr.FormatAsString());
            sb.Append("]");
            return sb.ToString();
        }

        public override void Write(ILittleEndianOutput out1)
        {

            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(ExternSheetIndex);
            WriteCoordinates(out1);
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public int ExternSheetIndex
        {
            get { return field_1_index_extern_sheet; }
            set { field_1_index_extern_sheet = value; }
        }
        /**
         * @return text representation of this cell reference that can be used in text 
         * formulas. The sheet name will Get properly delimited if required.
         */
        public String ToFormulaString(IFormulaRenderingWorkbook book)
        {
            return ExternSheetNameResolver.PrependSheetName(book, field_1_index_extern_sheet, FormatReferenceAsString());
        }

        public override String ToFormulaString()
        {
            throw new NotImplementedException("3D references need a workbook to determine formula text");
        }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_REF; }
        }
    }
}