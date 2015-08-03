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
    using NPOI.SS.Formula;
    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * <p>Title:        Area 3D Ptg - 3D reference (Sheet + Area)</p>
     * <p>Description:  Defined an area in Extern Sheet. </p>
     * <p>REFERENCE:  </p>
     * 
     * This is HSSF only, as it matches the HSSF file format way of
     *  referring to the sheet by an extern index. The XSSF equivalent
     *  is {@link Area3DPxg}
     */
    [Serializable]
    public class Area3DPtg : AreaPtgBase, WorkbookDependentFormula, IExternSheetReferenceToken
    {
        public const byte sid = 0x3b;
        private const int SIZE = 11; // 10 + 1 for Ptg
        private int field_1_index_extern_sheet;

        private BitField rowRelative = BitFieldFactory.GetInstance(0x8000);
        private BitField colRelative = BitFieldFactory.GetInstance(0x4000);


        public Area3DPtg(String arearef, int externIdx):base(arearef)
        {
            ExternSheetIndex=externIdx;

        }

        public Area3DPtg(AreaReference arearef, int externIdx):base(arearef)
        {
            ExternSheetIndex=(externIdx);
        }
        public Area3DPtg(ILittleEndianInput in1)
        {
            field_1_index_extern_sheet = in1.ReadShort();
            ReadCoordinates(in1);
        }

        public Area3DPtg(int firstRow, int lastRow, int firstColumn, int lastColumn,
                bool firstRowRelative, bool lastRowRelative, bool firstColRelative, bool lastColRelative,
                int externalSheetIndex) :
            base(firstRow, lastRow, firstColumn, lastColumn, firstRowRelative, lastRowRelative, firstColRelative, lastColRelative)
        {
            ExternSheetIndex= externalSheetIndex;
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name);
            sb.Append(" [");
            sb.Append("sheetIx=").Append(ExternSheetIndex);
            sb.Append(" ! ");
            sb.Append(FormatReferenceAsString());
            sb.Append("]");
            return sb.ToString();
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(field_1_index_extern_sheet);
            WriteCoordinates(out1);
        }
        public override int Size
        {
            get { return SIZE; }
        }

        public int ExternSheetIndex
        {
            get{return field_1_index_extern_sheet;}
            set { field_1_index_extern_sheet = value; }
        }


        /*public String Area{
            RangeAddress ra = new RangeAddress( FirstColumn,FirstRow + 1, LastColumn, LastRow + 1);
            String result = ra.GetAddress();

            return result;
        }*/

        public void SetArea(String ref1)
        {
            AreaReference ar = new AreaReference(ref1);

            CellReference frstCell = ar.FirstCell;
            CellReference lastCell = ar.LastCell;

            FirstRow=(short)frstCell.Row;
            FirstColumn=frstCell.Col;
            LastRow=(short)lastCell.Row;
            LastColumn=lastCell.Col;
            IsFirstColRelative=!frstCell.IsColAbsolute;
            IsLastColRelative=!lastCell.IsColAbsolute;
            IsFirstRowRelative=!frstCell.IsRowAbsolute;
            IsLastRowRelative=!lastCell.IsRowAbsolute;
        }

        public override String ToFormulaString()
        {
            throw new NotImplementedException("3D references need a workbook to determine formula text");
        }
        /**
 * @return text representation of this area reference that can be used in text
 *  formulas. The sheet name will get properly delimited if required.
 */
        public String ToFormulaString(IFormulaRenderingWorkbook book)
        {
            return ExternSheetNameResolver.PrependSheetName(book, field_1_index_extern_sheet, FormatReferenceAsString());
        }
        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_REF; }
        }

    }
}