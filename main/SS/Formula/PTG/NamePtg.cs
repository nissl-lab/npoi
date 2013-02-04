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
    using NPOI.Util;
    using NPOI.SS.Formula;


    /**
     *
     * @author  andy
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    [Serializable]
    public class NamePtg : OperandPtg, WorkbookDependentFormula
    {
        public const short sid = 0x23;
        private const int SIZE = 5;
        /** one-based index to defined name record */
        private int field_1_label_index;
        private short field_2_zero;   // reserved must be 0
        /**
 * @param nameIndex zero-based index to name within workbook
 */
        public NamePtg(int nameIndex)
        {
            field_1_label_index = 1 + nameIndex; // convert to 1-based
        }

        /** Creates new NamePtg */

        public NamePtg(ILittleEndianInput in1)
        {
            field_1_label_index = in1.ReadShort();
            field_2_zero = in1.ReadShort();
        }

        /**
         * @return zero based index to a defined name record in the LinkTable
         */
        public int Index
        {
            get { return field_1_label_index - 1; } // Convert to zero based
        }

        public override void Write(ILittleEndianOutput out1)
        {
		    out1.WriteByte(sid + PtgClass);
		    out1.WriteShort(field_1_label_index);
		    out1.WriteShort(field_2_zero);
        }


        public override int Size
        {
            get { return SIZE; }
        }

        public String ToFormulaString(IFormulaRenderingWorkbook book)
        {
            return book.GetNameText(this);
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