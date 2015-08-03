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
     * A Name, be that a Named Range or a Function / User Defined
     *  Function, addressed in the HSSF External Sheet style.
     *  
     * This is HSSF only, as it matches the HSSF file format way of
     *  referring to the sheet by an extern index. The XSSF equivalent
     *  is {@link NameXPxg}
     */
    [Serializable]
    public class NameXPtg : OperandPtg,WorkbookDependentFormula
    {
        public const short sid = 0x39;
        private const int SIZE = 7;

        /** index to REF entry in externsheet record */
        private int _sheetRefIndex;
        /** index to defined name or externname table(1 based) */
        private int _nameNumber;
        /** reserved must be 0 */
        private int _reserved;

        private NameXPtg(int sheetRefIndex, int nameNumber, int reserved)
        {
            _sheetRefIndex = sheetRefIndex;
            _nameNumber = nameNumber;
            _reserved = reserved;
        }

        /**
         * @param sheetRefIndex index to REF entry in externsheet record
         * @param nameIndex index to defined name or externname table
         */
        public NameXPtg(int sheetRefIndex, int nameIndex)
            :this(sheetRefIndex, nameIndex + 1, 0)
        {
            
        }

        public NameXPtg(ILittleEndianInput in1)
            :this(in1.ReadUShort(), in1.ReadUShort(), in1.ReadUShort())
        {

        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(_sheetRefIndex);
            out1.WriteShort(_nameNumber);
            out1.WriteShort(_reserved);

        }

        public override int Size
        {
            get { return SIZE; }
        }

        public String ToFormulaString(IFormulaRenderingWorkbook book)
        {
            // -1 to convert definedNameIndex from 1-based to zero-based
            return book.ResolveNameXText(this);
        }
        public override String ToFormulaString()
        {
            throw new NotImplementedException("3D references need a workbook to determine formula text");
        }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_VALUE; }
        }

        public int SheetRefIndex
        {
            get
            {
                return _sheetRefIndex;
            }
        }
        public int NameIndex
        {
            get
            {
                return _nameNumber - 1;
            }
        }
    }
}