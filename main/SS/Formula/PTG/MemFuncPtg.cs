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
    


    /**
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class MemFuncPtg : OperandPtg
    {

        public const byte sid = 0x29;
        private int field_1_len_ref_subexpression;

        /**Creates new function pointer from a byte array
         * usually called while Reading an excel file.
         */
        public MemFuncPtg(ILittleEndianInput in1)
            : this(in1.ReadUShort())
        {

        }

        public MemFuncPtg(int subExprLen)
        {
            field_1_len_ref_subexpression = subExprLen;
        }

        public override int Size
        {
            get { return 3; }
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(field_1_len_ref_subexpression);
        }

        public override String ToFormulaString()
        {
            return "";
        }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_REF; }
        }

        public int NumberOfOperands
        {
            get { return field_1_len_ref_subexpression; }
        }

        public int LenRefSubexpression
        {
            get { return field_1_len_ref_subexpression; }
        }
    }
}