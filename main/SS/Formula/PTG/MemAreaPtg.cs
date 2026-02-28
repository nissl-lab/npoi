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


    /**
     * @author Daniel Noll (daniel at nuix dot com dot au)
     */
    public class MemAreaPtg : OperandPtg
    {
        public const short sid = 0x26;
        private const int SIZE = 7;
        private int field_1_reserved;
        private int field_2_subex_len;

        /** Creates new MemAreaPtg */

        public MemAreaPtg(int subexLen)
        {
            field_1_reserved = 0;
            field_2_subex_len = subexLen;
        }

        public MemAreaPtg(ILittleEndianInput in1)
        {
            field_1_reserved = in1.ReadInt();
            field_2_subex_len = in1.ReadShort();
        }

        public int Reserved
        {
            get { return field_1_reserved; }
            set { field_1_reserved = value; }
        }

        public int LenRefSubexpression
        {
            get { return field_2_subex_len; }
            set { field_2_subex_len = value; }
        }

        public override void Write(ILittleEndianOutput out1)
        {
		    out1.WriteByte(sid + PtgClass);
		    out1.WriteInt(field_1_reserved);
		    out1.WriteShort(field_2_subex_len);
        }


        public override int Size
        {
            get { return SIZE; }
        }

        public override String ToFormulaString()
        {
            return ""; // TODO: Not sure how to format this. -- DN
        }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_VALUE; }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(this.GetType().Name).Append(" [len=");
            sb.Append(field_2_subex_len);
            sb.Append("]");
            return sb.ToString();
        }
    }
}