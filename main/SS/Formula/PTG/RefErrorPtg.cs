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
     * RefError - handles deleted cell reference
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class RefErrorPtg : OperandPtg
    {

        private const int SIZE = 5;
        public const byte sid = 0x2a;
        private int field_1_reserved;

        public RefErrorPtg()
        {
            field_1_reserved = 0;
        }

        public RefErrorPtg(ILittleEndianInput in1)
        {
            field_1_reserved = in1.ReadInt();

        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder("[RefError]\n");

            buffer.Append("reserved = ").Append(Reserved).Append("\n");
            return buffer.ToString();
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteInt(field_1_reserved);
        }
        public int Reserved
        {
            get{return field_1_reserved;}
            set { field_1_reserved = value; }
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public override String ToFormulaString()
        {
            //TODO -- should we store a cellreference instance in this ptg?? but .. memory is an Issue, i believe!
            return "#REF!";
        }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_REF; }
        }
    }
}