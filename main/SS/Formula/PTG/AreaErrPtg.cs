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

    using NPOI.HSSF.UserModel;

    /**
     * AreaErr - handles deleted cell area references.
     *
     * @author Daniel Noll (daniel at nuix dot com dot au)
     */
    public class AreaErrPtg : OperandPtg
    {
        public const byte sid = 0x2b;
        private int unused1;
        private int unused2;

        public AreaErrPtg(ILittleEndianInput in1)
        {
            // 8 bytes unused:
            unused1 = in1.ReadInt();
            unused2 = in1.ReadInt();
        }
        public AreaErrPtg()
        {
            unused1 = 0;
            unused2 = 0;
        }
        public override void Write(ILittleEndianOutput out1)
        {
		    out1.WriteByte(sid + PtgClass);
		    out1.WriteInt(unused1);
		    out1.WriteInt(unused2);
        }

        public override String ToFormulaString()
        {
            return HSSFErrorConstants.GetText(HSSFErrorConstants.ERROR_REF);
        }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_REF; }
        }

        public override int Size
        {
            get { return 9; }
        }
    }
}