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
     *
     * @author  andy
     * @author Jason Height (jheight at chariot dot net dot au)
     * @author Daniel Noll (daniel at nuix dot com dot au)
     */

    public class MemErrPtg : MemAreaPtg
    {
        public new const short sid = 0x27;
        // fix warning CS0414 "never used": private const int SIZE = 7;
	    private int field_1_reserved = 0;
	    private short field_2_subex_len = 0;

        /** Creates new MemErrPtg */

        public MemErrPtg(ILittleEndianInput in1)
            : base(in1)
        {
        
        }

        public override void Write(ILittleEndianOutput out1)
        {
		    out1.WriteByte(sid + PtgClass);
		    out1.WriteInt(field_1_reserved);
		    out1.WriteShort(field_2_subex_len);
        }


        public override String ToFormulaString()
        {
            return "ERR#";
        }
    }
}