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
    
    using NPOI.HSSF.Record;
    using NPOI.Util;


    /**
     *
     * @author  andy
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class UnknownPtg : Ptg
    {
        private short size = 1;

        /** Creates new UnknownPtg */

        public UnknownPtg()
        {
        }

        public UnknownPtg(RecordInputStream in1)
        {
            // doesn't need anything
        }

        public override bool IsBaseToken
        {
            get { return true; }
        }
        public override void Write(ILittleEndianOutput out1)
        {
        }

        public override int Size
        {
            get { return size; }
        }

        public override String ToFormulaString()
        {
            return "UNKNOWN";
        }
        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_VALUE; }
        }

        public override Object Clone()
        {
            return new UnknownPtg();
        }
    }
}