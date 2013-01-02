        
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

    /**
     * Ptg class to implement less than or equal
     *
     * @author fred at stsci dot edu
     */
    public class LessEqualPtg : ValueOperatorPtg
    {
        public const byte sid = 0x0a;

        public static ValueOperatorPtg instance = new LessEqualPtg();

        private LessEqualPtg()
        {
            // enforce singleton
        }

        protected override byte Sid
        {
            get { return sid; }
        }

        public override int NumberOfOperands
        {
            get { return 2; }
        }

        public override String ToFormulaString(String[] operands)
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append(operands[0]);
            buffer.Append("<=");
            buffer.Append(operands[1]);
            return buffer.ToString();
        }
    }
}