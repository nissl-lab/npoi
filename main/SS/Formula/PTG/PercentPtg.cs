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
     * Percent PTG.
     *
     * @author Daniel Noll (daniel at nuix.com.au)
     */
    public class PercentPtg : ValueOperatorPtg
    {
        public const int SIZE = 1;
        public const byte sid = 0x14;

        private const string PERCENT = "%";

        public static readonly ValueOperatorPtg instance = new PercentPtg();

        private PercentPtg()
        {
            // enforce singleton
        }

        protected override byte Sid
        {
            get { return sid; }
        }

        public override int NumberOfOperands
        {
            get { return 1; }
        }

        public override String ToFormulaString(String[] operands)
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append(operands[0]);
            buffer.Append(PERCENT);
            return buffer.ToString();
        }
    }
}
