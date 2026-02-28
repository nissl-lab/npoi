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
     * Greater than operator PTG ">"
     * @author  Cameron Riley (criley at ekmail.com)
     */
    public class GreaterThanPtg : ValueOperatorPtg
    {
        public const byte sid = 0x0D;
        private const string GREATERTHAN = ">";

        public static readonly ValueOperatorPtg instance = new GreaterThanPtg();

        private GreaterThanPtg()
        {
            // enforce singleton
        }

        protected override byte Sid
        {
            get { return sid; }
        }

        /**
         * Get the number of operands for the Less than operator
         * @return int the number of operands
         */
        public override int NumberOfOperands
        {
            get { return 2; }
        }

        /** 
         * Implementation of method from OperationsPtg
         * @param operands a String array of operands
         * @return String the Formula as a String
         */
        public override String ToFormulaString(String[] operands)
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append(operands[0]);
            buffer.Append(GREATERTHAN);
            buffer.Append(operands[1]);
            return buffer.ToString();
        }
    }
}