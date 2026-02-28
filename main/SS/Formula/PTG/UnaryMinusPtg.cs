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
     * Unary Plus operator
     * does not have any effect on the operand
     * @author Avik Sengupta
     */
    public class UnaryMinusPtg : ValueOperatorPtg
    {
        public const byte sid = 0x13;

        private const string MINUS = "-";

        public static readonly ValueOperatorPtg instance = new UnaryMinusPtg();

        private UnaryMinusPtg()
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

        /** implementation of method from OperationsPtg*/
        public override String ToFormulaString(String[] operands)
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append(MINUS);
            buffer.Append(operands[0]);
            return buffer.ToString();
        }
    }
}