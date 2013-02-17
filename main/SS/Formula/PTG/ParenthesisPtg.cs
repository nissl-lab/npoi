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
     * While formula tokens are stored in RPN order and thus do not need parenthesis for 
     * precedence reasons, Parenthesis tokens ARE written to Ensure that user entered
     * parenthesis are Displayed as-is on Reading back
     *
     * Avik Sengupta &lt;lists@aviksengupta.com&gt;
     * Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class ParenthesisPtg : ControlPtg
    {

        private const int SIZE = 1;
        public const byte sid = 0x15;

        public static ControlPtg instance = new ParenthesisPtg();
        private ParenthesisPtg()
        {
            // enforce singleton
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public override String ToFormulaString()
        {
            return "()";
        }


        public String ToFormulaString(String[] operands)
        {
            return "(" + operands[0] + ")";
        }
    }
}