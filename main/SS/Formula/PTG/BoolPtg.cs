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
     * bool (bool)
     * Stores a (java) bool value in a formula.
     * @author Paul Krause (pkrause at soundbite dot com)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class BoolPtg : ScalarConstantPtg
    {
        public const int SIZE = 2;
        public const byte sid = 0x1d;
        private bool field_1_value;

        public BoolPtg(ILittleEndianInput in1)
        {
            field_1_value = (in1.ReadByte() == 1);
        }


        public BoolPtg(String formulaToken)
        {
            field_1_value = (formulaToken.Equals("TRUE"));
        }

        public bool Value
        {
            get { return field_1_value; }
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteByte(field_1_value ? 1 : 0);
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public override String ToFormulaString()
        {
            return field_1_value ? "TRUE" : "FALSE";
        }
    }
}