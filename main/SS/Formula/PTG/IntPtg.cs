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
    using System.Globalization;
    


    /**
     * Integer (unsigned short integer)
     * Stores an Unsigned short value (java int) in a formula
     * @author  Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class IntPtg : ScalarConstantPtg
    {
        // 16 bit Unsigned integer
        private const int MIN_VALUE = 0x0000;
        private const int MAX_VALUE = 0xFFFF;

        /**
         * Excel represents integers 0..65535 with the tInt token. 
         * @return <c>true</c> if the specified value is within the range of values 
         * <c>IntPtg</c> can represent. 
         */
        public static bool IsInRange(int i)
        {
            return i >= MIN_VALUE && i <= MAX_VALUE;
        }

        public const int SIZE = 3;
        public const byte sid = 0x1e;
        private int field_1_value;

        public IntPtg(ILittleEndianInput in1)
            : this(in1.ReadUShort())
        {
            
        }

        public IntPtg(int value)
        {
            if (!IsInRange(value))
            {
                throw new ArgumentException("value is out of range: " + value);
            }
            field_1_value = value;
        }

        public int Value
        {
            get { return field_1_value; }
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(Value);
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public override String ToFormulaString()
        {
            return Value.ToString(CultureInfo.InvariantCulture);
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(field_1_value);
            sb.Append("]");
            return sb.ToString();
        }
    }
}