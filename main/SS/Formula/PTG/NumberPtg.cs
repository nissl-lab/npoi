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

    using NPOI.SS.Util;
    using System.Globalization;

    /**
     * Number
     * Stores a floating point value in a formula
     * value stored in a 8 byte field using IEEE notation
     * @author  Avik Sengupta
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class NumberPtg : ScalarConstantPtg
    {
        public const int SIZE = 9;
        public const byte sid = 0x1f;
        private double field_1_value;

        /** Create a NumberPtg from a byte array Read from disk */
        public NumberPtg(ILittleEndianInput in1)
        {
            field_1_value = in1.ReadDouble();
        }
        /** Create a NumberPtg from a string representation of  the number
         *  Number format is not checked, it is expected to be validated in the parser
         *   that calls this method. 
         *  @param value : String representation of a floating point number
         */
        public NumberPtg(String value)
            : this(Double.Parse(value, CultureInfo.InvariantCulture))
        {
            
        }

        public NumberPtg(double value)
        {
            field_1_value = value;
        }

        public double Value
        {
            get { return field_1_value; }
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteDouble(Value);
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public override String ToFormulaString()
        {
            return NumberToTextConverter.ToText(Value);
        }
    }
}