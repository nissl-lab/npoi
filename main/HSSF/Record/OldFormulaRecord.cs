/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;

    /**
     * Formula Record (0x0006 / 0x0206 / 0x0406) - holds a formula in
     *  encoded form, along with the value if a number
     */
    public class OldFormulaRecord : OldCellRecord
    {
        public const short biff2_sid = 0x0006;
        public const short biff3_sid = 0x0206;
        public const short biff4_sid = 0x0406;
        public const short biff5_sid = 0x0006;

        private SpecialCachedValue specialCachedValue;
        private double field_4_value;
        private short field_5_options;
        private Formula field_6_Parsed_expr;

        public OldFormulaRecord(RecordInputStream ris) :
            base(ris, ris.Sid == biff2_sid)
        {
            ;

            if (IsBiff2)
            {
                field_4_value = ris.ReadDouble();
            }
            else
            {
                long valueLongBits = ris.ReadLong();
                specialCachedValue = SpecialCachedValue.Create(valueLongBits);
                if (specialCachedValue == null)
                {
                    field_4_value = BitConverter.Int64BitsToDouble(valueLongBits);
                }
            }

            if (IsBiff2)
            {
                field_5_options = (short)ris.ReadUByte();
            }
            else
            {
                field_5_options = ris.ReadShort();
            }

            int expression_len = ris.ReadShort();
            int nBytesAvailable = ris.Available();
            field_6_Parsed_expr = Formula.Read(expression_len, ris, nBytesAvailable);
        }

        public CellType GetCachedResultType()
        {
            if (specialCachedValue == null)
            {
                return CellType.Numeric;
            }
            return specialCachedValue.GetValueType();
        }

        public bool GetCachedBooleanValue()
        {
            return specialCachedValue.GetBooleanValue();
        }
        public int GetCachedErrorValue()
        {
            return specialCachedValue.GetErrorValue();
        }

        /**
         * Get the calculated value of the formula
         *
         * @return calculated value
         */
        public double Value
        {
            get
            {
                return field_4_value;
            }
        }

        /**
         * Get the option flags
         *
         * @return bitmask
         */
        public short Options
        {
            get
            {
                return field_5_options;
            }
        }

        /**
         * @return the formula tokens. never <code>null</code>
         */
        public Ptg[] ParsedExpression
        {
            get
            {
                return field_6_Parsed_expr.Tokens;
            }
        }

        public Formula Formula
        {
            get
            {
                return field_6_Parsed_expr;
            }
        }

        protected override void AppendValueText(StringBuilder sb)
        {
            sb.Append("    .value       = ").Append(Value).Append("\n");
        }
        protected override String RecordName
        {
            get
            {
                return "Old Formula";
            }
        }
    }

}