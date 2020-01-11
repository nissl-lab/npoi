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

namespace NPOI.HSSF.Record.CF
{
    using System;
    using System.Text;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;

    /**
     * Threshold / value (CFVO) for Changes in Conditional Formatting
     */
    public abstract class Threshold
    {
        private byte type;
        private Formula formula;
        private double? value;

        protected Threshold()
        {
            type = (byte)RangeType.NUMBER.id;
            formula = Formula.Create(null);
            value = 0d;
        }

        /** Creates new Threshold */
        protected Threshold(ILittleEndianInput in1)
        {
            type = (byte)in1.ReadByte();
            short formulaLen = in1.ReadShort();
            if (formulaLen > 0)
            {
                formula = Formula.Read(formulaLen, in1);
            }
            else
            {
                formula = Formula.Create(null);
            }
            // Value is only there for non-formula, non min/max thresholds
            if (formulaLen == 0 && type != RangeType.MIN.id &&
                    type != RangeType.MAX.id)
            {
                value = in1.ReadDouble();
            }
        }

        public byte Type
        {
            get
            {
                return type;
            }
            set
            {
                this.type = value;

                // Ensure the value presence / absence is consistent for the new type
                if (type == RangeType.MIN.id || type == RangeType.MAX.id ||
                       type == RangeType.FORMULA.id)
                {
                    this.value = null;
                }
                else if (this.value == null)
                {
                    this.value = 0d;
                }
            }
        }
        public void SetType(int type)
        {
            this.type = (byte)type;
        }

        protected Formula Formula
        {
            get
            {
                return formula;
            }
        }
        public Ptg[] ParsedExpression
        {
            get { return formula.Tokens; }
            set {
                formula = Formula.Create(value);
                if (value.Length > 0)
                {
                    this.value = null;
                }
            }
        }

        public double? Value
        {
            get
            {
                return this.value;
            }
            set
            {
                this.value = value;
            }
        }

        public virtual int DataLength
        {
            get
            {
                int len = 1 + formula.EncodedSize;
                if (value != null)
                {
                    len += 8;
                }
                return len;
            }
            
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("    [CF Threshold]\n");
            buffer.Append("          .type    = ").Append(HexDump.ToHex(type)).Append("\n");
            buffer.Append("          .Formula = ").Append(Arrays.ToString(formula.Tokens)).Append("\n");
            buffer.Append("          .value   = ").Append(value).Append("\n");
            buffer.Append("    [/CF Threshold]\n");
            return buffer.ToString();
        }

        public void CopyTo(Threshold rec)
        {
            rec.type = type;
            rec.formula = formula;
            rec.value = value;
        }

        public virtual void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(type);
            if (formula.Tokens.Length == 0)
            {
                out1.WriteShort(0);
            }
            else
            {
                formula.Serialize(out1);
            }
            if (value != null)
            {
                out1.WriteDouble(value.Value);
            }
        }
    }
}
