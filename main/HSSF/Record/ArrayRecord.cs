/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * ARRAY (0x0221)<p/>
     * 
     * Treated in a similar way to SharedFormulaRecord
     * 
     * @author Josh Micich
     */
    public class ArrayRecord : SharedValueRecordBase
    {

        public const short sid = 0x0221;
        private const int OPT_ALWAYS_RECALCULATE = 0x0001;
        private const int OPT_CALCULATE_ON_OPEN = 0x0002;

        private int _options;
        private int _field3notUsed;
        private NPOI.SS.Formula.Formula _formula;

        public ArrayRecord(RecordInputStream in1)
            : base(in1)
        {

            _options = in1.ReadUShort();
            _field3notUsed = in1.ReadInt();
            int formulaTokenLen = in1.ReadUShort();
            int totalFormulaLen = in1.Available();
            _formula = NPOI.SS.Formula.Formula.Read(formulaTokenLen, in1, totalFormulaLen);
        }
        public ArrayRecord(NPOI.SS.Formula.Formula formula, CellRangeAddress8Bit range):base(range)
        {
            _options = 0; //YK: Excel 2007 leaves this field unset
            _field3notUsed = 0;
            _formula = formula;
        }
        public bool IsAlwaysRecalculate
        {
            get
            {
                return (_options & OPT_ALWAYS_RECALCULATE) != 0;
            }
        }
        public bool IsCalculateOnOpen
        {
            get
            {
                return (_options & OPT_CALCULATE_ON_OPEN) != 0;
            }
        }

        public Ptg[] FormulaTokens
        {
            get
            {
                return _formula.Tokens;
            }
        }

        protected override int ExtraDataSize
        {
            get
            {
                return 2 + 4
                    + _formula.EncodedSize;
            }
        }
        protected override void SerializeExtraData(ILittleEndianOutput out1)
        {
            out1.WriteShort(_options);
            out1.WriteInt(_field3notUsed);
            _formula.Serialize(out1);
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name).Append(" [ARRAY]\n");
            sb.Append(" range=").Append(Range.ToString()).Append("\n");
            sb.Append(" options=").Append(HexDump.ShortToHex(_options)).Append("\n");
            sb.Append(" notUsed=").Append(HexDump.IntToHex(_field3notUsed)).Append("\n");
            sb.Append(" formula:").Append("\n");
            Ptg[] ptgs = _formula.Tokens;
            for (int i = 0; i < ptgs.Length; i++)
            {
                Ptg ptg = ptgs[i];
                sb.Append(ptg.ToString()).Append(ptg.RVAType).Append("\n");
            }
            sb.Append("]");
            return sb.ToString();
        }

        public override Object Clone()
        {
            ArrayRecord rec = new ArrayRecord(_formula.Copy(), Range);

            // they both seem unused, but clone them nevertheless to have an exact copy
            rec._options = _options;
            rec._field3notUsed = _field3notUsed;

            return rec;
        }
    }
}