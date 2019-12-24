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
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;


    /**
     * Conditional Formatting Rule Record (0x01B1). 
     * 
     * <p>This is for the older-style Excel conditional formattings,
     *  new-style (Excel 2007+) also make use of {@link CFRule12Record}
     *  and {@link CFExRuleRecord} for their rules.
     */
    public class CFRuleRecord : CFRuleBase, ICloneable
    {
        public static short sid = 0x01B1;

        /** Creates new CFRuleRecord */
        private CFRuleRecord(byte conditionType, byte comparisonOperation)
                : base(conditionType, comparisonOperation)
        {

            SetDefaults();
        }

        private CFRuleRecord(byte conditionType, byte comparisonOperation, Ptg[] formula1, Ptg[] formula2)
            : base(conditionType, comparisonOperation, formula1, formula2)
        {

            SetDefaults();
        }
        private void SetDefaults() {
            // Set modification flags to 1: by default options are not modified
            formatting_options = modificationBits.SetValue(formatting_options, -1);
            // Set formatting block flags to 0 (no formatting blocks)
            formatting_options = fmtBlockBits.SetValue(formatting_options, 0);
            formatting_options = undocumented.Clear(formatting_options);
            unchecked
            {
                formatting_not_used = (short)0x8002; // Excel seems to write this value, but it doesn't seem to care what it Reads
            }

            _fontFormatting = null;
            _borderFormatting = null;
            _patternFormatting = null;
        }

        /**
         * Creates a new comparison operation rule
         */
        public static CFRuleRecord Create(HSSFSheet sheet, String formulaText) {
            Ptg[] formula1 = ParseFormula(formulaText, sheet);
            return new CFRuleRecord(CONDITION_TYPE_FORMULA, ComparisonOperator.NO_COMPARISON,
                    formula1, null);
        }
        /**
         * Creates a new comparison operation rule
         */
        public static CFRuleRecord Create(HSSFSheet sheet, byte comparisonOperation,
                String formulaText1, String formulaText2) {
            Ptg[] formula1 = ParseFormula(formulaText1, sheet);
            Ptg[] formula2 = ParseFormula(formulaText2, sheet);
            return new CFRuleRecord(CONDITION_TYPE_CELL_VALUE_IS, comparisonOperation, formula1, formula2);
        }

        public CFRuleRecord(RecordInputStream in1) {
            ConditionType = ((byte)in1.ReadByte());
            ComparisonOperation = ((byte)in1.ReadByte());
            int field_3_formula1_len = in1.ReadUShort();
            int field_4_formula2_len = in1.ReadUShort();
            ReadFormatOptions(in1);

            // "You may not use unions, intersections or array constants in Conditional Formatting criteria"
            Formula1 = (Formula.Read(field_3_formula1_len, in1));
            Formula2 = (Formula.Read(field_4_formula2_len, in1));
        }

        public override short Sid
        {
            get { return sid; }
        }

        /**
         * called by the class that is responsible for writing this sucker.
         * Subclasses should implement this so that their data is passed back in a
         * byte array.
         *
         * @param out the stream to write to
         */
        public override void Serialize(ILittleEndianOutput out1) {
            int formula1Len = GetFormulaSize(Formula1);
            int formula2Len = GetFormulaSize(Formula2);

            out1.WriteByte(ConditionType);
            out1.WriteByte(ComparisonOperation);
            out1.WriteShort(formula1Len);
            out1.WriteShort(formula2Len);

            SerializeFormattingBlock(out1);

            Formula1.SerializeTokens(out1);
            Formula2.SerializeTokens(out1);
        }

        protected override int DataSize
        {
            get
            {
                return 6 + FormattingBlockSize +
                   GetFormulaSize(Formula1) +
                   GetFormulaSize(Formula2);
            }
            
        }

        public override String ToString() {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[CFRULE]\n");
            buffer.Append("    .condition_type   =").Append(ConditionType).Append("\n");
            buffer.Append("    OPTION FLAGS=0x").Append(HexDump.ToHex(Options)).Append("\n");
            if (ContainsFontFormattingBlock) {
                buffer.Append(_fontFormatting.ToString()).Append("\n");
            }
            if (ContainsBorderFormattingBlock) {
                buffer.Append(_borderFormatting.ToString()).Append("\n");
            }
            if (ContainsPatternFormattingBlock) {
                buffer.Append(_patternFormatting.ToString()).Append("\n");
            }
            buffer.Append("    Formula 1 =").Append(Arrays.ToString(Formula1.Tokens)).Append("\n");
            buffer.Append("    Formula 2 =").Append(Arrays.ToString(Formula2.Tokens)).Append("\n");
            buffer.Append("[/CFRULE]\n");
            return buffer.ToString();
        }

        public override Object Clone() {
            CFRuleRecord rec = new CFRuleRecord(ConditionType, ComparisonOperation);
            base.CopyTo(rec);
            return rec;
        }
    }

}