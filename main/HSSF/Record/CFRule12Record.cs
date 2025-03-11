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
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.Record.Common;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using ExtendedColorR = NPOI.HSSF.Record.Common.ExtendedColor;

    /**
     * Conditional Formatting v12 Rule Record (0x087A). 
     * 
     * <p>This is for newer-style Excel conditional formattings,
     *  from Excel 2007 onwards.
     *  
     * <p>{@link CFRuleRecord} is used where the condition type is
     *  {@link #CONDITION_TYPE_CELL_VALUE_IS} or {@link #CONDITION_TYPE_FORMULA},
     *  this is only used for the other types
     */
    public class CFRule12Record : CFRuleBase, IFutureRecord, ICloneable
    {
        public static short sid = 0x087A;

        private FtrHeader futureHeader;
        private int ext_formatting_length;
        private byte[] ext_formatting_data;
        private Formula formula_scale;
        private byte ext_opts;
        private int priority;
        private int template_type;
        private byte template_param_length;
        private byte[] template_params;

        private DataBarFormatting data_bar;
        private IconMultiStateFormatting multistate;
        private ColorGradientFormatting color_gradient;
        // TODO Parse this, see #58150
        private byte[] filter_data;

        /** Creates new CFRuleRecord */
        private CFRule12Record(byte conditionType, byte comparisonOperation)
                : base(conditionType, comparisonOperation)
        {

            SetDefaults();
        }

        private CFRule12Record(byte conditionType, byte comparisonOperation, Ptg[] formula1, Ptg[] formula2, Ptg[] formulaScale)
                : base(conditionType, comparisonOperation, formula1, formula2)
        {

            SetDefaults();
            this.formula_scale = Formula.Create(formulaScale);
        }
        private void SetDefaults() {
            futureHeader = new FtrHeader();
            futureHeader.RecordType = (/*setter*/sid);

            ext_formatting_length = 0;
            ext_formatting_data = new byte[4];

            formula_scale = Formula.Create(Ptg.EMPTY_PTG_ARRAY);

            ext_opts = 0;
            priority = 0;
            template_type = ConditionType;
            template_param_length = 16;
            template_params = new byte[template_param_length];
        }

        /**
         * Creates a new comparison operation rule
         */
        public static CFRule12Record Create(HSSFSheet sheet, String formulaText) {
            Ptg[] formula1 = ParseFormula(formulaText, sheet);
            return new CFRule12Record(CONDITION_TYPE_FORMULA, ComparisonOperator.NO_COMPARISON,
                    formula1, null, null);
        }
        /**
         * Creates a new comparison operation rule
         */
        public static CFRule12Record Create(HSSFSheet sheet, byte comparisonOperation,
                String formulaText1, String formulaText2) {
            Ptg[] formula1 = ParseFormula(formulaText1, sheet);
            Ptg[] formula2 = ParseFormula(formulaText2, sheet);
            return new CFRule12Record(CONDITION_TYPE_CELL_VALUE_IS, comparisonOperation,
                    formula1, formula2, null);
        }
        /**
         * Creates a new comparison operation rule
         */
        public static CFRule12Record Create(HSSFSheet sheet, byte comparisonOperation,
                String formulaText1, String formulaText2, String formulaTextScale) {
            Ptg[] formula1 = ParseFormula(formulaText1, sheet);
            Ptg[] formula2 = ParseFormula(formulaText2, sheet);
            Ptg[] formula3 = ParseFormula(formulaTextScale, sheet);
            return new CFRule12Record(CONDITION_TYPE_CELL_VALUE_IS, comparisonOperation,
                    formula1, formula2, formula3);
        }
        /// <summary>
        /// Creates a new Data Bar formatting
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static CFRule12Record Create(HSSFSheet sheet, ExtendedColorR color) {
            CFRule12Record r = new CFRule12Record(CONDITION_TYPE_DATA_BAR,
                                                  ComparisonOperator.NO_COMPARISON);
            DataBarFormatting dbf = r.CreateDataBarFormatting();
            dbf.Color = color;
            dbf.PercentMin = (byte)0;
            dbf.PercentMax = (byte)100;

            DataBarThreshold min = new DataBarThreshold();
            min.SetType(RangeType.MIN.id);
            dbf.ThresholdMin = min;

            DataBarThreshold max = new DataBarThreshold();
            max.SetType(RangeType.MAX.id);
            dbf.ThresholdMax = max;

            return r;
        }
        /// <summary>
        /// Creates a new Icon Set / Multi-State formatting
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="iconSet"></param>
        /// <returns></returns>
        public static CFRule12Record Create(HSSFSheet sheet, IconSet iconSet) {
            Threshold[] ts = new Threshold[iconSet.num];
            for (int i = 0; i < ts.Length; i++) {
                ts[i] = new IconMultiStateThreshold();
            }

            CFRule12Record r = new CFRule12Record(CONDITION_TYPE_ICON_SET,
                                                  ComparisonOperator.NO_COMPARISON);
            IconMultiStateFormatting imf = r.CreateMultiStateFormatting();
            imf.IconSet = iconSet;
            imf.Thresholds = ts;
            return r;
        }
        /// <summary>
        /// Creates a new Color Scale / Color Gradient formatting
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static CFRule12Record CreateColorScale(HSSFSheet sheet) {
            int numPoints = 3;
            ExtendedColorR[] colors = new ExtendedColorR[numPoints];
            ColorGradientThreshold[] ts = new ColorGradientThreshold[numPoints];
            for (int i = 0; i < ts.Length; i++) {
                ts[i] = new ColorGradientThreshold();
                colors[i] = new ExtendedColorR();
            }

            CFRule12Record r = new CFRule12Record(CONDITION_TYPE_COLOR_SCALE,
                                                  ComparisonOperator.NO_COMPARISON);
            ColorGradientFormatting cgf = r.CreateColorGradientFormatting();
            cgf.NumControlPoints = (/*setter*/numPoints);
            cgf.Thresholds = (/*setter*/ts);
            cgf.Colors = (/*setter*/colors);
            return r;
        }

        public CFRule12Record(RecordInputStream in1) {
            futureHeader = new FtrHeader(in1);
            ConditionType = ((byte)in1.ReadByte());
            ComparisonOperation = ((byte)in1.ReadByte());
            int field_3_formula1_len = in1.ReadUShort();
            int field_4_formula2_len = in1.ReadUShort();

            ext_formatting_length = in1.ReadInt();
            ext_formatting_data = new byte[0];
            if (ext_formatting_length == 0) {
                // 2 bytes reserved
                in1.ReadUShort();
            } else {
                int len = ReadFormatOptions(in1);
                if (len < ext_formatting_length) {
                    ext_formatting_data = new byte[ext_formatting_length - len];
                    in1.ReadFully(ext_formatting_data);
                }
            }

            Formula1 = (Formula.Read(field_3_formula1_len, in1));
            Formula2 = (Formula.Read(field_4_formula2_len, in1));

            int formula_scale_len = in1.ReadUShort();
            formula_scale = Formula.Read(formula_scale_len, in1);

            ext_opts = (byte)in1.ReadByte();
            priority = in1.ReadUShort();
            template_type = in1.ReadUShort();
            template_param_length = (byte)in1.ReadByte();
            if (template_param_length == 0 || template_param_length == 16) {
                template_params = new byte[template_param_length];
                in1.ReadFully(template_params);
            } else {
                //logger.Log(POILogger.WARN, "CF Rule v12 template params length should be 0 or 16, found " + template_param_length);
                in1.ReadRemainder();
            }

            byte type = ConditionType;
            if (type == CONDITION_TYPE_COLOR_SCALE) {
                color_gradient = new ColorGradientFormatting(in1);
            } else if (type == CONDITION_TYPE_DATA_BAR) {
                data_bar = new DataBarFormatting(in1);
            } else if (type == CONDITION_TYPE_FILTER) {
                filter_data = in1.ReadRemainder();
            } else if (type == CONDITION_TYPE_ICON_SET) {
                multistate = new IconMultiStateFormatting(in1);
            }
        }

        public bool ContainsDataBarBlock() {
            return (data_bar != null);
        }
        public DataBarFormatting DataBarFormatting
        {
            get { return data_bar; }
        }
        public DataBarFormatting CreateDataBarFormatting() {
            if (data_bar != null) return data_bar;

            // Convert, Setup and return
            ConditionType = (CONDITION_TYPE_DATA_BAR);
            data_bar = new DataBarFormatting();
            return data_bar;
        }

        public bool ContainsMultiStateBlock() {
            return (multistate != null);
        }
        public IconMultiStateFormatting MultiStateFormatting
        {
            get { return multistate; }
        }
        public IconMultiStateFormatting CreateMultiStateFormatting() {
            if (multistate != null) return multistate;

            // Convert, Setup and return
            ConditionType = (CONDITION_TYPE_ICON_SET);
            multistate = new IconMultiStateFormatting();
            return multistate;
        }

        public bool ContainsColorGradientBlock() {
            return (color_gradient != null);
        }
        public ColorGradientFormatting ColorGradientFormatting
        {
            get { return color_gradient; }
        }
        public ColorGradientFormatting CreateColorGradientFormatting() {
            if (color_gradient != null) return color_gradient;

            // Convert, Setup and return
            ConditionType = (CONDITION_TYPE_COLOR_SCALE);
            color_gradient = new ColorGradientFormatting();
            return color_gradient;
        }

        /**
         * Get the stack of the scale expression as a list
         *
         * @return list of tokens (casts stack to a list and returns it!)
         * this method can return null is we are unable to create Ptgs from
         *	 existing excel file
         * callers should check for null!
         */
        public Ptg[] ParsedExpressionScale
        {
            get
            {
                return formula_scale.Tokens;
            }
            set
            {
                formula_scale = Formula.Create(value);
            }
        }

        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        /**
         * called by the class that is responsible for writing this sucker.
         * Subclasses should implement this so that their data is passed back in a
         * byte array.
         *
         * @param out the stream to write to
         */
        public override void Serialize(ILittleEndianOutput out1) {
            futureHeader.Serialize(out1);

            int formula1Len = GetFormulaSize(Formula1);
            int formula2Len = GetFormulaSize(Formula2);

            out1.WriteByte(ConditionType);
            out1.WriteByte(ComparisonOperation);
            out1.WriteShort(formula1Len);
            out1.WriteShort(formula2Len);

            // TODO Update ext_formatting_length
            if (ext_formatting_length == 0) {
                out1.WriteInt(0);
                out1.WriteShort(0);
            } else {
                out1.WriteInt(ext_formatting_length);
                SerializeFormattingBlock(out1);
                out1.Write(ext_formatting_data);
            }

            Formula1.SerializeTokens(out1);
            Formula2.SerializeTokens(out1);
            out1.WriteShort(GetFormulaSize(formula_scale));
            formula_scale.SerializeTokens(out1);

            out1.WriteByte(ext_opts);
            out1.WriteShort(priority);
            out1.WriteShort(template_type);
            out1.WriteByte(template_param_length);
            out1.Write(template_params);

            byte type = ConditionType;
            if (type == CONDITION_TYPE_COLOR_SCALE) {
                color_gradient.Serialize(out1);
            } else if (type == CONDITION_TYPE_DATA_BAR) {
                data_bar.Serialize(out1);
            } else if (type == CONDITION_TYPE_FILTER) {
                out1.Write(filter_data);
            } else if (type == CONDITION_TYPE_ICON_SET) {
                multistate.Serialize(out1);
            }
        }

        protected override int DataSize
        {
            get
            {
                int len = FtrHeader.GetDataSize() + 6;
                if (ext_formatting_length == 0)
                {
                    len += 6;
                }
                else
                {
                    len += 4 + FormattingBlockSize + ext_formatting_data.Length;
                }
                len += GetFormulaSize(Formula1);
                len += GetFormulaSize(Formula2);
                len += 2 + GetFormulaSize(formula_scale);
                len += 6 + template_params.Length;

                byte type = ConditionType;
                if (type == CONDITION_TYPE_COLOR_SCALE)
                {
                    len += color_gradient.DataLength;
                }
                else if (type == CONDITION_TYPE_DATA_BAR)
                {
                    len += data_bar.DataLength;
                }
                else if (type == CONDITION_TYPE_FILTER)
                {
                    len += filter_data.Length;
                }
                else if (type == CONDITION_TYPE_ICON_SET)
                {
                    len += multistate.DataLength;
                }
                return len;
            }
            
        }

        public override String ToString() {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[CFRULE12]\n");
            buffer.Append("    .condition_type=").Append(ConditionType).Append("\n");
            buffer.Append("    .dxfn12_length =0x").Append(HexDump.ToHex(ext_formatting_length)).Append("\n");
            buffer.Append("    .option_flags  =0x").Append(HexDump.ToHex(Options)).Append("\n");
            if (ContainsFontFormattingBlock) {
                buffer.Append(_fontFormatting.ToString()).Append("\n");
            }
            if (ContainsBorderFormattingBlock) {
                buffer.Append(_borderFormatting.ToString()).Append("\n");
            }
            if (ContainsPatternFormattingBlock) {
                buffer.Append(_patternFormatting.ToString()).Append("\n");
            }
            buffer.Append("    .dxfn12_ext=").Append(HexDump.ToHex(ext_formatting_data)).Append("\n");
            buffer.Append("    .Formula_1 =").Append(Arrays.ToString(Formula1.Tokens)).Append("\n");
            buffer.Append("    .Formula_2 =").Append(Arrays.ToString(Formula2.Tokens)).Append("\n");
            buffer.Append("    .Formula_S =").Append(Arrays.ToString(formula_scale.Tokens)).Append("\n");
            buffer.Append("    .ext_opts  =").Append(ext_opts).Append("\n");
            buffer.Append("    .priority  =").Append(priority).Append("\n");
            buffer.Append("    .template_type  =").Append(template_type).Append("\n");
            buffer.Append("    .template_params=").Append(HexDump.ToHex(template_params)).Append("\n");
            buffer.Append("    .filter_data    =").Append(HexDump.ToHex(filter_data)).Append("\n");
            if (color_gradient != null) {
                buffer.Append(color_gradient);
            }
            if (multistate != null) {
                buffer.Append(multistate);
            }
            if (data_bar != null) {
                buffer.Append(data_bar);
            }
            buffer.Append("[/CFRULE12]\n");
            return buffer.ToString();
        }

        public override Object Clone() {
            CFRule12Record rec = new CFRule12Record(ConditionType, ComparisonOperation);
            rec.futureHeader.AssociatedRange = (/*setter*/futureHeader.AssociatedRange.Copy());

            base.CopyTo(rec);

            // use min() to gracefully handle cases where the length-property and the array-lenght do not match
            // we saw some such files in circulation
            rec.ext_formatting_length = Math.Min(ext_formatting_length, ext_formatting_data.Length);
            rec.ext_formatting_data = new byte[ext_formatting_length];
            Array.Copy(ext_formatting_data, 0, rec.ext_formatting_data, 0, rec.ext_formatting_length);

            rec.formula_scale = formula_scale.Copy();

            rec.ext_opts = ext_opts;
            rec.priority = priority;
            rec.template_type = template_type;
            rec.template_param_length = template_param_length;
            rec.template_params = new byte[template_param_length];
            Array.Copy(template_params, 0, rec.template_params, 0, template_param_length);

            if (color_gradient != null) {
                rec.color_gradient = (ColorGradientFormatting)color_gradient.Clone();
            }
            if (multistate != null) {
                rec.multistate = (IconMultiStateFormatting)multistate.Clone();
            }
            if (data_bar != null) {
                rec.data_bar = (DataBarFormatting)data_bar.Clone();
            }
            if (filter_data != null) {
                rec.filter_data = new byte[filter_data.Length];
                Array.Copy(filter_data, 0, rec.filter_data, 0, filter_data.Length);
            }

            return rec;
        }

        public short GetFutureRecordType() {
            return futureHeader.RecordType;
        }
        public FtrHeader GetFutureHeader() {
            return futureHeader;
        }
        public CellRangeAddress GetAssociatedRange() {
            return futureHeader.AssociatedRange;
        }

        public int Priority
        {
            get
            {
                return priority;
            }
            set {
                this.priority = value;
            }
        }
    }

}