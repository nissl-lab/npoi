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

    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;


    /**
     * Conditional Formatting Rules. This can hold old-style rules
     *   
     * 
     * <p>This is for the older-style Excel conditional formattings,
     *  new-style (Excel 2007+) also make use of {@link CFRule12Record}
     *  and {@link CFExRuleRecord} for their rules.
     */
    public abstract class CFRuleBase : StandardRecord, ICloneable
    {
        public static class ComparisonOperator {
            public static byte NO_COMPARISON = 0;
            public static byte BETWEEN = 1;
            public static byte NOT_BETWEEN = 2;
            public static byte EQUAL = 3;
            public static byte NOT_EQUAL = 4;
            public static byte GT = 5;
            public static byte LT = 6;
            public static byte GE = 7;
            public static byte LE = 8;
            public static byte max_operator = 8;
        }

        private byte condition_type;
        // The only kinds that CFRuleRecord handles
        public const byte CONDITION_TYPE_CELL_VALUE_IS = 1;
        public const byte CONDITION_TYPE_FORMULA = 2;
        // These are CFRule12Rule only
        public const byte CONDITION_TYPE_COLOR_SCALE = 3;
        public const byte CONDITION_TYPE_DATA_BAR = 4;
        public const byte CONDITION_TYPE_FILTER = 5;
        public const byte CONDITION_TYPE_ICON_SET = 6;

        private byte comparison_operator;

        public static int TEMPLATE_CELL_VALUE = 0x0000;
        public static int TEMPLATE_FORMULA = 0x0001;
        public static int TEMPLATE_COLOR_SCALE_FORMATTING = 0x0002;
        public static int TEMPLATE_DATA_BAR_FORMATTING = 0x0003;
        public static int TEMPLATE_ICON_SET_FORMATTING = 0x0004;
        public static int TEMPLATE_FILTER = 0x0005;
        public static int TEMPLATE_UNIQUE_VALUES = 0x0007;
        public static int TEMPLATE_CONTAINS_TEXT = 0x0008;
        public static int TEMPLATE_CONTAINS_BLANKS = 0x0009;
        public static int TEMPLATE_CONTAINS_NO_BLANKS = 0x000A;
        public static int TEMPLATE_CONTAINS_ERRORS = 0x000B;
        public static int TEMPLATE_CONTAINS_NO_ERRORS = 0x000C;
        public static int TEMPLATE_TODAY = 0x000F;
        public static int TEMPLATE_TOMORROW = 0x0010;
        public static int TEMPLATE_YESTERDAY = 0x0011;
        public static int TEMPLATE_LAST_7_DAYS = 0x0012;
        public static int TEMPLATE_LAST_MONTH = 0x0013;
        public static int TEMPLATE_NEXT_MONTH = 0x0014;
        public static int TEMPLATE_THIS_WEEK = 0x0015;
        public static int TEMPLATE_NEXT_WEEK = 0x0016;
        public static int TEMPLATE_LAST_WEEK = 0x0017;
        public static int TEMPLATE_THIS_MONTH = 0x0018;
        public static int TEMPLATE_ABOVE_AVERAGE = 0x0019;
        public static int TEMPLATE_BELOW_AVERAGE = 0x001A;
        public static int TEMPLATE_DUPLICATE_VALUES = 0x001B;
        public static int TEMPLATE_ABOVE_OR_EQUAL_TO_AVERAGE = 0x001D;
        public static int TEMPLATE_BELOW_OR_EQUAL_TO_AVERAGE = 0x001E;

        internal static BitField modificationBits = bf(0x003FFFFF); // Bits: font,align,bord,patt,prot
        internal static BitField alignHor = bf(0x00000001); // 0 = Horizontal alignment modified
        internal static BitField alignVer = bf(0x00000002); // 0 = Vertical alignment modified
        internal static BitField alignWrap = bf(0x00000004); // 0 = Text wrapped flag modified
        internal static BitField alignRot = bf(0x00000008); // 0 = Text rotation modified
        internal static BitField alignJustLast = bf(0x00000010); // 0 = Justify last line flag modified
        internal static BitField alignIndent = bf(0x00000020); // 0 = Indentation modified
        internal static BitField alignShrin = bf(0x00000040); // 0 = Shrink to fit flag modified
        internal static BitField mergeCell = bf(0x00000080); // Normally 1, 0 = Merge Cell flag modified
        internal static BitField protLocked = bf(0x00000100); // 0 = Cell locked flag modified
        internal static BitField protHidden = bf(0x00000200); // 0 = Cell hidden flag modified
        internal static BitField bordLeft = bf(0x00000400); // 0 = Left border style and colour modified
        internal static BitField bordRight = bf(0x00000800); // 0 = Right border style and colour modified
        internal static BitField bordTop = bf(0x00001000); // 0 = Top border style and colour modified
        internal static BitField bordBot = bf(0x00002000); // 0 = Bottom border style and colour modified
        internal static BitField bordTlBr = bf(0x00004000); // 0 = Top-left to bottom-right border flag modified
        internal static BitField bordBlTr = bf(0x00008000); // 0 = Bottom-left to top-right border flag modified
        internal static BitField pattStyle = bf(0x00010000); // 0 = Pattern style modified
        internal static BitField pattCol = bf(0x00020000); // 0 = Pattern colour modified
        internal static BitField pattBgCol = bf(0x00040000); // 0 = Pattern background colour modified
        internal static BitField notUsed2 = bf(0x00380000); // Always 111 (ifmt / ifnt / 1)
        internal static BitField undocumented = bf(0x03C00000); // Undocumented bits
        internal static BitField fmtBlockBits = bf(0x7C000000); // Bits: font,align,bord,patt,prot
        internal static BitField font = bf(0x04000000); // 1 = Record Contains font formatting block
        internal static BitField align = bf(0x08000000); // 1 = Record Contains alignment formatting block
        internal static BitField bord = bf(0x10000000); // 1 = Record Contains border formatting block
        internal static BitField patt = bf(0x20000000); // 1 = Record Contains pattern formatting block
        internal static BitField prot = bf(0x40000000); // 1 = Record Contains protection formatting block
        internal static BitField alignTextDir = bf(0x80000000); // 0 = Text direction modified

        private static BitField bf(long i) {
            return BitFieldFactory.GetInstance((int)i);
        }

        protected int formatting_options;
        protected short formatting_not_used; // TODO Decode this properly

        protected FontFormatting _fontFormatting;
        protected BorderFormatting _borderFormatting;
        protected PatternFormatting _patternFormatting;

        private Formula formula1;
        private Formula formula2;

        /** Creates new CFRuleRecord */
        protected CFRuleBase(byte conditionType, byte comparisonOperation) {
            ConditionType = (conditionType);
            ComparisonOperation = (comparisonOperation);
            formula1 = Formula.Create(Ptg.EMPTY_PTG_ARRAY);
            formula2 = Formula.Create(Ptg.EMPTY_PTG_ARRAY);
        }
        protected CFRuleBase(byte conditionType, byte comparisonOperation, Ptg[] formula1, Ptg[] formula2)
                : this(conditionType, comparisonOperation)
        {

            this.formula1 = Formula.Create(formula1);
            this.formula2 = Formula.Create(formula2);
        }
        protected CFRuleBase() { }

        protected int ReadFormatOptions(RecordInputStream in1) {
            formatting_options = in1.ReadInt();
            formatting_not_used = in1.ReadShort();

            int len = 6;

            if (ContainsFontFormattingBlock) {
                _fontFormatting = new FontFormatting(in1);
                len += _fontFormatting.DataLength;
            }

            if (ContainsBorderFormattingBlock) {
                _borderFormatting = new BorderFormatting(in1);
                len += _borderFormatting.DataLength;
            }

            if (ContainsPatternFormattingBlock) {
                _patternFormatting = new PatternFormatting(in1);
                len += _patternFormatting.DataLength;
            }

            return len;
        }

        public byte ConditionType
        {
            get { return condition_type; }
            set
            {
                if ((this is CFRuleRecord))
                {
                    if (value == CONDITION_TYPE_CELL_VALUE_IS ||
                        value == CONDITION_TYPE_FORMULA)
                    {
                        // Good, valid combination
                    }
                    else
                    {
                        throw new ArgumentException("CFRuleRecord only accepts Value-Is and Formula types");
                    }
                }
                this.condition_type = value;
            }
        }

        public byte ComparisonOperation
        {
            get { return comparison_operator; }
            set
            {
                if (value < 0 || value > ComparisonOperator.max_operator)
                    throw new ArgumentException(
                            "Valid operators are only in the range 0 to " + ComparisonOperator.max_operator);

                this.comparison_operator = value;
            }
        }

        public bool ContainsFontFormattingBlock
        {
            get { return GetOptionFlag(font); }
        }

        public FontFormatting FontFormatting
        {
            get
            {
                if (ContainsFontFormattingBlock)
                {
                    return _fontFormatting;
                }
                return null;
            }
            set
            {
                _fontFormatting = value;
                SetOptionFlag(value != null, font);
            }
        }

        public bool ContainsAlignFormattingBlock() {
            return GetOptionFlag(align);
        }
        public void SetAlignFormattingUnChanged() {
            SetOptionFlag(false, align);
        }

        public bool ContainsBorderFormattingBlock
        {
            get { return GetOptionFlag(bord); }
            
        }

        public BorderFormatting BorderFormatting {
            get
            {
                if (ContainsBorderFormattingBlock)
                {
                    return _borderFormatting;
                }
                return null;
            }
            set
            {
                _borderFormatting = value;
                SetOptionFlag(value != null, bord);
            }
        }

        public bool ContainsPatternFormattingBlock
        {
            get { return GetOptionFlag(patt); }
        }

        public PatternFormatting PatternFormatting
        {
            get
            {
                if (ContainsPatternFormattingBlock)
                {
                    return _patternFormatting;
                }
                return null;
            }
            set
            {
                _patternFormatting = value;
                SetOptionFlag(value != null, patt);
            }
        }

        public bool ContainsProtectionFormattingBlock() {
            return GetOptionFlag(prot);
        }
        public void SetProtectionFormattingUnChanged() {
            SetOptionFlag(false, prot);
        }

        /**
         * Get the option flags
         *
         * @return bit mask
         */
        public int Options
        {
            get
            {
                return formatting_options;
            }
        }

        private bool IsModified(BitField field) {
            return !field.IsSet(formatting_options);
        }
        private void SetModified(bool modified, BitField field) {
            formatting_options = field.SetBoolean(formatting_options, !modified);
        }

        public bool IsLeftBorderModified
        {
            get
            {
                return IsModified(bordLeft);
            }
            set
            {
                SetModified(value, bordLeft);
            }
        }

        public bool IsRightBorderModified {
            get
            {
                return IsModified(bordRight);
            }
            set
            {
                SetModified(value, bordRight);
            }
        }

        public bool IsTopBorderModified
        {
            get
            {
                return IsModified(bordTop);
            }
            set
            {
                SetModified(value, bordTop);
            }
        }

        public bool IsBottomBorderModified
        {
            get
            {
                return IsModified(bordBot);
            }
            set
            {
                SetModified(value, bordBot);
            }
        }

        public bool IsTopLeftBottomRightBorderModified
        {
            get
            {
                return IsModified(bordTlBr);
            }
            set
            {
                SetModified(value, bordTlBr);
            }
        }

        public bool IsBottomLeftTopRightBorderModified
        {
            get
            {
                return IsModified(bordBlTr);
            }
            set
            {
                SetModified(value, bordBlTr);
            }
        }

        public bool IsPatternStyleModified
        {
            get
            {
                return IsModified(pattStyle);
            }
            set
            {
                SetModified(value, pattStyle);
            }
        }

        public bool IsPatternColorModified
        {
            get
            {
                return IsModified(pattCol);
            }
            set
            {
                SetModified(value, pattCol);
            }
        }

        public bool IsPatternBackgroundColorModified
        {
            get
            {
                return IsModified(pattBgCol);
            }
            set
            {
                SetModified(value, pattBgCol);
            }
        }

        private bool GetOptionFlag(BitField field) {
            return field.IsSet(formatting_options);
        }
        private void SetOptionFlag(bool flag, BitField field) {
            formatting_options = field.SetBoolean(formatting_options, flag);
        }

        protected int FormattingBlockSize
        {
            get
            {
                return 6 +
              (ContainsFontFormattingBlock ? _fontFormatting.RawRecord.Length : 0) +
              (ContainsBorderFormattingBlock ? 8 : 0) +
              (ContainsPatternFormattingBlock ? 4 : 0);
            }
        }
        protected void SerializeFormattingBlock(ILittleEndianOutput out1) {
            out1.WriteInt(formatting_options);
            out1.WriteShort(formatting_not_used);

            if (ContainsFontFormattingBlock) {
                byte[] fontFormattingRawRecord = _fontFormatting.RawRecord;
                out1.Write(fontFormattingRawRecord);
            }

            if (ContainsBorderFormattingBlock) {
                _borderFormatting.Serialize(out1);
            }

            if (ContainsPatternFormattingBlock) {
                _patternFormatting.Serialize(out1);
            }
        }

        /**
         * Get the stack of the 1st expression as a list
         *
         * @return list of tokens (casts stack to a list and returns it!)
         * this method can return null is we are unable to create Ptgs from
         *	 existing excel file
         * callers should check for null!
         */
        public Ptg[] ParsedExpression1
        {
            get
            {
                return formula1.Tokens;
            }
            set
            {
                formula1 = Formula.Create(value);
            }
        }

        protected Formula Formula1
        {
            get
            {
                return formula1;
            }
            set
            {
                this.formula1 = value;
            }
        }

        /**
         * Get the stack of the 2nd expression as a list
         *
         * @return array of {@link Ptg}s, possibly <code>null</code>
         */
        public Ptg[] ParsedExpression2
        {
            get
            {
                return formula2.Tokens;
            }
            set
            {
                formula2 = Formula.Create(value);
            }
        }

        protected Formula Formula2
        {
            get
            {
                return formula2;
            }
            set
            {
                this.formula2 = value;
            }
        }

        /**
         * @param formula must not be <code>null</code>
         * @return encoded size of the formula tokens (does not include 2 bytes for ushort length)
         */
        protected static int GetFormulaSize(Formula formula) {
            return formula.EncodedTokenSize;
        }

        /**
         * TODO - parse conditional format formulas properly i.e. produce tRefN and tAreaN instead of tRef and tArea
         * this call will produce the wrong results if the formula Contains any cell references
         * One approach might be to apply the inverse of SharedFormulaRecord.ConvertSharedFormulas(Stack, int, int)
         * Note - two extra parameters (rowIx & colIx) will be required. They probably come from one of the Region objects.
         *
         * @return <code>null</code> if <tt>formula</tt> was null.
         */
        public static Ptg[] ParseFormula(String formula, HSSFSheet sheet) {
            if (formula == null) {
                return null;
            }
            int sheetIndex = sheet.Workbook.GetSheetIndex(sheet);
            return HSSFFormulaParser.Parse(formula, sheet.Workbook as HSSFWorkbook, FormulaType.Cell, sheetIndex);
        }

        protected void CopyTo(CFRuleBase rec) {
            rec.condition_type = condition_type;
            rec.comparison_operator = comparison_operator;

            rec.formatting_options = formatting_options;
            rec.formatting_not_used = formatting_not_used;
            if (ContainsFontFormattingBlock) {
                rec._fontFormatting = (FontFormatting)_fontFormatting.Clone();
            }
            if (ContainsBorderFormattingBlock) {
                rec._borderFormatting = (BorderFormatting)_borderFormatting.Clone();
            }
            if (ContainsPatternFormattingBlock) {
                rec._patternFormatting = (PatternFormatting)_patternFormatting.Clone();
            }

            rec.formula1 = (formula1.Copy());
            rec.formula2 = (formula2.Copy());
        }
    }

}