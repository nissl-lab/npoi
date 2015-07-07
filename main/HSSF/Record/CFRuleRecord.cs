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

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.HSSF.Record.CF;
    using FR=NPOI.SS.Formula;

    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;

    //public enum ComparisonOperator : byte
    //{
    //    NO_COMPARISON = 0,
    //    BETWEEN = 1,
    //    NOT_BETWEEN = 2,
    //    EQUAL = 3,
    //    NOT_EQUAL = 4,
    //    GT = 5,
    //    LT = 6,
    //    GE = 7,
    //    LE = 8
    //}
    /**
     * Conditional Formatting Rule Record.
     * @author Dmitriy Kumshayev
     */
    public class CFRuleRecord : StandardRecord
    {

        public const short sid = 0x01B1;



        private byte field_1_condition_type;
        public const byte CONDITION_TYPE_CELL_VALUE_IS = 1;
        public const byte CONDITION_TYPE_FORMULA = 2;

        private byte field_2_comparison_operator;

        private int field_5_options;

        private static BitField modificationBits = bf(0x003FFFFF); // Bits: font,align,bord,patt,prot
        private static BitField alignHor = bf(0x00000001); // 0 = Horizontal alignment modified
        private static BitField alignVer = bf(0x00000002); // 0 = Vertical alignment modified
        private static BitField alignWrap = bf(0x00000004); // 0 = Text wrapped flag modified
        private static BitField alignRot = bf(0x00000008); // 0 = Text rotation modified
        private static BitField alignJustLast = bf(0x00000010); // 0 = Justify last line flag modified
        private static BitField alignIndent = bf(0x00000020); // 0 = Indentation modified
        private static BitField alignShrin = bf(0x00000040); // 0 = Shrink to fit flag modified
        private static BitField notUsed1 = bf(0x00000080); // Always 1
        private static BitField protLocked = bf(0x00000100); // 0 = Cell locked flag modified
        private static BitField protHidden = bf(0x00000200); // 0 = Cell hidden flag modified
        private static BitField bordLeft = bf(0x00000400); // 0 = Left border style and colour modified
        private static BitField bordRight = bf(0x00000800); // 0 = Right border style and colour modified
        private static BitField bordTop = bf(0x00001000); // 0 = Top border style and colour modified
        private static BitField bordBot = bf(0x00002000); // 0 = Bottom border style and colour modified
        private static BitField bordTlBr = bf(0x00004000); // 0 = Top-left to bottom-right border flag modified
        private static BitField bordBlTr = bf(0x00008000); // 0 = Bottom-left to top-right border flag modified
        private static BitField pattStyle = bf(0x00010000); // 0 = Pattern style modified
        private static BitField pattCol = bf(0x00020000); // 0 = Pattern colour modified
        private static BitField pattBgCol = bf(0x00040000); // 0 = Pattern background colour modified
        private static BitField notUsed2 = bf(0x00380000); // Always 111
        private static BitField Undocumented = bf(0x03C00000); // Undocumented bits
        private static BitField fmtBlockBits = bf(0x7C000000); // Bits: font,align,bord,patt,prot
        private static BitField font = bf(0x04000000); // 1 = Record Contains font formatting block
        private static BitField align = bf(0x08000000); // 1 = Record Contains alignment formatting block
        private static BitField bord = bf(0x10000000); // 1 = Record Contains border formatting block
        private static BitField patt = bf(0x20000000); // 1 = Record Contains pattern formatting block
        private static BitField prot = bf(0x40000000); // 1 = Record Contains protection formatting block
        private static BitField alignTextDir = bf(unchecked((int)0x80000000)); // 0 = Text direction modified


        private static BitField bf(int i)
        {
            return BitFieldFactory.GetInstance(i);
        }
        private short field_6_not_used;

        private FontFormatting _fontFormatting;

        // fix warning CS0414 "never used": private byte field_8_align_text_break;
        // fix warning CS0414 "never used": private byte field_9_align_text_rotation_angle;
        // fix warning CS0414 "never used": private short field_10_align_indentation;
        // fix warning CS0414 "never used": private short field_11_relative_indentation;
        // fix warning CS0414 "never used": private short field_12_not_used;

        private BorderFormatting _borderFormatting;

        private PatternFormatting _patternFormatting;

        private FR.Formula field_17_formula1;
        private FR.Formula field_18_formula2;

        /** Creates new CFRuleRecord */
        private CFRuleRecord(byte conditionType, ComparisonOperator comparisonOperation)
        {
            field_1_condition_type = conditionType;
            field_2_comparison_operator =(byte) comparisonOperation;

            // Set modification flags to 1: by default options are not modified
            field_5_options = modificationBits.SetValue(field_5_options, -1);
            // Set formatting block flags to 0 (no formatting blocks)
            field_5_options = fmtBlockBits.SetValue(field_5_options, 0);
            field_5_options = Undocumented.Clear(field_5_options);

            //TODO:: check what's this field used for
            field_6_not_used = unchecked((short)0x8002); // Excel seems to Write this value, but it doesn't seem to care what it Reads
            _fontFormatting = null;
            //field_8_align_text_break = 0;
            //field_9_align_text_rotation_angle = 0;
            //field_10_align_indentation = 0;
            //field_11_relative_indentation = 0;
            //field_12_not_used = 0;
            _borderFormatting = null;
            _patternFormatting = null;
            field_17_formula1 = FR.Formula.Create(Ptg.EMPTY_PTG_ARRAY);
            field_18_formula2 = FR.Formula.Create(Ptg.EMPTY_PTG_ARRAY);
        }

        private CFRuleRecord(byte conditionType, ComparisonOperator comparisonOperation, Ptg[] formula1, Ptg[] formula2)
            :this(conditionType, comparisonOperation)
        {
            
            //field_1_condition_type = CONDITION_TYPE_CELL_VALUE_IS;
            //field_2_comparison_operator = (byte)comparisonOperation;
            field_17_formula1 = FR.Formula.Create(formula1);
            field_18_formula2 = FR.Formula.Create(formula2);
        }
        /**
         * get the stack of the 1st expression as a list
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
                return field_17_formula1.Tokens;
            }
            set { field_17_formula1 = FR.Formula.Create(value); }
        }
        /**
         * get the stack of the 2nd expression as a list
         *
         * @return list of tokens (casts stack to a list and returns it!)
         * this method can return null is we are unable to create Ptgs from 
         *	 existing excel file
         * callers should check for null!
         */

        public Ptg[] ParsedExpression2
        {
            get
            {
                return field_18_formula2.Tokens;
            }
            set { field_18_formula2 = FR.Formula.Create(value); }
        }

        /**
         * Creates a new comparison operation rule
         */
        [Obsolete]
        public static CFRuleRecord Create(HSSFWorkbook workbook, String formulaText)
        {
            Ptg[] formula1 = ParseFormula(formulaText, workbook);
            return new CFRuleRecord(CONDITION_TYPE_FORMULA, ComparisonOperator.NoComparison,
                    formula1, null);
        }
        /**
         * Creates a new comparison operation rule
         */
        [Obsolete]
        public static CFRuleRecord Create(HSSFWorkbook workbook, ComparisonOperator comparisonOperation,
                String formulaText1, String formulaText2)
        {
            Ptg[] formula1 = ParseFormula(formulaText1, workbook);
            Ptg[] formula2 = ParseFormula(formulaText2, workbook);
            return new CFRuleRecord(CONDITION_TYPE_CELL_VALUE_IS, comparisonOperation, formula1, formula2);
        }
        public static CFRuleRecord Create(HSSFSheet sheet, String formulaText)
        {
            Ptg[] formula1 = ParseFormula(formulaText, sheet);
            return new CFRuleRecord(CONDITION_TYPE_FORMULA, ComparisonOperator.NoComparison,
                    formula1, null);
        }
        /**
         * Creates a new comparison operation rule
         */
        public static CFRuleRecord Create(HSSFSheet sheet, byte comparisonOperation,
                String formulaText1, String formulaText2)
        {
            Ptg[] formula1 = ParseFormula(formulaText1, sheet);
            Ptg[] formula2 = ParseFormula(formulaText2, sheet);
            return new CFRuleRecord(CONDITION_TYPE_CELL_VALUE_IS, (ComparisonOperator)comparisonOperation, formula1, formula2);
        }
        public CFRuleRecord(RecordInputStream in1)
        {
            field_1_condition_type = (byte)in1.ReadByte();
            field_2_comparison_operator = (byte)in1.ReadByte();
            int field_3_formula1_len = in1.ReadUShort();
            int field_4_formula2_len = in1.ReadUShort();
            field_5_options = in1.ReadInt();
            field_6_not_used = in1.ReadShort();

            if (ContainsFontFormattingBlock)
            {
                _fontFormatting = new FontFormatting(in1);
            }

            if (ContainsBorderFormattingBlock)
            {
                _borderFormatting = new BorderFormatting(in1);
            }

            if (ContainsPatternFormattingBlock)
            {
                _patternFormatting = new PatternFormatting(in1);
            }
            field_17_formula1 = FR.Formula.Read(field_3_formula1_len, in1);
            field_18_formula2 = FR.Formula.Read(field_4_formula2_len, in1);
    }

        public byte ConditionType
        {
            get { return field_1_condition_type; }
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
                else
                {
                    return null;
                }
            }
            set
            {
                this._fontFormatting = value;
                SetOptionFlag(_fontFormatting != null, font);
            }
        }

        public bool ContainsAlignFormattingBlock
        {
            get { return GetOptionFlag(align); }
        }
        public void SetAlignFormattingUnChanged()
        {
            SetOptionFlag(false, align);
        }

        public bool ContainsBorderFormattingBlock
        {
            get { return GetOptionFlag(bord); }
        }
        public BorderFormatting BorderFormatting
        {
            get
            {
                if (ContainsBorderFormattingBlock)
                {
                    return _borderFormatting;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                this._borderFormatting = value;
                SetOptionFlag(_borderFormatting != null, bord);
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
                else
                {
                    return null;
                }
            }
            set 
            {
                this._patternFormatting = value;
                SetOptionFlag(_patternFormatting != null, patt);
            }
        }

        public bool ContainsProtectionFormattingBlock
        {
            get { return GetOptionFlag(prot); }
        }
        public void SetProtectionFormattingUnChanged()
        {
            SetOptionFlag(false, prot);
        }

        public byte ComparisonOperation
        {
            get { return field_2_comparison_operator; }
            set { field_2_comparison_operator = value; }
        }


        /**
         * Get the option flags
         *
         * @return bit mask
         */
        public int Options
        {
            get { return field_5_options; }
        }

        private bool IsModified(BitField field)
        {
            return !field.IsSet(field_5_options);
        }

        private void SetModified(bool modified, BitField field)
        {
            field_5_options = field.SetBoolean(field_5_options, !modified);
        }

        public bool IsLeftBorderModified
        {
            get { return IsModified(bordLeft); }
            set { SetModified(value, bordLeft);}
        }

        public bool IsRightBorderModified
        {
            get { return IsModified(bordRight); }
            set { SetModified(value, bordRight); }
        }

        public bool IsTopBorderModified
        {
            get { return IsModified(bordTop); }
            set { SetModified(value, bordTop);}
        }

        public bool IsBottomBorderModified
        {
            get { return IsModified(bordBot); }
            set { SetModified(value, bordBot); }
        }

        public bool IsTopLeftBottomRightBorderModified
        {
            get{return IsModified(bordTlBr);}
            set {  SetModified(value, bordTlBr);}
        }

        public bool IsBottomLeftTopRightBorderModified
        {
            get { return IsModified(bordBlTr); }
            set { SetModified(value, bordBlTr);}
        }

        public bool IsPatternStyleModified
        {
            get { return IsModified(pattStyle); }
            set { SetModified(value, pattStyle);}
        }

        public bool IsPatternColorModified
        {
            get { return IsModified(pattCol); }
            set { SetModified(value, pattCol);}
        }

        public bool IsPatternBackgroundColorModified
        {
            get { return IsModified(pattBgCol); }
            set {SetModified(value, pattBgCol); }
        }

        private bool GetOptionFlag(BitField field)
        {
            return field.IsSet(field_5_options);
        }

        private void SetOptionFlag(bool flag, BitField field)
        {
            field_5_options = field.SetBoolean(field_5_options, flag);
        }



        public override short Sid
        {
            get { return sid; }
        }

        /**
         * @param ptgs may be <c>null</c>
         * @return encoded size of the formula
         */
        private int GetFormulaSize(FR.Formula formula)
        {
            return formula.EncodedTokenSize;
        }

        /**
         * called by the class that Is responsible for writing this sucker.
         * Subclasses should implement this so that their data Is passed back in a
         * byte array.
         *
         * @param offset to begin writing at
         * @param data byte array containing instance data
         * @return number of bytes written
         */
        public override void Serialize(ILittleEndianOutput out1)
        {
            int formula1Len=GetFormulaSize(field_17_formula1);
            int formula2Len=GetFormulaSize(field_18_formula2);

            out1.WriteByte(field_1_condition_type);
            out1.WriteByte(field_2_comparison_operator);
            out1.WriteShort(formula1Len);
            out1.WriteShort(formula2Len);
            out1.WriteInt(field_5_options);
            out1.WriteShort(field_6_not_used);

            if (ContainsFontFormattingBlock) {
                byte[] fontFormattingRawRecord  = _fontFormatting.GetRawRecord();
                out1.Write(fontFormattingRawRecord);
            }

            if (ContainsBorderFormattingBlock) {
                _borderFormatting.Serialize(out1);
            }

            if (ContainsPatternFormattingBlock) {
                _patternFormatting.Serialize(out1);
            }

            field_17_formula1.SerializeTokens(out1);
            field_18_formula2.SerializeTokens(out1);
        }


        protected override int DataSize
        {
            get
            {
                int retval = 12 +
                            (ContainsFontFormattingBlock ? _fontFormatting.GetRawRecord().Length : 0) +
                            (ContainsBorderFormattingBlock ? 8 : 0) +
                            (ContainsPatternFormattingBlock? 4 : 0) +
                            GetFormulaSize(field_17_formula1) +
                            GetFormulaSize(field_18_formula2)
                            ;
                return retval;
            }
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[CFRULE]\n");
            buffer.Append("    .condition_type   =").Append(field_1_condition_type).Append("\n");
            buffer.Append("    OPTION FLAGS=0x").Append(string.Format("{0:X}",Options)).Append("\n");
            if (ContainsFontFormattingBlock)
            {
                buffer.Append(_fontFormatting.ToString()).Append("\n");
            }
            if (ContainsBorderFormattingBlock)
            {
                buffer.Append(_borderFormatting.ToString()).Append("\n");
            }
            if (ContainsPatternFormattingBlock)
            {
                buffer.Append(_patternFormatting.ToString()).Append("\n");
            }
            buffer.Append("    Formula 1 =").Append(Arrays.ToString(field_17_formula1.Tokens)).Append("\n");
            buffer.Append("    Formula 2 =").Append(Arrays.ToString(field_18_formula2.Tokens)).Append("\n");
            buffer.Append("[/CFRULE]\n");
            return buffer.ToString();
        }
 

        public override Object Clone()
        {
            CFRuleRecord rec = new CFRuleRecord(field_1_condition_type, (ComparisonOperator)field_2_comparison_operator);
            rec.field_5_options = field_5_options;
            rec.field_6_not_used = field_6_not_used;
            if (ContainsFontFormattingBlock)
            {
                rec._fontFormatting = (FontFormatting)_fontFormatting.Clone();
            }
            if (ContainsBorderFormattingBlock)
            {
                rec._borderFormatting = (BorderFormatting)_borderFormatting.Clone();
            }
            if (ContainsPatternFormattingBlock)
            {
                rec._patternFormatting = (PatternFormatting)_patternFormatting.Clone();
            }
            if (field_17_formula1 != null)
            {
                rec.field_17_formula1 = field_17_formula1.Copy();
            }
            if (field_18_formula2 != null)
            {
                rec.field_18_formula2 = field_18_formula2.Copy();
            }

            return rec;
        }

        /**
         * TODO - Parse conditional format formulas properly i.e. produce tRefN and tAreaN instead of tRef and tArea
         * this call will produce the wrong results if the formula Contains any cell references
         * One approach might be to apply the inverse of SharedFormulaRecord.ConvertSharedFormulas(Stack, int, int)
         * Note - two extra parameters (rowIx &amp;colIx) will be required. They probably come from one of the Region objects.
         * 
         * @return <c>null</c> if <c>formula</c> was null.
         */
        private static Ptg[] ParseFormula(String formula, HSSFWorkbook workbook)
        {
            if (formula == null)
            {
                return null;
            }
            return HSSFFormulaParser.Parse(formula, workbook);
        }
        /**
     * TODO - parse conditional format formulas properly i.e. produce tRefN and tAreaN instead of tRef and tArea
     * this call will produce the wrong results if the formula contains any cell references
     * One approach might be to apply the inverse of SharedFormulaRecord.convertSharedFormulas(Stack, int, int)
     * Note - two extra parameters (rowIx &amp; colIx) will be required. They probably come from one of the Region objects.
     *
     * @return <code>null</code> if <c>formula</c> was null.
     */
        private static Ptg[] ParseFormula(String formula, HSSFSheet sheet)
        {
            if (formula == null)
            {
                return null;
            }
            int sheetIndex = sheet.Workbook.GetSheetIndex(sheet);
            return HSSFFormulaParser.Parse(formula, (HSSFWorkbook)sheet.Workbook, FormulaType.Cell, sheetIndex);
        }
    }
}