/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed Under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

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
    using NPOI.SS.Util;

    using NPOI.Util;

    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula;
    


    /**
     * Title:        DATAVALIDATION Record (0x01BE)<p/>
     * Description:  This record stores data validation Settings and a list of cell ranges
     *               which contain these Settings. The data validation Settings of a sheet
     *               are stored in a sequential list of DV records. This list Is followed by
     *               DVAL record(s)
     * @author Dragos Buleandra (dragos.buleandra@trade2b.ro)
     * @version 2.0-pre
     */
    public class DVRecord : StandardRecord
    {
        private static readonly UnicodeString NULL_TEXT_STRING = new UnicodeString("\0");


        public const short sid = 0x01BE;
        /** Option flags */
        private int _option_flags;
        /** Title of the prompt box */
        private UnicodeString _promptTitle;
        /** Title of the error box */
        private UnicodeString _errorTitle;
        /** Text of the prompt box */
        private UnicodeString _promptText;
        /** Text of the error box */
        private UnicodeString _errorText;
        /** Not used - Excel seems to always write 0x3FE0 */
        private short _not_used_1 = 0x3FE0;
        /** Formula data for first condition (RPN token array without size field) */
        private NPOI.SS.Formula.Formula _formula1;
        /** Not used - Excel seems to always write 0x0000 */
        private short _not_used_2 = 0x0000;
        /** Formula data for second condition (RPN token array without size field) */
        private NPOI.SS.Formula.Formula _formula2;
        /** Cell range address list with all affected ranges */
        private CellRangeAddressList _regions;


        public const int STRING_PROMPT_TITLE = 0;
        public const int STRING_ERROR_TITLE = 1;
        public const int STRING_PROMPT_TEXT = 2;
        public const int STRING_ERROR_TEXT = 3;

        /**
         * Option flags field
         * @see org.apache.poi.hssf.util.HSSFDataValidation utility class
         */
        private BitField opt_data_type = new BitField(0x0000000F);
        private BitField opt_error_style = new BitField(0x00000070);
        private BitField opt_string_list_formula = new BitField(0x00000080);
        private BitField opt_empty_cell_allowed = new BitField(0x00000100);
        private BitField opt_suppress_dropdown_arrow = new BitField(0x00000200);
        private BitField opt_show_prompt_on_cell_selected = new BitField(0x00040000);
        private BitField opt_show_error_on_invalid_value = new BitField(0x00080000);
        private BitField opt_condition_operator = new BitField(0x00F00000);

        public DVRecord()
        {
        }
        public DVRecord(int validationType, int operator1, int errorStyle, bool emptyCellAllowed,
            bool suppressDropDownArrow, bool isExplicitList,
            bool showPromptBox, String promptTitle, String promptText,
            bool showErrorBox, String errorTitle, String errorText,
            Ptg[] formula1, Ptg[] formula2,
            CellRangeAddressList regions)
        {

            int flags = 0;
            flags = opt_data_type.SetValue(flags, validationType);
            flags = opt_condition_operator.SetValue(flags, operator1);
            flags = opt_error_style.SetValue(flags, errorStyle);
            flags = opt_empty_cell_allowed.SetBoolean(flags, emptyCellAllowed);
            flags = opt_suppress_dropdown_arrow.SetBoolean(flags, suppressDropDownArrow);
            flags = opt_string_list_formula.SetBoolean(flags, isExplicitList);
            flags = opt_show_prompt_on_cell_selected.SetBoolean(flags, showPromptBox);
            flags = opt_show_error_on_invalid_value.SetBoolean(flags, showErrorBox);
            _option_flags = flags;
            _promptTitle = ResolveTitleText(promptTitle);
            _promptText = ResolveTitleText(promptText);
            _errorTitle = ResolveTitleText(errorTitle);
            _errorText = ResolveTitleText(errorText);
            _formula1 = NPOI.SS.Formula.Formula.Create(formula1);
            _formula2 = NPOI.SS.Formula.Formula.Create(formula2);
            _regions = regions;
        }

        /**
         * Constructs a DV record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public DVRecord(RecordInputStream in1)
        {
            _option_flags = in1.ReadInt();

            _promptTitle = ReadUnicodeString(in1);
            _errorTitle = ReadUnicodeString(in1);
            _promptText = ReadUnicodeString(in1);
            _errorText = ReadUnicodeString(in1);

            int field_size_first_formula = in1.ReadUShort();
            _not_used_1 = in1.ReadShort();

            //read first formula data condition
            _formula1 = NPOI.SS.Formula.Formula.Read(field_size_first_formula, in1);

            int field_size_sec_formula = in1.ReadUShort();
            _not_used_2 = in1.ReadShort();

            //read sec formula data condition
            _formula2 = NPOI.SS.Formula.Formula.Read(field_size_sec_formula, in1);

            //read cell range address list with all affected ranges
            _regions = new CellRangeAddressList(in1);
        }
        /**
         * When entered via the UI, Excel translates empty string into "\0"
         * While it is possible to encode the title/text as empty string (Excel doesn't exactly crash),
         * the resulting tool-tip text / message box looks wrong.  It is best to do the same as the 
         * Excel UI and encode 'not present' as "\0". 
         */
        private static UnicodeString ResolveTitleText(String str)
        {
            if (str == null || str.Length < 1)
            {
                return NULL_TEXT_STRING;
            }
            return new UnicodeString(str);
        }

        private static String ResolveTitleString(UnicodeString us)
        {
            if (us == null || us.Equals(NULL_TEXT_STRING))
            {
                return null;
            }
            return us.String;
        }

        private static UnicodeString ReadUnicodeString(RecordInputStream in1)
        {
            return new UnicodeString(in1);
        }
        /**
         * Get the condition data type
         * @return the condition data type
         * @see org.apache.poi.hssf.util.HSSFDataValidation utility class
         */
        public int DataType
        {
            get
            {
                return this.opt_data_type.GetValue(this._option_flags);
            }
            set { this._option_flags = this.opt_data_type.SetValue(this._option_flags, value); }
        }



        /**
         * Get the condition error style
         * @return the condition error style
         * @see org.apache.poi.hssf.util.HSSFDataValidation utility class
         */
        public int ErrorStyle
        {
            get
            {
                return this.opt_error_style.GetValue(this._option_flags);
            }
            set { this._option_flags = this.opt_error_style.SetValue(this._option_flags, value); }
        }


        /**
         * return true if in list validations the string list Is explicitly given in the formula, false otherwise
         * @return true if in list validations the string list Is explicitly given in the formula, false otherwise
         * @see org.apache.poi.hssf.util.HSSFDataValidation utility class
         */
        public bool ListExplicitFormula
        {
            get
            {
                return (this.opt_string_list_formula.IsSet(this._option_flags));
            }
            set { this._option_flags = this.opt_string_list_formula.SetBoolean(this._option_flags, value); }
        }



        /**
         * return true if empty values are allowed in cells, false otherwise
         * @return if empty values are allowed in cells, false otherwise
         * @see org.apache.poi.hssf.util.HSSFDataValidation utility class
         */
        public bool EmptyCellAllowed
        {
            get
            {
                return (this.opt_empty_cell_allowed.IsSet(this._option_flags));
            }
            set { this._option_flags = this.opt_empty_cell_allowed.SetBoolean(this._option_flags, value); }
        }

        /**
          * @return <code>true</code> if drop down arrow should be suppressed when list validation is
          * used, <code>false</code> otherwise
         */
        public bool SuppressDropdownArrow
        {
            get
            {
                return (opt_suppress_dropdown_arrow.IsSet(_option_flags));
            }
        }
        /**
         * return true if a prompt window should appear when cell Is selected, false otherwise
         * @return if a prompt window should appear when cell Is selected, false otherwise
         * @see org.apache.poi.hssf.util.HSSFDataValidation utility class
         */
        public bool ShowPromptOnCellSelected
        {
            get
            {
                return (this.opt_show_prompt_on_cell_selected.IsSet(this._option_flags));
            }
        }


        /**
         * return true if an error window should appear when an invalid value Is entered in the cell, false otherwise
         * @return if an error window should appear when an invalid value Is entered in the cell, false otherwise
         * @see org.apache.poi.hssf.util.HSSFDataValidation utility class
         */
        public bool ShowErrorOnInvalidValue
        {
            get
            {
                return (this.opt_show_error_on_invalid_value.IsSet(this._option_flags));
            }
            set { this._option_flags = this.opt_show_error_on_invalid_value.SetBoolean(this._option_flags, value); }
        }



        /**
         * Get the condition operator
         * @return the condition operator
         * @see org.apache.poi.hssf.util.HSSFDataValidation utility class
         */
        public int ConditionOperator
        {
            get
            {
                return this.opt_condition_operator.GetValue(this._option_flags);
            }
            set
            {
                this._option_flags = this.opt_condition_operator.SetValue(this._option_flags, value);
            }
        }

        public String PromptTitle
        {
            get
            {
                return ResolveTitleString(_promptTitle);
            }
        }

        public String ErrorTitle
        {
            get
            {
                return ResolveTitleString(_errorTitle);
            }
        }

        public String PromptText
        {
            get
            {
                return ResolveTitleString(_promptText);
            }
        }

        public String ErrorText
        {
            get
            {
                return ResolveTitleString(_errorText);
            }
        }

        public Ptg[] Formula1
        {
            get
            {
                return Formula.GetTokens(_formula1);
            }
        }

        public Ptg[] Formula2
        {
            get
            {
                return  Formula.GetTokens(_formula2);
            }
        }

        public CellRangeAddressList CellRangeAddress
        {
            get
            {
                return this._regions;
            }
            set { this._regions = value; }
        }

        /**
         * Gets the option flags field.
         * @return options - the option flags field
         */
        public int OptionFlags
        {
            get
            {
                return this._option_flags;
            }
        }

        public override String ToString()
        {
            /* @todo DVRecord string representation */
            StringBuilder buffer = new StringBuilder();

            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {

            out1.WriteInt(_option_flags);

            SerializeUnicodeString(_promptTitle, out1);
            SerializeUnicodeString(_errorTitle, out1);
            SerializeUnicodeString(_promptText, out1);
            SerializeUnicodeString(_errorText, out1);
            out1.WriteShort(_formula1.EncodedTokenSize);
            out1.WriteShort(_not_used_1);
            _formula1.SerializeTokens(out1);

            out1.WriteShort(_formula2.EncodedTokenSize);
            out1.WriteShort(_not_used_2);
            _formula2.SerializeTokens(out1);

            _regions.Serialize(out1);
        }
        private static void SerializeUnicodeString(UnicodeString us, ILittleEndianOutput out1)
        {
            StringUtil.WriteUnicodeString(out1, us.String);
        }

        private static int GetUnicodeStringSize(UnicodeString us)
        {
            String str = us.String;
            return 3 + str.Length * (StringUtil.HasMultibyte(str) ? 2 : 1);
        }
        protected override int DataSize
        {
            get
            {
                int size = 4 + 2 + 2 + 2 + 2;//header+options_field+first_formula_size+first_unused+sec_formula_size+sec+unused;
                size += GetUnicodeStringSize(_promptTitle);
                size += GetUnicodeStringSize(_errorTitle);
                size += GetUnicodeStringSize(_promptText);
                size += GetUnicodeStringSize(_errorText);
                size += _formula1.EncodedTokenSize;
                size += _formula2.EncodedTokenSize;
                size += _regions.Size;
                return size;
            }
        }

        public override short Sid
        {
            get { return DVRecord.sid; }
        }

        /**
         * Clones the object. Uses serialisation, as the
         *  contents are somewhat complex
         */
        public override Object Clone()
        {
            return CloneViaReserialise();
        }
    }
}