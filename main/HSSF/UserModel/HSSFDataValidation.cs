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

namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.SS.Util;
    /**
     * Title: HSSFDataValidation
     * Description: Utilty class for creating data validation cells
     * Copyright: Copyright (c) 2004
     * Company: 
     * @author Dragos Buleandra (dragos.buleandra@trade2b.ro)
     * @version 2.0-pre
     */

    public class HSSFDataValidation
    {
        /**
 * Error style constants for error box
 */
        public enum ERRORSTYLE:int
        {
            /** STOP style */
            STOP = 0x00,
            /** WARNING style */
            WARNING = 0x01,
            /** INFO style */
            INFO = 0x02
        }
        /// <summary>
        ///  Validation data type constants
        /// </summary>
        public enum DataType : int
        { 
            ANY=0x00,
            INTEGER=0X01,
            DECIMAL=0X02,
            /// <summary>
            /// List type ( combo box type )
            /// </summary>
            LIST=0X03,
            DATE=0X04,
            TIME=0X05,
            /// <summary>
            /// String Length type
            /// </summary>
            TEXT_LENGTH=0X06,
            /// <summary>
            /// Formula ( custom ) type
            /// </summary>
            FORMULA=0X07
        }
        /// <summary>
        /// Condition operator
        /// </summary>
        public enum OperatorType : int
        { 
            BETWEEN = 0x00,
            NOT_BETWEEN = 0x01,
            EQUAL = 0x02,
            NOT_EQUAL = 0x03,
            GREATER_THAN = 0x04,
            LESS_THAN = 0x05,
            GREATER_OR_EQUAL = 0x06,
            LESS_OR_EQUAL = 0x07        
        }

        private String _prompt_title = null;
        private String _prompt_text = null;
        private String _error_title = null;
        private String _error_text = null;
        private String _string_first_formula = null;
        private String _string_sec_formula = null;

        private DataType _data_type = DataType.ANY;
        private ERRORSTYLE _error_style = ERRORSTYLE.STOP;
        private bool _list_explicit_formula = true;
        private bool _empty_cell_allowed = true;
        private bool _surpress_dropdown_arrow = false;
        private bool _show_prompt_box = true;
        private bool _show_error_box = true;
        private OperatorType _operator = OperatorType.BETWEEN;
        private DVConstraint _constraint;


        private CellRangeAddressList _regions;

        /**
         * Empty constructor
         */
        public HSSFDataValidation()
        {
        }

        /**
 * Constructor which initializes the cell range on which this object will be
 * applied
 * @param constraint 
 */
        public HSSFDataValidation(CellRangeAddressList regions, DVConstraint constraint)
        {
            _regions = regions;
            _constraint = constraint;
        }
        public DVRecord CreateDVRecord(HSSFWorkbook workbook)
        {

            FormulaPair fp = _constraint.CreateFormulas(workbook);

            return new DVRecord(_constraint.GetValidationType(),
                    _constraint.Operator,
                    (int)_error_style, _empty_cell_allowed, SuppressDropDownArrow,
                    _constraint.IsExplicitList,
                    this._show_prompt_box, _prompt_title, _prompt_text,
                    this._show_error_box, _error_title, _error_text,
                    fp.Formula1, fp.Formula2,
                    _regions);
        }
        /**
         * Set the type of this object
         * @param data_type The type
         * @see DATA_TYPE_ANY, DATA_TYPE_INTEGER, DATA_TYPE_DECIMNAL, DATA_TYPE_LIST, DATA_TYPE_DATE,
         *      DATA_TYPE_TIME, DATA_TYPE_TEXT_LENTGH, DATA_TYPE_FORMULA
         */
        public DataType DataValidationType
        {
            set { this._data_type = value; }
            get { return this._data_type; }
        }

        /**
         * Sets the error style for error box
         * @param error_style Error style constant
         * @see ERROR_STYLE_STOP, ERROR_STYLE_WARNING, ERROR_STYLE_INFO
         */
        public ERRORSTYLE ErrorStyle
        {
            set { this._error_style = value; }
            get { return _error_style; }
        }

        /**
         * If this object has an explicit formula . This is useful only for list data validation object
         * @param explicit True if use an explicit formula
         */
        public bool ExplicitListFormula
        {
            set { this._list_explicit_formula = value; }
            get
            {
                if (this._data_type != DataType.LIST)
                {
                    return false;
                }
                return this._list_explicit_formula;
            }
        }


        /**
         * Sets if this object allows empty as a valid value
         * @param allowed True if this object should treats empty as valid value , false otherwise
         */
        public bool EmptyCellAllowed
        {
            set { this._empty_cell_allowed = value; }
            get { return this._empty_cell_allowed; }
        }

        /**
         * Useful only list validation objects .
         * This method always returns false if the object Isn't a list validation object
         * @return True if a list should Display the values into a drop down list , false otherwise .
         * @see SetDataValidationType( int data_type )
         */
        public bool SuppressDropDownArrow
        {
            get
            {
                if (this._data_type != DataType.LIST)
                {
                    return false;
                }
                return this._surpress_dropdown_arrow;
            }
            set
            {
                this._surpress_dropdown_arrow = value;
            }
        }

        /**
         * @param show True if an prompt box should be Displayed , false otherwise
         */
        public bool ShowPromptBox
        {
            get
            {
                if ((this.PromptBoxText == null) && (this.PromptBoxTitle == null))
                {
                    return false;
                }
                return this._show_prompt_box;
            }
            set
            {
                this._show_prompt_box = value;
            }
        }

        /**
         * @return True if an error box should be Displayed , false otherwise
         */
        public bool ShowErrorBox
        {
            get
            {
                if ((this.ErrorBoxText == null) && (this.ErrorBoxTitle == null))
                {
                    return false;
                }
                return this._show_error_box;
            }
            set
            {
                this._show_error_box = value;
            }
        }

        /**
         * Sets the operator involved in the formula whic governs this object
         * Example : if you wants that a cell to accept only values between 1 and 5 , which
         * mathematically means 1 <= value <= 5 , then the operator should be OPERATOR_BETWEEN
         * @param operator A constant for operator
         * @see OPERATOR_BETWEEN, OPERATOR_NOT_BETWEEN, OPERATOR_EQUAL, OPERATOR_NOT_EQUAL
         * OPERATOR_GREATER_THAN, OPERATOR_LESS_THAN, OPERATOR_GREATER_OR_EQUAL,
         * OPERATOR_LESS_OR_EQUAL
         */
        
        public OperatorType Operator
        {
            set { this._operator = value; }
            get { return this._operator; }
        }

        /**
         * Sets the title and text for the prompt box . Prompt box is Displayed when the user
         * selects a cell which belongs to this validation object . In order for a prompt box
         * to be Displayed you should also use method SetShowPromptBox( bool show )
         * @param title The prompt box's title
         * @param text The prompt box's text
         * @see SetShowPromptBox( bool show )
         */
        public void CreatePromptBox(String title, String text)
        {
            this._prompt_title = title;
            this._prompt_text = text;
            this.ShowPromptBox = true;
        }

        /**
         * Returns the prompt box's title
         * @return Prompt box's title or null
         */
        public String PromptBoxTitle
        {
            get { return this._prompt_title; }
        }

        /**
         * Returns the prompt box's text
         * @return Prompt box's text or null
         */
        public String PromptBoxText
        {
            get { return this._prompt_text; }
        }

        /**
         * Sets the title and text for the error box . Error box is Displayed when the user
         * enters an invalid value int o a cell which belongs to this validation object .
         * In order for an error box to be Displayed you should also use method
         * SetShowErrorBox( bool show )
         * @param title The error box's title
         * @param text The error box's text
         * @see SetShowErrorBox( bool show )
         */
        public void CreateErrorBox(String title, String text)
        {
            this._error_title = title;
            this._error_text = text;
            this.ShowErrorBox = true;
        }

        /**
         * Returns the error box's title
         * @return Error box's title or null
         */
        public String ErrorBoxTitle
        {
            get { return this._error_title; }
        }

        /**
         * Returns the error box's text
         * @return Error box's text or null
         */
        public String ErrorBoxText
        {
            get { return this._error_text; }
        }

        /**
         * Sets the first formula for this object .
         * A formula is divided into three parts : first formula , operator and second formula .
         * In other words , a formula Contains a left oprand , an operator and a right operand.
         * This is the general rule . An example is 1<= value <= 5 . In this case ,
         * the left operand ( or the first formula ) is the number 1 . The operator Is
         * OPERATOR_BETWEEN and the right operand ( or the second formula ) is 5 .
         * @param formula
         */
        public string FirstFormula
        {
            set { this._string_first_formula = value; }
            get { return this._string_first_formula; }
        }

        /**
         * Sets the first formula for this object .
         * A formula is divided into three parts : first formula , operator and second formula .
         * In other words , a formula Contains a left oprand , an operator and a right operand.
         * This is the general rule . An example is 1<= value <=5 . In this case ,
         * the left operand ( or the first formula ) is the number 1 . The operator Is
         * OPERATOR_BETWEEN and the right operand ( or the second formula ) is 5 .
         * But there are cases when a second formula Isn't needed :
         * You want somethink like : all values less than 5 . In this case , there's only a first
         * formula ( in our case 5 ) and the operator OPERATOR_LESS_THAN
         * @param formula
         */
        public string SecondFormula
        {
            set { this._string_sec_formula = value; }
            get { return this._string_sec_formula; }
        }

    }
}