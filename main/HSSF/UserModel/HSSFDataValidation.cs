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

namespace NPOI.HSSF.UserModel
{
    using System;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     *Utility class for creating data validation cells
     * 
     * @author Dragos Buleandra (dragos.buleandra@trade2b.ro)
     */
    public class HSSFDataValidation : IDataValidation
    {
        private String _prompt_title;
        private String _prompt_text;
        private String _error_title;
        private String _error_text;

        private int _errorStyle = ERRORSTYLE.STOP;
        private bool _emptyCellAllowed = true;
        private bool _suppress_dropdown_arrow = false;
        private bool _ShowPromptBox = true;
        private bool _ShowErrorBox = true;
        private CellRangeAddressList _regions;
        private DVConstraint _constraint;

        /**
         * Constructor which Initializes the cell range on which this object will be
         * applied
         * @param constraint 
         */
        public HSSFDataValidation(CellRangeAddressList regions, IDataValidationConstraint constraint)
        {
            _regions = regions;

            //FIXME: This cast can be avoided.
            _constraint = (DVConstraint)constraint;
        }


        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidation#getConstraint()
         */
        public IDataValidationConstraint ValidationConstraint
        {
            get
            {
                return _constraint;
            }
        }

        public DVConstraint Constraint
        {
            get
            {
                return _constraint;
            }
        }

        public CellRangeAddressList Regions
        {
            get
            {
                return _regions;
            }
        }
        public int ErrorStyle
        {
            get
            {
                return _errorStyle;
            }
            set 
            { 
                _errorStyle = value; 
            }
        }

        public bool EmptyCellAllowed
        {
            get
            {
                return _emptyCellAllowed;
            }
            set
            {
                _emptyCellAllowed = value;
            }
        }

        public bool SuppressDropDownArrow
        {
            get
            {
                if (_constraint.GetValidationType() == ValidationType.LIST)
                {
                    return _suppress_dropdown_arrow;
                }
                return false;
            }
            set
            {
                _suppress_dropdown_arrow = value;
            }
        }
        public bool ShowPromptBox
        {
            get
            {
                return _ShowPromptBox;
            }
            set
            {
                _ShowPromptBox = value;
            }
        }
        public bool ShowErrorBox
        {
            get
            {
                return _ShowErrorBox;
            }
            set
            {
                _ShowErrorBox = value;
            }
        }



        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidation#CreatePromptBox(java.lang.String, java.lang.String)
         */
        public void CreatePromptBox(String title, String text)
        {
            _prompt_title = title;
            _prompt_text = text;
            this.ShowPromptBox = (/*setter*/true);
        }

        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidation#getPromptBoxTitle()
         */
        public String PromptBoxTitle
        {
            get
            {
                return _prompt_title;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidation#getPromptBoxText()
         */
        public String PromptBoxText
        {
            get
            {
                return _prompt_text;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidation#CreateErrorBox(java.lang.String, java.lang.String)
         */
        public void CreateErrorBox(String title, String text)
        {
            _error_title = title;
            _error_text = text;
            this.ShowErrorBox = (/*setter*/true);
        }

        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidation#getErrorBoxTitle()
         */
        public String ErrorBoxTitle
        {
            get
            {
                return _error_title;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidation#getErrorBoxText()
         */
        public String ErrorBoxText
        {
            get
            {
                return _error_text;
            }
        }

        public DVRecord CreateDVRecord(HSSFSheet sheet)
        {

            NPOI.HSSF.UserModel.DVConstraint.FormulaPair fp = _constraint.CreateFormulas(sheet);

            return new DVRecord(_constraint.GetValidationType(),
                    _constraint.Operator,
                    _errorStyle, _emptyCellAllowed, SuppressDropDownArrow,
                    _constraint.GetValidationType() == ValidationType.LIST && _constraint.ExplicitListValues != null,
                    _ShowPromptBox, _prompt_title, _prompt_text,
                    _ShowErrorBox, _error_title, _error_text,
                    fp.Formula1, fp.Formula2,
                    _regions);
        }
    }

}