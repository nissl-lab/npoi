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
using System.Collections.Generic;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using System;
using System.Text;
using NPOI.SS.Util;
namespace NPOI.XSSF.UserModel
{

    /**
     * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
     *
     */
    public class XSSFDataValidation : IDataValidation
    {
        private CT_DataValidation ctDdataValidation;
        private XSSFDataValidationConstraint validationConstraint;
        private CellRangeAddressList regions;

        internal static Dictionary<int, ST_DataValidationOperator> operatorTypeMappings = new Dictionary<int, ST_DataValidationOperator>();
        internal static Dictionary<ST_DataValidationOperator, int> operatorTypeReverseMappings = new Dictionary<ST_DataValidationOperator, int>();
        internal static Dictionary<int, ST_DataValidationType> validationTypeMappings = new Dictionary<int, ST_DataValidationType>();
        internal static Dictionary<ST_DataValidationType, int> validationTypeReverseMappings = new Dictionary<ST_DataValidationType, int>();
        internal static Dictionary<int, ST_DataValidationErrorStyle> errorStyleMappings = new Dictionary<int, ST_DataValidationErrorStyle>();
        static XSSFDataValidation()
        {

            errorStyleMappings[ERRORSTYLE.INFO]= ST_DataValidationErrorStyle.information;
            errorStyleMappings[ERRORSTYLE.STOP]= ST_DataValidationErrorStyle.stop;
            errorStyleMappings[ERRORSTYLE.WARNING]= ST_DataValidationErrorStyle.warning;

            
            operatorTypeMappings[OperatorType.BETWEEN] =  ST_DataValidationOperator.between;
            operatorTypeMappings[OperatorType.NOT_BETWEEN] =  ST_DataValidationOperator.notBetween;
            operatorTypeMappings[OperatorType.EQUAL] =  ST_DataValidationOperator.equal;
            operatorTypeMappings[OperatorType.NOT_EQUAL] =  ST_DataValidationOperator.notEqual;
            operatorTypeMappings[OperatorType.GREATER_THAN] =  ST_DataValidationOperator.greaterThan;
            operatorTypeMappings[OperatorType.GREATER_OR_EQUAL] =  ST_DataValidationOperator.greaterThanOrEqual;
            operatorTypeMappings[OperatorType.LESS_THAN] =  ST_DataValidationOperator.lessThan;
            operatorTypeMappings[OperatorType.LESS_OR_EQUAL] =  ST_DataValidationOperator.lessThanOrEqual;

            foreach (KeyValuePair<int, ST_DataValidationOperator> entry in operatorTypeMappings)
            {
                operatorTypeReverseMappings[entry.Value]=entry.Key;
            }

            validationTypeMappings[ValidationType.FORMULA] =  ST_DataValidationType.custom;
            validationTypeMappings[ValidationType.DATE] =  ST_DataValidationType.date;
            validationTypeMappings[ValidationType.DECIMAL] =  ST_DataValidationType.@decimal;
            validationTypeMappings[ValidationType.LIST] =  ST_DataValidationType.list;
            validationTypeMappings[ValidationType.ANY] =  ST_DataValidationType.none;
            validationTypeMappings[ValidationType.TEXT_LENGTH] =  ST_DataValidationType.textLength;
            validationTypeMappings[ValidationType.TIME] =  ST_DataValidationType.time;
            validationTypeMappings[ValidationType.INTEGER] =  ST_DataValidationType.whole;

            foreach (KeyValuePair<int, ST_DataValidationType> entry in validationTypeMappings)
            {
                validationTypeReverseMappings[entry.Value]= entry.Key;
            }
        }


        public XSSFDataValidation(CellRangeAddressList regions, CT_DataValidation ctDataValidation)
            : this(GetConstraint(ctDataValidation), regions, ctDataValidation)
        {
        }

        public XSSFDataValidation(XSSFDataValidationConstraint constraint, CellRangeAddressList regions, CT_DataValidation ctDataValidation)
            : base()
        {

            this.validationConstraint = constraint;
            this.ctDdataValidation = ctDataValidation;
            this.regions = regions;
        }

        internal CT_DataValidation GetCTDataValidation()
        {
            return ctDdataValidation;
        }



        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#CreateErrorBox(java.lang.String, java.lang.String)
         */
        public void CreateErrorBox(String title, String text)
        {
            ctDdataValidation.errorTitle = (title);
            ctDdataValidation.error = (text);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#CreatePromptBox(java.lang.String, java.lang.String)
         */
        public void CreatePromptBox(String title, String text)
        {
            ctDdataValidation.promptTitle = (title);
            ctDdataValidation.prompt = (text);
        }

        public bool EmptyCellAllowed
        {
            get
            {
                return ctDdataValidation.allowBlank;
            }
            set 
            {
                ctDdataValidation.allowBlank =value;
            }
        }

        public String ErrorBoxText
        {
            get
            {
                return ctDdataValidation.error;
            }
        }

        public String ErrorBoxTitle
        {
            get
            {
                return ctDdataValidation.errorTitle;
            }
        }

        public int ErrorStyle
        {
            get
            {
                return (int)ctDdataValidation.errorStyle;
            }
            set 
            {

                ctDdataValidation.errorStyle = (errorStyleMappings[value]);
            }
        }

        public String PromptBoxText
        {
            get
            {
                return ctDdataValidation.prompt;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getPromptBoxTitle()
         */
        public String PromptBoxTitle
        {
            get
            {
                return ctDdataValidation.promptTitle;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getShowErrorBox()
         */
        public bool ShowErrorBox
        {
            get
            {
                return ctDdataValidation.showErrorMessage;
            }
            set 
            {
                ctDdataValidation.showErrorMessage = value;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getShowPromptBox()
         */
        public bool ShowPromptBox
        {
            get
            {
                return ctDdataValidation.showInputMessage;
            }
            set 
            {
                ctDdataValidation.showInputMessage = value;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getSuppressDropDownArrow()
         */
        public bool SuppressDropDownArrow
        {
            get
            {
                return !ctDdataValidation.showDropDown;
            }
            set 
            {
                if (validationConstraint.GetValidationType() == ValidationType.LIST)
                {
                    ctDdataValidation.showDropDown = (!value);
                }
            }
        }

        public IDataValidationConstraint ValidationConstraint
        {
            get
            {
                return validationConstraint;
            }
        }

        public CellRangeAddressList Regions
        {
            get
            {
                return regions;
            }
        }

        public String PrettyPrint()
        {
            StringBuilder builder = new StringBuilder();
            foreach (CellRangeAddress Address in regions.CellRangeAddresses)
            {
                builder.Append(Address.FormatAsString());
            }
            builder.Append(" => ");
            builder.Append(this.validationConstraint.PrettyPrint());
            return builder.ToString();
        }

        private static XSSFDataValidationConstraint GetConstraint(CT_DataValidation ctDataValidation)
        {
            XSSFDataValidationConstraint constraint = null;
            String formula1 = ctDataValidation.formula1;
            String formula2 = ctDataValidation.formula2;
            ST_DataValidationOperator operator1 = ctDataValidation.@operator;
            ST_DataValidationType type = ctDataValidation.type;
            int validationType = XSSFDataValidation.validationTypeReverseMappings[type];
            int operatorType = XSSFDataValidation.operatorTypeReverseMappings[operator1];
            constraint = new XSSFDataValidationConstraint(validationType, operatorType, formula1, formula2);
            return constraint;
        }
    }

}

