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
        static Dictionary<ST_DataValidationOperator, int> operatorTypeReverseMappings = new Dictionary<ST_DataValidationOperator, int>();
        static Dictionary<int, ST_DataValidationType> validationTypeMappings = new Dictionary<int, ST_DataValidationType>();
        static Dictionary<ST_DataValidationType, int> validationTypeReverseMappings = new Dictionary<ST_DataValidationType, int>();
        static Dictionary<int, ST_DataValidationErrorStyle> errorStyleMappings = new Dictionary<int, ST_DataValidationErrorStyle>();
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


        XSSFDataValidation(CellRangeAddressList regions, CT_DataValidation ctDataValidation)
            : base()
        {

            this.validationConstraint = GetConstraint(ctDataValidation);
            this.ctDdataValidation = ctDataValidation;
            this.regions = regions;
            this.ctDdataValidation.errorStyle = (ST_DataValidationErrorStyle.stop);
            this.ctDdataValidation.allowBlank = (true);
        }

        public XSSFDataValidation(XSSFDataValidationConstraint constraint, CellRangeAddressList regions, CT_DataValidation ctDataValidation)
            : base()
        {

            this.validationConstraint = constraint;
            this.ctDdataValidation = ctDataValidation;
            this.regions = regions;
            this.ctDdataValidation.errorStyle = (ST_DataValidationErrorStyle.stop);
            this.ctDdataValidation.allowBlank = (true);
        }

        CT_DataValidation GetCtDdataValidation()
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

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getEmptyCellAllowed()
         */
        public bool GetEmptyCellAllowed()
        {
            return ctDdataValidation.allowBlank;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getErrorBoxText()
         */
        public String GetErrorBoxText()
        {
            return ctDdataValidation.error;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getErrorBoxTitle()
         */
        public String GetErrorBoxTitle()
        {
            return ctDdataValidation.errorTitle;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getErrorStyle()
         */
        public int GetErrorStyle()
        {
            return (int)ctDdataValidation.errorStyle;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getPromptBoxText()
         */
        public String GetPromptBoxText()
        {
            return ctDdataValidation.prompt;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getPromptBoxTitle()
         */
        public String GetPromptBoxTitle()
        {
            return ctDdataValidation.promptTitle;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getShowErrorBox()
         */
        public bool GetShowErrorBox()
        {
            return ctDdataValidation.showErrorMessage;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getShowPromptBox()
         */
        public bool GetShowPromptBox()
        {
            return ctDdataValidation.showInputMessage;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getSuppressDropDownArrow()
         */
        public bool GetSuppressDropDownArrow()
        {
            return !ctDdataValidation.showDropDown;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#getValidationConstraint()
         */
        public IDataValidationConstraint GetValidationConstraint()
        {
            return validationConstraint;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#setEmptyCellAllowed(bool)
         */
        public void SetEmptyCellAllowed(bool allowed)
        {
            ctDdataValidation.allowBlank = (allowed);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#setErrorStyle(int)
         */
        public void SetErrorStyle(int errorStyle)
        {
            ctDdataValidation.errorStyle = (errorStyleMappings[errorStyle]);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#setShowErrorBox(bool)
         */
        public void SetShowErrorBox(bool show)
        {
            ctDdataValidation.showErrorMessage = show;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#setShowPromptBox(bool)
         */
        public void SetShowPromptBox(bool show)
        {
            ctDdataValidation.showInputMessage=show;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidation#setSuppressDropDownArrow(bool)
         */
        public void SetSuppressDropDownArrow(bool suppress)
        {
            if (validationConstraint.GetValidationType() == ValidationType.LIST)
            {
                ctDdataValidation.showDropDown=(!suppress);
            }
        }

        public CellRangeAddressList GetRegions()
        {
            return regions;
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

        private XSSFDataValidationConstraint GetConstraint(CT_DataValidation ctDataValidation)
        {
            XSSFDataValidationConstraint constraint = null;
            String formula1 = ctDataValidation.formula1;
            String formula2 = ctDataValidation.formula2;
            Enum operator1 = ctDataValidation.@operator;
            ST_DataValidationType type = ctDataValidation.type;
            int validationType = XSSFDataValidation.validationTypeReverseMappings[type];
            int operatorType = XSSFDataValidation.operatorTypeReverseMappings.Get(operator1);
            constraint = new XSSFDataValidationConstraint(validationType, operatorType, formula1, formula2);
            return constraint;
        }
    }

}

