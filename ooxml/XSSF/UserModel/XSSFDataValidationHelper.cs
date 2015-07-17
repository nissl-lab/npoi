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
namespace NPOI.XSSF.UserModel
{

    using NPOI.SS.UserModel;
    using System;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using System.Collections.Generic;
    using NPOI.SS.Util;

    /**
     * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
     *
     */
    public class XSSFDataValidationHelper : IDataValidationHelper
    {
        private XSSFSheet xssfSheet;


        public XSSFDataValidationHelper(XSSFSheet xssfSheet)
            : base()
        {

            this.xssfSheet = xssfSheet;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationHelper#CreateDateConstraint(int, java.lang.String, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateDateConstraint(int operatorType, String formula1, String formula2, String dateFormat)
        {
            return new XSSFDataValidationConstraint(ValidationType.DATE, operatorType, formula1, formula2);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationHelper#CreateDecimalConstraint(int, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateDecimalConstraint(int operatorType, String formula1, String formula2)
        {
            return new XSSFDataValidationConstraint(ValidationType.DECIMAL, operatorType, formula1, formula2);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationHelper#CreateExplicitListConstraint(java.lang.String[])
         */
        public IDataValidationConstraint CreateExplicitListConstraint(String[] listOfValues)
        {
            return new XSSFDataValidationConstraint(listOfValues);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationHelper#CreateFormulaListConstraint(java.lang.String)
         */
        public IDataValidationConstraint CreateFormulaListConstraint(String listFormula)
        {
            return new XSSFDataValidationConstraint(ValidationType.LIST, listFormula);
        }



        public IDataValidationConstraint CreateNumericConstraint(int validationType, int operatorType, String formula1, String formula2)
        {
            if (validationType == ValidationType.INTEGER)
            {
                return CreateintConstraint(operatorType, formula1, formula2);
            }
            else if (validationType == ValidationType.DECIMAL)
            {
                return CreateDecimalConstraint(operatorType, formula1, formula2);
            }
            else if (validationType == ValidationType.TEXT_LENGTH)
            {
                return CreateTextLengthConstraint(operatorType, formula1, formula2);
            }
            return null;
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationHelper#CreateintConstraint(int, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateintConstraint(int operatorType, String formula1, String formula2)
        {
            return new XSSFDataValidationConstraint(ValidationType.INTEGER, operatorType, formula1, formula2);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationHelper#CreateTextLengthConstraint(int, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateTextLengthConstraint(int operatorType, String formula1, String formula2)
        {
            return new XSSFDataValidationConstraint(ValidationType.TEXT_LENGTH, operatorType, formula1, formula2);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationHelper#CreateTimeConstraint(int, java.lang.String, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateTimeConstraint(int operatorType, String formula1, String formula2)
        {
            return new XSSFDataValidationConstraint(ValidationType.TIME, operatorType, formula1, formula2);
        }

        public IDataValidationConstraint CreateCustomConstraint(String formula)
        {
            return new XSSFDataValidationConstraint(ValidationType.FORMULA, formula);
        }

        /* (non-Javadoc)
         * @see NPOI.ss.usermodel.DataValidationHelper#CreateValidation(NPOI.ss.usermodel.DataValidationConstraint, NPOI.ss.util.CellRangeAddressList)
         */
        public IDataValidation CreateValidation(IDataValidationConstraint constraint, CellRangeAddressList cellRangeAddressList)
        {
            XSSFDataValidationConstraint dataValidationConstraint = (XSSFDataValidationConstraint)constraint;
            CT_DataValidation newDataValidation = new CT_DataValidation();

            int validationType = constraint.GetValidationType();
            switch (validationType)
            {
                case ValidationType.LIST:
                    newDataValidation.type = (ST_DataValidationType.list);
                    newDataValidation.formula1 = (constraint.Formula1);
                    break;
                case ValidationType.ANY:
                    newDataValidation.type = ST_DataValidationType.none;
                    break;
                case ValidationType.TEXT_LENGTH:
                    newDataValidation.type = ST_DataValidationType.textLength;
                    break;
                case ValidationType.DATE:
                    newDataValidation.type = ST_DataValidationType.date;
                    break;
                case ValidationType.INTEGER:
                    newDataValidation.type = ST_DataValidationType.whole;
                    break;
                case ValidationType.DECIMAL:
                    newDataValidation.type = ST_DataValidationType.@decimal;
                    break;
                case ValidationType.TIME:
                    newDataValidation.type = ST_DataValidationType.time;
                    break;
                case ValidationType.FORMULA:
                    newDataValidation.type = ST_DataValidationType.custom;
                    break;
                default:
                    newDataValidation.type = ST_DataValidationType.none;
                    break;
            }

            if (validationType != ValidationType.ANY && validationType != ValidationType.LIST)
            {
                newDataValidation.@operator = ST_DataValidationOperator.between;
                if (XSSFDataValidation.operatorTypeMappings.ContainsKey(constraint.Operator))
                    newDataValidation.@operator = XSSFDataValidation.operatorTypeMappings[constraint.Operator];

                if (constraint.Formula1 != null)
                {
                    newDataValidation.formula1 = (constraint.Formula1);
                }
                if (constraint.Formula2 != null)
                {
                    newDataValidation.formula2 = (constraint.Formula2);
                }
            }

            CellRangeAddress[] cellRangeAddresses = cellRangeAddressList.CellRangeAddresses;
            string sqref = string.Empty;
            for (int i = 0; i < cellRangeAddresses.Length; i++)
            {
                CellRangeAddress cellRangeAddress = cellRangeAddresses[i];
                if(sqref.Length==0)
                    sqref = cellRangeAddress.FormatAsString();
                else
                    sqref = " " + cellRangeAddress.FormatAsString();
            }
            newDataValidation.sqref = sqref;
            newDataValidation.allowBlank = (true);
            newDataValidation.errorStyle = ST_DataValidationErrorStyle.stop;
            return new XSSFDataValidation(dataValidationConstraint, cellRangeAddressList, newDataValidation);
        }
    }
}


