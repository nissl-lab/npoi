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

    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
     * 
     */
    public class HSSFDataValidationHelper : IDataValidationHelper
    {
        private HSSFSheet sheet;

        public HSSFDataValidationHelper(HSSFSheet sheet)
            : base()
        {

            this.sheet = sheet;
        }

        /*
         * (non-Javadoc)
         * 
         * @see
         * NPOI.SS.UserModel.DataValidationHelper#CreateDateConstraint
         * (int, java.lang.String, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateDateConstraint(int operatorType, String formula1, String formula2, String dateFormat)
        {
            return DVConstraint.CreateDateConstraint(operatorType, formula1, formula2, dateFormat);
        }

        /*
         * (non-Javadoc)
         * 
         * @see
         * NPOI.SS.UserModel.DataValidationHelper#CreateExplicitListConstraint
         * (java.lang.String[])
         */
        public IDataValidationConstraint CreateExplicitListConstraint(String[] listOfValues)
        {
            return DVConstraint.CreateExplicitListConstraint(listOfValues);
        }

        /*
         * (non-Javadoc)
         * 
         * @see
         * NPOI.SS.UserModel.DataValidationHelper#CreateFormulaListConstraint
         * (java.lang.String)
         */
        public IDataValidationConstraint CreateFormulaListConstraint(String listFormula)
        {
            return DVConstraint.CreateFormulaListConstraint(listFormula);
        }



        public IDataValidationConstraint CreateNumericConstraint(int validationType, int operatorType, String formula1, String formula2)
        {
            return DVConstraint.CreateNumericConstraint(validationType, operatorType, formula1, formula2);
        }

        public IDataValidationConstraint CreateintConstraint(int operatorType, String formula1, String formula2)
        {
            return DVConstraint.CreateNumericConstraint(ValidationType.INTEGER, operatorType, formula1, formula2);
        }

        /*
         * (non-Javadoc)
         * 
         * @see
         * NPOI.SS.UserModel.DataValidationHelper#CreateNumericConstraint
         * (int, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateDecimalConstraint(int operatorType, String formula1, String formula2)
        {
            return DVConstraint.CreateNumericConstraint(ValidationType.DECIMAL, operatorType, formula1, formula2);
        }

        /*
         * (non-Javadoc)
         * 
         * @see
         * NPOI.SS.UserModel.DataValidationHelper#CreateTextLengthConstraint
         * (int, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateTextLengthConstraint(int operatorType, String formula1, String formula2)
        {
            return DVConstraint.CreateNumericConstraint(ValidationType.TEXT_LENGTH, operatorType, formula1, formula2);
        }

        /*
         * (non-Javadoc)
         * 
         * @see
         * NPOI.SS.UserModel.DataValidationHelper#CreateTimeConstraint
         * (int, java.lang.String, java.lang.String, java.lang.String)
         */
        public IDataValidationConstraint CreateTimeConstraint(int operatorType, String formula1, String formula2)
        {
            return DVConstraint.CreateTimeConstraint(operatorType, formula1, formula2);
        }



        public IDataValidationConstraint CreateCustomConstraint(String formula)
        {
            return DVConstraint.CreateCustomFormulaConstraint(formula);
        }

        /*
         * (non-Javadoc)
         * 
         * @see
         * NPOI.SS.UserModel.DataValidationHelper#CreateValidation(org
         * .apache.poi.SS.UserModel.DataValidationConstraint,
         * NPOI.SS.Util.CellRangeAddressList)
         */
        public IDataValidation CreateValidation(IDataValidationConstraint constraint, CellRangeAddressList cellRangeAddressList)
        {
            return new HSSFDataValidation(cellRangeAddressList, constraint);
        }
    }

}