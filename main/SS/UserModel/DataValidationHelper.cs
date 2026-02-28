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
namespace NPOI.SS.UserModel
{
    using System;

    using NPOI.SS.Util;

    /**
     * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
     * 
     */
    public interface IDataValidationHelper
    {

        IDataValidationConstraint CreateFormulaListConstraint(String listFormula);

        IDataValidationConstraint CreateExplicitListConstraint(String[] listOfValues);

        IDataValidationConstraint CreateNumericConstraint(int validationType, int operatorType, String formula1, String formula2);

        IDataValidationConstraint CreateTextLengthConstraint(int operatorType, String formula1, String formula2);

        IDataValidationConstraint CreateDecimalConstraint(int operatorType, String formula1, String formula2);

        IDataValidationConstraint CreateintConstraint(int operatorType, String formula1, String formula2);

        IDataValidationConstraint CreateDateConstraint(int operatorType, String formula1, String formula2, String dateFormat);

        IDataValidationConstraint CreateTimeConstraint(int operatorType, String formula1, String formula2);

        IDataValidationConstraint CreateCustomConstraint(String formula);

        IDataValidation CreateValidation(IDataValidationConstraint constraint, CellRangeAddressList cellRangeAddressList);
    }

}