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

namespace NPOI.XSSF.UserModel.Helpers
{
    using System;
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.UserModel;

    /**
     * XSSF-specific code for working with ignored errors
     */
    public class XSSFIgnoredErrorHelper
    {
        
        public static bool IsSet(IgnoredErrorType errorType, CT_IgnoredError error)
        {
            switch (errorType)
            {
                case IgnoredErrorType.CalculatedColumn:
                    return error.calculatedColumn;
                case IgnoredErrorType.EmptyCellReference:
                    return error.emptyCellReference;
                case IgnoredErrorType.EvaluationError:
                    return error.evalError;
                case IgnoredErrorType.Formula:
                    return error.formula;
                case IgnoredErrorType.FormulaRange:
                    return error.formulaRange;
                case IgnoredErrorType.ListDataValidation:
                    return error.listDataValidation;
                case IgnoredErrorType.NumberStoredAsText:
                    return error.numberStoredAsText;
                case IgnoredErrorType.TwoDigitTextYear:
                    return error.twoDigitTextYear;
                case IgnoredErrorType.UnlockedFormula:
                    return error.unlockedFormula;
                default:
                    throw new InvalidOperationException();
            }
        }

        public static void Set(IgnoredErrorType errorType, CT_IgnoredError error)
        {
            switch (errorType)
            {
                case IgnoredErrorType.CalculatedColumn:
                    error.calculatedColumn = true;
                    break;
                case IgnoredErrorType.EmptyCellReference:
                    error.emptyCellReference = true;
                    break;
                case IgnoredErrorType.EvaluationError:
                    error.evalError = true;
                    break;
                case IgnoredErrorType.Formula:
                    error.formula = true;
                    break;
                case IgnoredErrorType.FormulaRange:
                    error.formulaRange = true;
                    break;
                case IgnoredErrorType.ListDataValidation:
                    error.listDataValidation = true;
                    break;
                case IgnoredErrorType.NumberStoredAsText:
                    error.numberStoredAsText = true;
                    break;
                case IgnoredErrorType.TwoDigitTextYear:
                    error.twoDigitTextYear = true;
                    break;
                case IgnoredErrorType.UnlockedFormula:
                    error.unlockedFormula = true;
                    break;
                default:
                    throw new InvalidOperationException();
            }
        }

        public static void AddIgnoredErrors(CT_IgnoredError err, String ref1, params IgnoredErrorType[] ignoredErrorTypes)
        {
            err.sqref.Clear();
            err.sqref.Add(ref1);
            foreach (IgnoredErrorType errType in ignoredErrorTypes)
            {
                XSSFIgnoredErrorHelper.Set(errType, err);
            }
        }

        public static ISet<IgnoredErrorType> GetErrorTypes(CT_IgnoredError err)
        {
            ISet<IgnoredErrorType> result = new HashSet<IgnoredErrorType>();
            foreach (IgnoredErrorType errType in IgnoredErrorTypeValues.Values)
            {
                if (XSSFIgnoredErrorHelper.IsSet(errType, err))
                {
                    result.Add(errType);
                }
            }
            return result;
        }
    }

}