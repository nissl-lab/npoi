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

    /// <summary>
    /// <para>
    /// Types of ignored workbook and worksheet error.
    /// </para>
    /// <para>
    /// TODO Implement these for HSSF too, using FeatFormulaErr2,
    ///  see bugzilla bug #46136 for details
    /// </para>
    /// </summary>
    public enum IgnoredErrorType
    {
        /// <summary>
        /// ????. Probably XSSF-only.
        /// </summary>
        CalculatedColumn,

        /// <summary>
        /// Whether to check for references to empty cells.
        /// HSSF + XSSF.
        /// </summary>
        EmptyCellReference,

        /// <summary>
        /// Whether to check for calculation/Evaluation errors.
        /// HSSF + XSSF.
        /// </summary>
        EvaluationError,

        /// <summary>
        /// Whether to check formulas in the range of the shared feature
        ///  that are inconsistent with formulas in neighbouring cells.
        /// HSSF + XSSF.
        /// </summary>
        Formula,

        /// <summary>
        /// Whether to check formulas in the range of the shared feature
        /// with references to less than the entirety of a range Containing
        /// continuous data.
        /// HSSF + XSSF.
        /// </summary>
        FormulaRange,

        /// <summary>
        /// ????. Is this XSSF-specific the same as performDataValidation
        ///  in HSSF?
        /// </summary>
        ListDataValidation,

        /// <summary>
        /// Whether to check the format of string values and warn
        ///  if they look to actually be numeric values.
        /// HSSF + XSSF.
        /// </summary>
        NumberStoredAsText,

        /// <summary>
        /// ????. Is this XSSF-specific the same as CheckDateTimeFormats
        ///  in HSSF?
        /// </summary>
        TwoDigitTextYear,

        /// <summary>
        /// Whether to check for unprotected formulas.
        /// HSSF + XSSF.
        /// </summary>
        UnlockedFormula
    }

    public static class IgnoredErrorTypeValues
    {
        public static IgnoredErrorType[] Values = new IgnoredErrorType[] {
            IgnoredErrorType.CalculatedColumn,
            IgnoredErrorType.EmptyCellReference,
            IgnoredErrorType.EvaluationError,
            IgnoredErrorType.Formula,
            IgnoredErrorType.FormulaRange,
            IgnoredErrorType.ListDataValidation,
            IgnoredErrorType.NumberStoredAsText,
            IgnoredErrorType.TwoDigitTextYear,
            IgnoredErrorType.UnlockedFormula
        };
    }
}