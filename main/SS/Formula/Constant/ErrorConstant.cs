/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.Constant
{
    using System;
    using System.Text;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /// <summary>
    /// Represents a constant error code value as encoded in a constant values array.
    /// This class is a type-safe wrapper for a 16-bit int value performing a similar job to
    /// <c>ErrorEval</c>
    /// </summary>
    ///<remarks> @author Josh Micich</remarks>
    public class ErrorConstant
    {
        private static readonly ErrorConstant NULL = new ErrorConstant(FormulaError.NULL.Code);
        private static readonly ErrorConstant DIV_0 = new ErrorConstant(FormulaError.DIV0.Code);
        private static readonly ErrorConstant VALUE = new ErrorConstant(FormulaError.VALUE.Code);
        private static readonly ErrorConstant REF = new ErrorConstant(FormulaError.REF.Code);
        private static readonly ErrorConstant NAME = new ErrorConstant(FormulaError.NAME.Code);
        private static readonly ErrorConstant NUM = new ErrorConstant(FormulaError.NUM.Code);
        private static readonly ErrorConstant NA = new ErrorConstant(FormulaError.NA.Code);

        private readonly int _errorCode;

        /// <summary>
        /// Initializes a new instance of the <see cref="ErrorConstant"/> class.
        /// </summary>
        /// <param name="errorCode">The error code.</param>
        private ErrorConstant(int errorCode)
        {
            _errorCode = errorCode;
        }

        /// <summary>
        /// Gets the error code.
        /// </summary>
        /// <value>The error code.</value>
        public int ErrorCode
        {
            get { return _errorCode; }
        }
        /// <summary>
        /// Gets the text.
        /// </summary>
        /// <value>The text.</value>
        public String Text
        {
            get
            {
                if (FormulaError.IsValidCode(_errorCode))
                {
                    return FormulaError.ForInt(_errorCode).String;
                }
                return "unknown error code (" + _errorCode + ")";
            }
        }

        /// <summary>
        /// Values the of.
        /// </summary>
        /// <param name="errorCode">The error code.</param>
        /// <returns></returns>
        public static ErrorConstant ValueOf(int errorCode)
        {
            if (FormulaError.IsValidCode(errorCode))
            {
                switch ((FormulaErrorEnum)errorCode)
                {
                    case FormulaErrorEnum.NULL: return NULL;
                    case FormulaErrorEnum.DIV_0: return DIV_0;
                    case FormulaErrorEnum.VALUE: return VALUE;
                    case FormulaErrorEnum.REF: return REF;
                    case FormulaErrorEnum.NAME: return NAME;
                    case FormulaErrorEnum.NUM: return NUM;
                    case FormulaErrorEnum.NA: return NA;
                    default: break;
                }
            }
            Console.Error.WriteLine("Warning - Unexpected error code (" + errorCode + ")");
            return new ErrorConstant(errorCode);
        }
        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(Text);
            sb.Append("]");
            return sb.ToString();
        }
    }
}