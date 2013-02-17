/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/

namespace NPOI.SS.Formula.Eval
{
    using System;
    using System.Text;
    using NPOI.HSSF.UserModel;
    

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    public class ErrorEval : ValueEval
    {

        // convenient access to namespace
        private const HSSFErrorConstants EC = null;

        /** <b>#NULL!</b>  - Intersection of two cell ranges is empty */
        public static readonly ErrorEval NULL_INTERSECTION = new ErrorEval(HSSFErrorConstants.ERROR_NULL);
        /** <b>#DIV/0!</b> - Division by zero */
        public static readonly ErrorEval DIV_ZERO = new ErrorEval(HSSFErrorConstants.ERROR_DIV_0);
        /** <b>#VALUE!</b> - Wrong type of operand */
        public static readonly ErrorEval VALUE_INVALID = new ErrorEval(HSSFErrorConstants.ERROR_VALUE);
        /** <b>#REF!</b> - Illegal or deleted cell reference */
        public static readonly ErrorEval REF_INVALID = new ErrorEval(HSSFErrorConstants.ERROR_REF);
        /** <b>#NAME?</b> - Wrong function or range name */
        public static readonly ErrorEval NAME_INVALID = new ErrorEval(HSSFErrorConstants.ERROR_NAME);
        /** <b>#NUM!</b> - Value range overflow */
        public static readonly ErrorEval NUM_ERROR = new ErrorEval(HSSFErrorConstants.ERROR_NUM);
        /** <b>#N/A</b> - Argument or function not available */
        public static readonly ErrorEval NA = new ErrorEval(HSSFErrorConstants.ERROR_NA);


        // POI internal error codes
        private const int CIRCULAR_REF_ERROR_CODE = unchecked((int)0xFFFFFFC4);
        private const int FUNCTION_NOT_IMPLEMENTED_CODE = unchecked((int)0xFFFFFFE2);

        public static readonly ErrorEval FUNCTION_NOT_IMPLEMENTED = new ErrorEval(FUNCTION_NOT_IMPLEMENTED_CODE);
        // Note - Excel does not seem to represent this condition with an error code
        public static readonly ErrorEval CIRCULAR_REF_ERROR = new ErrorEval(CIRCULAR_REF_ERROR_CODE);


        /**
         * Translates an Excel internal error code into the corresponding POI ErrorEval instance
         * @param errorCode
         */
        public static ErrorEval ValueOf(int errorCode)
        {
            switch (errorCode)
            {
                case HSSFErrorConstants.ERROR_NULL: return NULL_INTERSECTION;
                case HSSFErrorConstants.ERROR_DIV_0: return DIV_ZERO;
                case HSSFErrorConstants.ERROR_VALUE: return VALUE_INVALID;
                case HSSFErrorConstants.ERROR_REF: return REF_INVALID;
                case HSSFErrorConstants.ERROR_NAME: return NAME_INVALID;
                case HSSFErrorConstants.ERROR_NUM: return NUM_ERROR;
                case HSSFErrorConstants.ERROR_NA: return NA;
                // non-std errors (conditions modeled as errors by POI)
                case CIRCULAR_REF_ERROR_CODE: return CIRCULAR_REF_ERROR;
                case FUNCTION_NOT_IMPLEMENTED_CODE: return FUNCTION_NOT_IMPLEMENTED;
            }
            throw new Exception("Unexpected error code (" + errorCode + ")");
        }

        /**
         * Converts error codes to text.  Handles non-standard error codes OK.  
         * For debug/test purposes (and for formatting error messages).
         * @return the String representation of the specified Excel error code.
         */
        public static String GetText(int errorCode)
        {
            if (HSSFErrorConstants.IsValidCode(errorCode))
            {
                return HSSFErrorConstants.GetText(errorCode);
            }
            // It is desirable to make these (arbitrary) strings look clearly different from any other
            // value expression that might appear in a formula.  In Addition these error strings should
            // look Unlike the standard Excel errors.  Hence tilde ('~') was used.
            switch (errorCode)
            {
                case CIRCULAR_REF_ERROR_CODE: return "~CIRCULAR~REF~";
                case FUNCTION_NOT_IMPLEMENTED_CODE: return "~FUNCTION~NOT~IMPLEMENTED~";
            }
            return "~non~std~err(" + errorCode + ")~";
        }

        private int _errorCode;
        /**
         * @param errorCode an 8-bit value
         */
        private ErrorEval(int errorCode)
        {
            _errorCode = errorCode;
        }

        public int ErrorCode
        {
            get{return _errorCode;}
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(GetText(_errorCode));
            sb.Append("]");
            return sb.ToString();
        }
    }
}