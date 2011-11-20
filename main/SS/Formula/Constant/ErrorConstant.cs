/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.SS.Formula.Constant
{
    using NPOI.SS.UserModel;
    using System.Text;
    using System;
    /**
     * Represents a constant error code value as encoded in a constant values array. <p/>
     * 
     * This class is a type-safe wrapper for a 16-bit int value performing a similar job to 
     * <tt>ErrorEval</tt>.
     * 
     * @author Josh Micich
     */
    public class ErrorConstant
    {
        // convenient access to name space
        //private static ErrorConstants EC = null;

        private static ErrorConstant NULL = new ErrorConstant(ErrorConstants.ERROR_NULL);
        private static ErrorConstant DIV_0 = new ErrorConstant(ErrorConstants.ERROR_DIV_0);
        private static ErrorConstant VALUE = new ErrorConstant(ErrorConstants.ERROR_VALUE);
        private static ErrorConstant REF = new ErrorConstant(ErrorConstants.ERROR_REF);
        private static ErrorConstant NAME = new ErrorConstant(ErrorConstants.ERROR_NAME);
        private static ErrorConstant NUM = new ErrorConstant(ErrorConstants.ERROR_NUM);
        private static ErrorConstant NA = new ErrorConstant(ErrorConstants.ERROR_NA);

        private int _errorCode;

        private ErrorConstant(int errorCode)
        {
            _errorCode = errorCode;
        }

        public int ErrorCode
        {
            get
            {
                return _errorCode;
            }
        }
        public string Text
        {
            get
            {
                if (ErrorConstants.IsValidCode(_errorCode))
                {
                    return ErrorConstants.GetText(_errorCode);
                }
                return "unknown error code (" + _errorCode + ")";
            }
        }

        public static ErrorConstant ValueOf(int errorCode)
        {
            switch (errorCode)
            {
                case ErrorConstants.ERROR_NULL: return NULL;
                case ErrorConstants.ERROR_DIV_0: return DIV_0;
                case ErrorConstants.ERROR_VALUE: return VALUE;
                case ErrorConstants.ERROR_REF: return REF;
                case ErrorConstants.ERROR_NAME: return NAME;
                case ErrorConstants.ERROR_NUM: return NUM;
                case ErrorConstants.ERROR_NA: return NA;
            }
            System.Console.WriteLine("Warning - unexpected error code (" + errorCode + ")");
            return new ErrorConstant(errorCode);
        }
        public String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(this.GetType().Name).Append(" [");
            sb.Append(Text);
            sb.Append("]");
            return sb.ToString();
        }
    }
}
