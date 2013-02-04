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

namespace NPOI.HSSF.UserModel
{
    using System;

    /// <summary>
    /// Contains raw Excel error codes (as defined in OOO's excelfileformat.pdf (2.5.6)
    /// @author  Michael Harhen
    /// </summary>
    public class HSSFErrorConstants
    {
        private HSSFErrorConstants()
        {
            // no instances of this class
        }

        /** <b>#NULL!</b>  - Intersection of two cell ranges is empty */
        public const int ERROR_NULL = 0x00;
        /** <b>#DIV/0!</b> - Division by zero */
        public const int ERROR_DIV_0 = 0x07;
        /** <b>#VALUE!</b> - Wrong type of operand */
        public const int ERROR_VALUE = 0x0F;
        /** <b>#REF!</b> - Illegal or deleted cell reference */
        public const int ERROR_REF = 0x17;
        /** <b>#NAME?</b> - Wrong function or range name */
        public const int ERROR_NAME = 0x1D;
        /** <b>#NUM!</b> - Value range overflow */
        public const int ERROR_NUM = 0x24;
        /** <b>#N/A</b> - Argument or function not available */
        public const int ERROR_NA = 0x2A;


        /// <summary>
        /// Gets standard Excel error literal for the specified error code.
        /// @throws ArgumentException if the specified error code is not one of the 7
        /// standard error codes
        /// </summary>
        /// <param name="errorCode">The error code.</param>
        /// <returns></returns>
        public static String GetText(int errorCode)
        {
            if (errorCode == ERROR_NULL)
            {
                return "#NULL!";
            }else if(errorCode == ERROR_DIV_0)
            {
                return "#DIV/0!";
            }else if(errorCode == ERROR_VALUE)
            {
                return "#VALUE!";
            }else if(errorCode == ERROR_REF)
            {
                return "#REF!";
            }else if(errorCode == ERROR_NAME)
            {
                return "#NAME?";
            }else if(errorCode == ERROR_NUM)
            {
                return "#NUM!";
            }
            else if (errorCode == ERROR_NA)
            {
                return "#N/A";
            }
            throw new ArgumentException("Bad error code (" + errorCode + ")");
        }

        /// <summary>
        /// Determines whether [is valid code] [the specified error code].
        /// </summary>
        /// <param name="errorCode">The error code.</param>
        /// <returns>
        /// 	<c>true</c> if the specified error code is a standard Excel error code.; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsValidCode(int errorCode)
        {
            // This method exists because it would be bad to force clients to catch 
            // ArgumentException if there were potential for passing an invalid error code.

            if (errorCode == ERROR_NULL
                || errorCode == ERROR_DIV_0
                || errorCode == ERROR_VALUE
                || errorCode == ERROR_REF
                || errorCode == ERROR_NAME
                || errorCode == ERROR_NUM
                || errorCode == ERROR_NA
                )
                {
                    return true;
                }

            return false;
        }
    }
}