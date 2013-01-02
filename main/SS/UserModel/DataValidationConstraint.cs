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


    public interface IDataValidationConstraint
    {

        /**
         * @return data validation type of this constraint
         * @see ValidationType
         */
        int GetValidationType();


        /**
        * @return the operator used for this constraint
        * @see OperatorType
        */
        /// <summary>
        /// get or set then comparison operator for this constraint
        /// </summary>
        int Operator { get; set; }


        String[] ExplicitListValues { get; set; }

        /// <summary>
        /// get or set the formula for expression 1. May be <code>null</code>
        /// </summary>
        string Formula1 { get; set; }


        /// <summary>
        /// get or set the formula for expression 2. May be <code>null</code>
        /// </summary>
        string Formula2 { get; set; }


        
    }
    /**
         * ValidationType enum
         */
    public static class ValidationType
    {
        /** 'Any value' type - value not restricted */
        public const int ANY = 0x00;
        /** int ('Whole number') type */
        public const int INTEGER = 0x01;
        /** Decimal type */
        public const int DECIMAL = 0x02;
        /** List type ( combo box type ) */
        public const int LIST = 0x03;
        /** Date type */
        public const int DATE = 0x04;
        /** Time type */
        public const int TIME = 0x05;
        /** String length type */
        public const int TEXT_LENGTH = 0x06;
        /** Formula ( 'Custom' ) type */
        public const int FORMULA = 0x07;
    }
    /**
     * Condition operator enum
     */
    public static class OperatorType
    {
       
        public const int BETWEEN = 0x00;
        public const int NOT_BETWEEN = 0x01;
        public const int EQUAL = 0x02;
        public const int NOT_EQUAL = 0x03;
        public const int GREATER_THAN = 0x04;
        public const int LESS_THAN = 0x05;
        public const int GREATER_OR_EQUAL = 0x06;
        public const int LESS_OR_EQUAL = 0x07;
        /** default value to supply when the operator type is not used */
        public const int IGNORED = BETWEEN;

        /* package */
        public static void ValidateSecondArg(int comparisonOperator, String paramValue)
        {
            switch (comparisonOperator)
            {
                case BETWEEN:
                    if (paramValue == null)
                    {
                        throw new ArgumentException("expr2 must be supplied for 'between' comparisons");
                    }
                    break;
                case NOT_BETWEEN:
                    if (paramValue == null)
                    {
                        throw new ArgumentException("expr2 must be supplied for 'between' comparisons");
                    }
                    break;
                // all other operators don't need second arg
            }
        }
    }
}