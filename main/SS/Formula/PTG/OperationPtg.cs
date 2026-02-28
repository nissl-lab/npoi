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

namespace NPOI.SS.Formula.PTG
{
    using System;
    /**
     * defines a Ptg that is an operation instead of an operand
     * @author  andy
     */
    [Serializable]
    public abstract class OperationPtg : Ptg
    {
        public const int TYPE_UNARY = 0;
        public const int TYPE_BINARY = 1;
        public const int TYPE_FUNCTION = 2;

        /**
         *  returns a string representation of the operations
         *  the Length of the input array should equal the number returned by 
         *  @see #GetNumberOfOperands
         *  
         */
        public abstract String ToFormulaString(String[] operands);

        /**
         * The number of operands expected by the operations
         */
        public abstract int NumberOfOperands { get; }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_VALUE; }
        }
    }
}