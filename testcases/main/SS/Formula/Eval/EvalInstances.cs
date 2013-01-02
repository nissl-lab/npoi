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

namespace TestCases.SS.Formula.Eval
{

    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.Eval;

    /**
     * Collects eval instances for easy access by Tests in this package
     *
     * @author Josh Micich
     */
    public class EvalInstances
    {
        private EvalInstances()
        {
            // no instances of this class
        }

        public static Function Add = TwoOperandNumericOperation.AddEval;
        public static Function Subtract = TwoOperandNumericOperation.SubtractEval;
        public static Function Multiply = TwoOperandNumericOperation.MultiplyEval;
        public static Function Divide = TwoOperandNumericOperation.DivideEval;

        public static Function Power = TwoOperandNumericOperation.PowerEval;

        public static Function Percent = PercentEval.instance;

        public static Function UnaryMinus = UnaryMinusEval.instance;
        public static Function UnaryPlus = UnaryPlusEval.instance;

        public static Function Equal = RelationalOperationEval.EqualEval;
        public static Function LessThan = RelationalOperationEval.LessThanEval;
        public static Function LessEqual = RelationalOperationEval.LessEqualEval;
        public static Function GreaterThan = RelationalOperationEval.GreaterThanEval;
        public static Function GreaterEqual = RelationalOperationEval.GreaterEqualEval;
        public static Function NotEqual = RelationalOperationEval.NotEqualEval;

        public static Function Range = RangeEval.instance;
        public static Function Concat = ConcatEval.instance;
    }

}