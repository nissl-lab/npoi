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


namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;

    /**
     * Implementation for the function COUNTIFS
     * <p>
     * Syntax: COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2])
     * </p>
     */

    public class Countifs : FreeRefFunction
    {
        public static FreeRefFunction instance = new Countifs();

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            Double result = double.NaN;
            if (args.Length == 0 || args.Length % 2 > 0)
            {
                return ErrorEval.VALUE_INVALID;
            }
            for (int i = 0; i < args.Length; )
            {
                ValueEval firstArg = args[i];
                ValueEval secondArg = args[i + 1];
                i += 2;
                NumberEval Evaluate = (NumberEval)new Countif().Evaluate(new ValueEval[] { firstArg, secondArg }, ec.RowIndex, ec.ColumnIndex);
                if (double.IsNaN(result))
                {
                    result = Evaluate.NumberValue;
                }
                else if (Evaluate.NumberValue < result)
                {
                    result = Evaluate.NumberValue;
                }
            }
            return new NumberEval(double.IsNaN(result) ? 0 : result);
        }
    }

}
