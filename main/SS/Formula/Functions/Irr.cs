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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using EFF = Excel.FinancialFunctions;
    using NPOI.SS.Formula.Eval;

    /// <summary>
    /// Calculates the internal rate of return.
    /// <para>
    ///     Syntax is IRR(values) or IRR(values,guess)
    /// </para>
    /// </summary>
    /// <see href="http://en.wikipedia.org/wiki/Internal_rate_of_return#Numerical_solution">Wikipedia on IRR</see>
    /// <see href="http://office.microsoft.com/en-us/excel-help/irr-HP005209146.aspx">Excel IRR</see>
    public class Irr : Function
    {

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length == 0 || args.Length > 2)
            {
                // Wrong number of arguments
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                double[] values = AggregateFunction.ValueCollector.CollectValues(args[0]);
                double guess;
                if (args.Length == 2)
                {
                    guess = NumericFunction.SingleOperandEvaluate(args[1], srcRowIndex, srcColumnIndex);
                }
                else
                {
                    guess = 0.1d;
                }
                double result = irr(values, guess);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        public static double irr(double[] income)
        {
            return irr(income, 0.1d);
        }

        public static double irr(double[] values, double guess)
        {
            try
            {
                return EFF.Financial.Irr(values, guess);
            }
            catch
            {
                return double.NaN;
            }
        }
    }
}