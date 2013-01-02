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
    using NPOI.SS.Formula.Eval;

    /**
     * Calculates the internal rate of return.
     *
     * Syntax is IRR(values) or IRR(values,guess)
     *
     * @author Marcel May
     * @author Yegor Kozlov
     *
     * @see <a href="http://en.wikipedia.org/wiki/Internal_rate_of_return#Numerical_solution">Wikipedia on IRR</a>
     * @see <a href="http://office.microsoft.com/en-us/excel-help/irr-HP005209146.aspx">Excel IRR</a>
     */
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

        /**
         * Computes the internal rate of return using an estimated irr of 10 percent.
         *
         * @param income the income values.
         * @return the irr.
         */
        public static double irr(double[] income)
        {
            return irr(income, 0.1d);
        }


        /**
         * Calculates IRR using the Newton-Raphson Method.
         * <p>
         * Starting with the guess, the method cycles through the calculation until the result
         * is accurate within 0.00001 percent. If IRR can't find a result that works
         * after 20 tries, the Double.NaN is returned.
         * </p>
         * <p>
         *   The implementation is inspired by the NewtonSolver from the Apache Commons-Math library,
         *   @see <a href="http://commons.apache.org">http://commons.apache.org</a>
         * </p>
         *
         * @param values        the income values.
         * @param guess         the initial guess of irr.
         * @return the irr value. The method returns <code>Double.NaN</code>
         *  if the maximum iteration count is exceeded
         *
         * @see <a href="http://en.wikipedia.org/wiki/Internal_rate_of_return#Numerical_solution">
         *     http://en.wikipedia.org/wiki/Internal_rate_of_return#Numerical_solution</a>
         * @see <a href="http://en.wikipedia.org/wiki/Newton%27s_method">
         *     http://en.wikipedia.org/wiki/Newton%27s_method</a>
         */
        public static double irr(double[] values, double guess)
        {
            int maxIterationCount = 20;
            double absoluteAccuracy = 1E-7;

            double x0 = guess;
            double x1;

            int i = 0;
            while (i < maxIterationCount)
            {

                // the value of the function (NPV) and its derivate can be calculated in the same loop
                double fValue = 0;
                double fDerivative = 0;
                for (int k = 0; k < values.Length; k++)
                {
                    fValue += values[k] / Math.Pow(1.0 + x0, k);
                    fDerivative += -k * values[k] / Math.Pow(1.0 + x0, k + 1);
                }

                // the essense of the Newton-Raphson Method
                x1 = x0 - fValue / fDerivative;

                if (Math.Abs(x1 - x0) <= absoluteAccuracy)
                {
                    return x1;
                }

                x0 = x1;
                ++i;
            }
            // maximum number of iterations is exceeded
            return Double.NaN;
        }
    }
}