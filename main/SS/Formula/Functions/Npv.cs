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
* distributed Under the License Is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 15, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;

    public class Npv : Function
    {
        [Obsolete]
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                double rate = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                result = Evaluate(rate, d1);
                NumericFunction.CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }
        [Obsolete]
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2)
        {
            double result;
            try
            {
                double rate = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                double d2 = NumericFunction.SingleOperandEvaluate(arg2, srcRowIndex, srcColumnIndex);
                result = Evaluate(rate, d1, d2);
                NumericFunction.CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }
        [Obsolete]
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2, ValueEval arg3)
        {
            double result;
            try
            {
                double rate = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                double d2 = NumericFunction.SingleOperandEvaluate(arg2, srcRowIndex, srcColumnIndex);
                double d3 = NumericFunction.SingleOperandEvaluate(arg3, srcRowIndex, srcColumnIndex);
                result = Evaluate(rate, d1, d2, d3);
                NumericFunction.CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            int nArgs = args.Length;
            if (nArgs < 2)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                double rate = NumericFunction.SingleOperandEvaluate(args[0], srcRowIndex, srcColumnIndex);
                // convert tail arguments into an array of doubles
                ValueEval[] vargs = new ValueEval[args.Length - 1];
                Array.Copy(args, 1, vargs, 0, vargs.Length);
                double[] values = AggregateFunction.ValueCollector.CollectValues(vargs);

                double result = FinanceLib.npv(rate, values);
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private static double Evaluate(double rate, params double[] ds)
        {
            double sum = 0;
            for (int i = 0; i < ds.Length; i++)
            {
                sum += ds[i] / Math.Pow(rate + 1, i);
            }
            return sum;
        }
    }

}