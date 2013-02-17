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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;


    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * Super class for all Evals for financial function evaluation.
     * 
     */
    public abstract class FinanceFunction : Function3Arg, Function4Arg
    {
        private static ValueEval DEFAULT_ARG3 = NumberEval.ZERO;
        private static ValueEval DEFAULT_ARG4 = BoolEval.FALSE;

        protected FinanceFunction()
        {

        }
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
        ValueEval arg2)
        {
            return Evaluate(srcRowIndex, srcColumnIndex, arg0, arg1, arg2, DEFAULT_ARG3);
        }
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2, ValueEval arg3)
        {
            return Evaluate(srcRowIndex, srcColumnIndex, arg0, arg1, arg2, arg3, DEFAULT_ARG4);
        }
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2, ValueEval arg3, ValueEval arg4)
        {
            double result;
            try
            {
                double d0 = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                double d2 = NumericFunction.SingleOperandEvaluate(arg2, srcRowIndex, srcColumnIndex);
                double d3 = NumericFunction.SingleOperandEvaluate(arg3, srcRowIndex, srcColumnIndex);
                double d4 = NumericFunction.SingleOperandEvaluate(arg4, srcRowIndex, srcColumnIndex);
                result = Evaluate(d0, d1, d2, d3, d4 != 0.0);
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
            switch (args.Length)
            {
                case 3:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2], DEFAULT_ARG3, DEFAULT_ARG4);
                case 4:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2], args[3], DEFAULT_ARG4);
                case 5:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1], args[2], args[3], args[4]);
            }
            return ErrorEval.VALUE_INVALID;
        }

        public double Evaluate(double[] ds)
        {
            // All finance functions have 3 to 5 args, first 4 are numbers, last is boolean
            // default for last 2 args are 0.0 and false
            // Text boolean literals are not valid for the last arg

            double arg3 = 0.0;
            double arg4 = 0.0;

            switch (ds.Length)
            {
                case 5:
                    arg4 = ds[4];
                    break;
                case 4:
                    arg3 = ds[3];
                    break;
                case 3:
                    break;
                default:
                    throw new ArgumentException("Wrong number of arguments");
            }
            return Evaluate(ds[0], ds[1], ds[2], arg3, arg4 != 0.0);
        }

        public abstract double Evaluate(double rate, double arg1, double arg2, double arg3, bool type);

        public static readonly Function FV = new Fv();
        public static readonly Function NPER = new Nper();
        public static readonly Function PMT = new Pmt();
        public static readonly Function PV = new Pv();
    }
}