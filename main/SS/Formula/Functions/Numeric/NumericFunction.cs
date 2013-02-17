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
/*
 * Created on May 14, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;

    public abstract class OneArg : Fixed1ArgFunction
    {
        protected OneArg()
        {
            // no fields to initialise
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            double result;
            try
            {
                double d = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                result = Evaluate(d);
                NumericFunction.CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }
        protected double Eval(ValueEval[] args, int srcCellRow, short srcCellCol)
        {
            if (args.Length != 1)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            double d = NumericFunction.SingleOperandEvaluate(args[0], srcCellRow, srcCellCol);
            return Evaluate(d);
        }
        public abstract double Evaluate(double d);
    }

    public abstract class TwoArg : Fixed2ArgFunction
    {
        protected TwoArg()
        {
            // no fields to initialise
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                double d0 = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                result = Evaluate(d0, d1);
                NumericFunction.CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }
        public abstract double Evaluate(double d0, double d1);
    }

    /**
     * @author Amol S. Deshmukh &lt; amolweb at yahoo dot com &gt;
     *
     */
    public abstract class NumericFunction : Function
    {
        public const double TEN = 10.0;
        public static readonly double LOG_10_TO_BASE_e = Math.Log(TEN);
        public const double E = Math.E;
        public const double PI = Math.PI;
        public const double ZERO = 0.0;

        public static double SingleOperandEvaluate(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            double result = OperandResolver.CoerceValueToDouble(ve);

            CheckValue(result);
            return result;
        }

        public static void CheckValue(double result)
        {
            if (Double.IsNaN(result) || Double.IsInfinity(result))
            {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
        }

        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            double result;
            try
            {
                result = Eval(args, srcCellRow, srcCellCol);
                CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }
        protected abstract double Eval(ValueEval[] evals, int srcCellRow, int srcCellCol);

        public static readonly Function ABS = new Abs();

        public static readonly Function COS = new Cos();
        public static readonly Function COSH = new Cosh();
        public static readonly Function ACOS = new Acos();
        public static readonly Function ACOSH = new Acosh();

        public static readonly Function ASIN = new Asin();
        public static readonly Function ASINH = new Asinh();
        public static readonly Function SIN = new Sin();
        public static readonly Function SINH = new Sinh();

        public static readonly Function TAN = new Tan();
        public static readonly Function TANH = new Tanh();
        public static readonly Function ATAN = new Atan();
        public static readonly Function ATANH = new Atanh();
        public static readonly Function ATAN2 = new Atan2();

        public static readonly Function DEGREES = new Degrees();
        public static readonly Function DOLLAR = new Dollar();
        public static readonly Function EXP = new Exp();
        public static readonly Function FACT = new Fact();
        public static readonly Function INT = new Int();
        public static readonly Function LN = new Ln();
        public static readonly Function LOG10 = new Log10();

        public static readonly Function RADIANS = new Radians();
        public static readonly Function SIGN = new Sign();
        public static readonly Function SQRT = new Sqrt();

        public static readonly Function CEILING = new Ceiling();
        public static readonly Function COMBIN =new Combin();
        public static readonly Function FLOOR = new Floor();
        public static readonly Function MOD = new Mod();
        public static readonly Function POWER = new Power();
        public static readonly Function ROUND = new Round();
        public static readonly Function ROUNDDOWN = new Rounddown();
        public static readonly Function ROUNDUP = new Roundup();

        public static readonly Function LOG = new Log();
        public static readonly Function TRUNC = new Trunc();
        public static readonly Function POISSON = new Poisson();
    }
}