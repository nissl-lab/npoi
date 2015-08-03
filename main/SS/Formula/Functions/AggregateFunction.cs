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

using NPOI.SS.Formula.Functions;
namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;

    public class AVEDEV : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return StatsLib.avedev(values);
        }
    }
    public class AVERAGE : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            if (values.Length < 1)
            {
                throw new EvaluationException(ErrorEval.DIV_ZERO);
            }
            return MathX.Average(values);
        }
    }
    public class DEVSQ : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return StatsLib.devsq(values);
        }
    }
    public class SUM : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return MathX.Sum(values);
        }
    }
    public class LARGE : AggregateFunction
    {
        protected internal override double Evaluate(double[] ops)
        {
            if (ops.Length < 2)
            {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
            double[] values = new double[ops.Length - 1];
            int k = (int)ops[ops.Length - 1];
            System.Array.Copy(ops, 0, values, 0, values.Length);
            return StatsLib.kthLargest(values, k);
        }
    }
    public class MAX : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return values.Length > 0 ? MathX.Max(values) : 0;
        }
    }
    public class MIN : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return values.Length > 0 ? MathX.Min(values) : 0;
        }
    }
    public class MEDIAN : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return StatsLib.median(values);
        }
    }
    public class PRODUCT : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return MathX.Product(values);
        }
    }
    public class SMALL : AggregateFunction
    {
        protected internal override double Evaluate(double[] ops)
        {
            if (ops.Length < 2)
            {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
            double[] values = new double[ops.Length - 1];
            int k = (int)ops[ops.Length - 1];
            System.Array.Copy(ops, 0, values, 0, values.Length);
            return StatsLib.kthSmallest(values, k);
        }
    }
    public class STDEV : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            if (values.Length < 1)
            {
                throw new EvaluationException(ErrorEval.DIV_ZERO);
            }
            return StatsLib.stdev(values);
        }
    }
    public class SUMSQ : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return MathX.Sumsq(values);
        }
    }
    public class VAR : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            if (values.Length < 1)
            {
                throw new EvaluationException(ErrorEval.DIV_ZERO);
            }
            return StatsLib.var(values);
        }
    };
    public class VARP : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            if (values.Length < 1)
            {
                throw new EvaluationException(ErrorEval.DIV_ZERO);
            }
            return StatsLib.varp(values);
        }
    };

    public class SubtotalInstance : AggregateFunction
    {
        private AggregateFunction _func;
        public SubtotalInstance(AggregateFunction func)
        {
            _func = func;
        }

        protected internal override double Evaluate(double[] values)
        {
            return _func.Evaluate(values);
        }
        /**
                 *  ignore nested subtotals.
                 */
        public override bool IsSubtotalCounted
        {
            get
            {
                return false;
            }
        }
    }


    public class LargeSmall : Fixed2ArgFunction
    {
        private bool _isLarge;
        protected LargeSmall(bool isLarge)
        {
            _isLarge = isLarge;
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0,
                ValueEval arg1)
        {
            double dn;
            try
            {
                ValueEval ve1 = OperandResolver.GetSingleValue(arg1, srcRowIndex, srcColumnIndex);
                dn = OperandResolver.CoerceValueToDouble(ve1);
            }
            catch (EvaluationException)
            {
                // all errors in the second arg translate to #VALUE!
                return ErrorEval.VALUE_INVALID;
            }
            // weird Excel behaviour on second arg
            if (dn < 1.0)
            {
                // values between 0.0 and 1.0 result in #NUM!
                return ErrorEval.NUM_ERROR;
            }
            // all other values are rounded up to the next integer
            int k = (int)Math.Ceiling(dn);

            double result;
            try
            {
                double[] ds = NPOI.SS.Formula.Functions.AggregateFunction.ValueCollector.CollectValues(arg0);
                if (k > ds.Length)
                {
                    return ErrorEval.NUM_ERROR;
                }
                result = _isLarge ? StatsLib.kthLargest(ds, k) : StatsLib.kthSmallest(ds, k);
                NumericFunction.CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return new NumberEval(result);
        }
    }

    /**
     *  Returns the k-th percentile of values in a range. You can use this function to establish a threshold of
     *  acceptance. For example, you can decide to examine candidates who score above the 90th percentile.
     *
     *  PERCENTILE(array,k)
     *  Array     is the array or range of data that defines relative standing.
     *  K     is the percentile value in the range 0..1, inclusive.
     *
     * <strong>Remarks</strong>
     * <ul>
     *     <li>if array is empty or Contains more than 8,191 data points, PERCENTILE returns the #NUM! error value.</li>
     *     <li>If k is nonnumeric, PERCENTILE returns the #VALUE! error value.</li>
     *     <li>If k is &lt; 0 or if k &gt; 1, PERCENTILE returns the #NUM! error value.</li>
     *     <li>If k is not a multiple of 1/(n - 1), PERCENTILE interpolates to determine the value at the k-th percentile.</li>
     * </ul>
     */
    public class Percentile : Fixed2ArgFunction
    {
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0,
                ValueEval arg1)
        {
            double dn;
            try
            {
                ValueEval ve1 = OperandResolver.GetSingleValue(arg1, srcRowIndex, srcColumnIndex);
                dn = OperandResolver.CoerceValueToDouble(ve1);
            }
            catch (EvaluationException)
            {
                // all errors in the second arg translate to #VALUE!
                return ErrorEval.VALUE_INVALID;
            }
            if (dn < 0 || dn > 1)
            { // has to be percentage
                return ErrorEval.NUM_ERROR;
            }

            double result;
            try
            {
                double[] ds = NPOI.SS.Formula.Functions.AggregateFunction.ValueCollector.CollectValues(arg0);
                int N = ds.Length;

                if (N == 0 || N > 8191)
                {
                    return ErrorEval.NUM_ERROR;
                }

                double n = (N - 1) * dn + 1;
                if (n == 1d)
                {
                    result = StatsLib.kthSmallest(ds, 1);
                }
                else if (n == N)
                {
                    result = StatsLib.kthLargest(ds, 1);
                }
                else
                {
                    int k = (int)n;
                    double d = n - k;
                    result = StatsLib.kthSmallest(ds, k) + d
                            * (StatsLib.kthSmallest(ds, k + 1) - StatsLib.kthSmallest(ds, k));
                }

                NumericFunction.CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return new NumberEval(result);
        }
    }


    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    public abstract class AggregateFunction : MultiOperandNumericFunction
    {
        public static Function SubtotalInstance(Function func)
        {
            AggregateFunction arg = (AggregateFunction)func;
            return new SubtotalInstance(arg);
        }
        internal class ValueCollector : MultiOperandNumericFunction
        {
            private static ValueCollector instance = new ValueCollector();
            public ValueCollector() :
                base(false, false)
            {
            }
            public static double[] CollectValues(params ValueEval[] operands)
            {
                return instance.GetNumberArray(operands);
            }
            protected internal override double Evaluate(double[] values)
            {
                throw new InvalidOperationException("should not be called");
            }
        }
        protected AggregateFunction()
            : base(false, false)
        {

        }

        public static readonly Function AVEDEV = new AVEDEV();
        public static readonly Function AVERAGE = new AVERAGE();
        public static readonly Function DEVSQ = new DEVSQ();
        public static readonly Function LARGE = new LARGE();
        public static readonly Function MAX = new MAX();
        public static readonly Function MEDIAN = new MEDIAN();
        public static readonly Function MIN = new MIN();
        public static readonly Function PRODUCT = new PRODUCT();
        public static readonly Function SMALL = new SMALL();
        public static readonly Function STDEV = new STDEV();
        public static readonly Function SUM = new SUM();
        public static readonly Function SUMSQ = new SUMSQ();
        public static readonly Function VAR = new VAR();
        public static readonly Function VARP = new VARP();
        public static readonly Function PERCENTILE = new Percentile();
    }
}
