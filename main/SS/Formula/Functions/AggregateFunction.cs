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
            return MathX.average(values);
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
            return MathX.sum(values);
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
            return values.Length > 0 ? MathX.max(values) : 0;
        }
    }
    public class MIN : AggregateFunction
    {
        protected internal override double Evaluate(double[] values)
        {
            return values.Length > 0 ? MathX.min(values) : 0;
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
            return MathX.product(values);
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
            return MathX.sumsq(values);
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
    }
}
