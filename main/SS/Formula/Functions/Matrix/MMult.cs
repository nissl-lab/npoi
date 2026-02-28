using MathNet.Numerics.LinearAlgebra;
using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;
using static NPOI.SS.Formula.Functions.MatrixFunction;

namespace NPOI.SS.Formula.Functions
{
    public class MMulti: TwoArrayArg
    {
        private MutableValueCollector instance = new MutableValueCollector(false, false);
        protected override double[] CollectValues(ValueEval arg)
        {
            double[] values = instance.collectValues(arg);

            /* handle case where MMULT is operating on an array that is not completely filled*/
            if (arg is AreaEval && values.Length == 1)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            return values;
        }
        protected override double[,] Evaluate(double[,] d1, double[,] d2)
        {
            var first= Matrix<double>.Build.DenseOfArray(d1);
            var second = Matrix<double>.Build.DenseOfArray(d2);

            return (first * second).ToArray();
        }
    }
}
