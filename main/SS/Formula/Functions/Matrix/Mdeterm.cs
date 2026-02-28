using MathNet.Numerics.LinearAlgebra;
using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;
using static NPOI.SS.Formula.Functions.MatrixFunction;

namespace NPOI.SS.Formula.Functions
{
    public class Mdeterm:OneArrayArg
    {
        private MutableValueCollector instance = new MutableValueCollector(false, false);
        protected override double[] CollectValues(ValueEval arg)
        {
            double[] values = instance.collectValues(arg);

            /* handle case where MDETERM is operating on an array that that is not completely filled*/
            if (arg is AreaEval && values.Length == 1)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            return values;
        }
        protected override double[,] Evaluate(double[,] d1)
        {
            if (d1.GetLength(0) != d1.GetLength(1))
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            var temp = Matrix<double>.Build.DenseOfArray(d1);
            var result = new double[1, 1];
            result[0, 0] = temp.LU().Determinant;
            return result;
        }
    }
}
