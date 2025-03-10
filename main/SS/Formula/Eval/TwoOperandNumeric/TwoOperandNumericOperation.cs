using System;
using NPOI.SS.Formula.Functions;

namespace NPOI.SS.Formula.Eval
{

    public abstract class TwoOperandNumericOperation : Fixed2ArgFunction, ArrayFunction
    {
        protected double SingleOperandEvaluate(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, srcCellCol);
            return OperandResolver.CoerceValueToDouble(ve);
        }
        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }
            Func<double, double, double> func = this.Evaluate;
            return new ArrayEval(func).Evaluate(srcRowIndex, srcColumnIndex, args[0], args[1]);
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                double d0 = SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                result = Evaluate(d0, d1);
                if (result == 0.0)
                { // this '==' matches +0.0 and -0.0
                    // Excel Converts -0.0 to +0.0 for '*', '/', '%', '+' and '^'
                    if (!(this is SubtractEval))
                    {
                        return NumberEval.ZERO;
                    }
                }
                if (Double.IsNaN(result) || Double.IsInfinity(result))
                {
                    return ErrorEval.NUM_ERROR;
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }

        public abstract double Evaluate(double d0, double d1);

        public static NPOI.SS.Formula.Functions.Function AddEval = new AddEval();
        public static NPOI.SS.Formula.Functions.Function DivideEval = new DivideEval();
        public static NPOI.SS.Formula.Functions.Function MultiplyEval = new MultiplyEval();
        public static NPOI.SS.Formula.Functions.Function PowerEval = new PowerEval();
        public static NPOI.SS.Formula.Functions.Function SubtractEval = new SubtractEval();

        private sealed class ArrayEval : MatrixFunction.TwoArrayArg
        {
            readonly Func<double, double, double> _evaluateFunc = null;
            public ArrayEval(Func<double, double, double> evalFunc)
            {
                _evaluateFunc = evalFunc;
            }

            private readonly MatrixFunction.MutableValueCollector instance = new MatrixFunction.MutableValueCollector(false, true);

            protected override double[] CollectValues(ValueEval arg)
            {
                return instance.collectValues(arg);
            }

            protected override double[,] Evaluate(double[,] d1, double[,] d2)
            {
                int width = d1.GetLength(1) < d2.GetLength(1) ? d1.GetLength(1) : d2.GetLength(1);
                int height = (d1.GetLength(0) < d2.GetLength(0)) ? d1.GetLength(0) : d2.GetLength(0);

                double[,] result = new double[height, width];

                for (int j = 0; j < height; j++) {
                    for (int i = 0; i < width; i++) {
                        result[j, i] = _evaluateFunc(d1[j, i], d2[j, i]);

                    }
                }
                return result;
            }
        }
    }
}