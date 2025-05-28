using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class MatrixFunction
    {
        public static Function MDETERM = new Mdeterm();
        public static Function TRANSPOSE = new Transpose();
        public static Function MMULT = new MMulti();
        public static Function MINVERSE = new Minverse();

        /* retrieves 1D array from 2D array after calculations */
        private static double[] extractDoubleArray(double[,] matrix)
        {
            int idx = 0;

            if (matrix == null || matrix.GetLength(0) < 1 || matrix.GetLength(1) < 1)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }

            double[] vector = new double[matrix.GetLength(0) * matrix.GetLength(1)];

            for (int j = 0; j < matrix.GetLength(0); j++)
            {
                for (int i = 0; i < matrix.GetLength(1); i++)
                {
                    vector[idx++] = matrix[j,i];
                }
            }
            return vector;
        }
        public static void CheckValues(double[] results)
        {
            for (int idx = 0; idx < results.Length; idx++)
            {
                if (Double.IsNaN(results[idx]) || Double.IsInfinity(results[idx]))
                {
                    throw new EvaluationException(ErrorEval.NUM_ERROR);
                }
            }
        }
        /* converts 1D array to 2D array for calculations */
        private static double[,] fillDoubleArray(double[] vector, int rows, int cols)
        {
            int i = 0, j = 0;

            if (rows < 1 || cols < 1 || vector.Length < 1)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }

            double[,] matrix = new double[rows, cols];

            for (int idx = 0; idx < vector.Length; idx++)
            {
                if (j < matrix.GetLength(0))
                {
                    if (i == matrix.GetLength(1))
                    {
                        i = 0;
                        j++;
                    }
                    matrix[j, i++] = vector[idx];
                }
            }

            return matrix;
        }
        public abstract class OneArrayArg : Fixed1ArgFunction
        {
            protected OneArrayArg()
            {
                //no fields to initialize
            }

            public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
            {
                if (arg0 is AreaEval eval)
                {
                    double[] result = null;
                    double[,] resultArray;
                    int width = 1, height = 1;

                    try
                    {
                        double[] values = CollectValues(eval);
                        double[,] array = fillDoubleArray(values, eval.Height, eval.Width);
                        resultArray = Evaluate(array);
                        width = resultArray.GetLength(1);
                        height = resultArray.GetLength(0);
                        result = extractDoubleArray(resultArray);

                        CheckValues(result);
                    }
                    catch (EvaluationException e)
                    {
                        return e.GetErrorEval();
                    }

                    ValueEval[] vals = new ValueEval[result.Length];

                    for (int idx = 0; idx < result.Length; idx++)
                    {
                        vals[idx] = new NumberEval(result[idx]);
                    }

                    if (result.Length == 1)
                    {
                        return vals[0];
                    }
                    else
                    {
                        /* find a better solution */
                        return new CacheAreaEval(eval.FirstRow, eval.FirstColumn,
                                                eval.FirstRow + height - 1,
                                                eval.FirstColumn + width - 1, vals);
                    }
                }
                else
                {
                    double[,] result = null;
                    try
                    {
                        double value = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                        double[,] temp = { { value } };
                        result = Evaluate(temp);
                        NumericFunction.CheckValue(result[0, 0]);
                    }
                    catch (EvaluationException e)
                    {
                        return e.GetErrorEval();
                    }

                    return new NumberEval(result[0, 0]);
                }
            }

            protected abstract double[,] Evaluate(double[,] d1);
            protected abstract double[] CollectValues(ValueEval arg);
        }
    public abstract class TwoArrayArg : Fixed2ArgFunction
        {
            protected TwoArrayArg()
            {
                //no fields to initialize
            }


            public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
            {
                double[] result;
                int width = 1, height = 1;

                try
                {
                    double[,] array0, array1, resultArray;

                    if (arg0 is AreaEval eval)
                    {
                        try
                        {
                            double[] values = CollectValues(eval);
                            array0 = fillDoubleArray(values, eval.Height, eval.Width);
                        }
                        catch (EvaluationException e)
                        {
                            return e.GetErrorEval();
                        }
                    }
                    else
                    {
                        try
                        {
                            double value = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                            array0 = new double[,] { { value } };
                        }
                        catch (EvaluationException e)
                        {
                            return e.GetErrorEval();
                        }
                    }

                    if (arg1 is AreaEval areaEval)
                    {
                        try
                        {
                            double[] values = CollectValues(areaEval);
                            array1 = fillDoubleArray(values, areaEval.Height, areaEval.Width);
                        }
                        catch (EvaluationException e)
                        {
                            return e.GetErrorEval();
                        }
                    }
                    else
                    {
                        try
                        {
                            double value = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                            array1 = new double[,] { { value } };
                        }
                        catch (EvaluationException e)
                        {
                            return e.GetErrorEval();
                        }
                    }

                    resultArray = Evaluate(array0, array1);
                    width = resultArray.GetLength(1);
                    height = resultArray.GetLength(0);
                    result = extractDoubleArray(resultArray);
                    CheckValues(result);
                }
                catch (EvaluationException e)
                {
                    return e.GetErrorEval();
                }
                catch (ArgumentException )
                {
                    return ErrorEval.VALUE_INVALID;
                }


                ValueEval[] vals = new ValueEval[result.Length];

                for (int idx = 0; idx < result.Length; idx++)
                {
                    vals[idx] = new NumberEval(result[idx]);
                }

                if (result.Length == 1)
                    return vals[0];
                else
                {
                    return new CacheAreaEval(((AreaEval)arg0).FirstRow, ((AreaEval)arg0).FirstColumn,
                            ((AreaEval)arg0).FirstRow + height - 1,
                            ((AreaEval)arg0).FirstColumn + width - 1, vals);
                }

            }

            protected abstract double[,] Evaluate(double[,] d1, double[,] d2);
            protected abstract double[] CollectValues(ValueEval arg);

        }


    public class MutableValueCollector : MultiOperandNumericFunction
        {
            public MutableValueCollector(bool isReferenceBoolCounted, bool isBlankCounted) :
                base(isReferenceBoolCounted, isBlankCounted)
            {

            }
            public double[] collectValues(params ValueEval[] operands)
            {
                return GetNumberArray(operands);
            }
            protected internal override double Evaluate(double[] values)
            {
                throw new InvalidOperationException("should not be called");
            }
        }
    }
}
