using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public class Trend : Function
    {
        MatrixFunction.MutableValueCollector collector = new MatrixFunction.MutableValueCollector(false, false);
        private class TrendResults
        {
            public double[] vals;
            public int resultWidth;
            public int resultHeight;

            public TrendResults(double[] vals, int resultWidth, int resultHeight)
            {
                this.vals = vals;
                this.resultWidth = resultWidth;
                this.resultHeight = resultHeight;
            }
        }
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            throw new NotImplementedException();
        }
        private static double[,] evalToArray(ValueEval arg)
        {
            double[,] ar;
            ValueEval eval;
            if (arg is MissingArgEval)
            {
                return new double[0,0];
            }
            if (arg is RefEval)
            {
                RefEval re = (RefEval)arg;
                if (re.NumberOfSheets > 1)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                eval = re.GetInnerValueEval(re.FirstSheetIndex);
            }
            else
            {
                eval = arg;
            }
            if (eval == null)
            {
                throw new ArgumentNullException("Parameter may not be null.");
            }

            if (eval is AreaEval)
            {
                AreaEval ae = (AreaEval)eval;
                int w = ae.Width;
                int h = ae.Height;
                ar = new double[h,w];
                for (int i = 0; i < h; i++)
                {
                    for (int j = 0; j < w; j++)
                    {
                        ValueEval ve = ae.GetRelativeValue(i, j);
                        if (!(ve is NumericValueEval))
                        {
                            throw new EvaluationException(ErrorEval.VALUE_INVALID);
                        }
                        ar[i,j] = ((NumericValueEval)ve).NumberValue;
                    }
                }
            }
            else if (eval is NumericValueEval)
            {
                ar = new double[1,1];
                ar[0,0] = ((NumericValueEval)eval).NumberValue;
            }
            else
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }

            return ar;
        }

        private static double[,] getDefaultArrayOneD(int w)
        {
            double[,] array = new double[w,1];
            for (int i = 0; i < w; i++)
            {
                array[i,0] = i + 1;
            }
            return array;
        }

        private static double[] flattenArray(double[,] twoD)
        {
            if (twoD.Length < 1)
            {
                return new double[0];
            }
            double[] oneD = new double[twoD.Length * twoD.GetLength(1)];
            for (int i = 0; i < twoD.Length; i++)
            {
                for (int j = 0; j < twoD.GetLength(1); j++)
                {
                    oneD[i * twoD.GetLength(1) + j] = twoD[i,j];
                }
            }
            return oneD;
        }

        private static double[,] flattenArrayToRow(double[,] twoD)
        {
            if (twoD.Length < 1)
            {
                return new double[0,0];
            }
            double[,] oneD = new double[twoD.Length * twoD[0].Length,1];
            for (int i = 0; i < twoD.Length; i++)
            {
                for (int j = 0; j < twoD[0].Length; j++)
                {
                    oneD[i * twoD[0].Length + j,0] = twoD[i,j];
                }
            }
            return oneD;
        }

        private static double[,] switchRowsColumns(double[,] array)
        {
            double[][] newArray = new double[array[0].Length,array.Length];
            for (int i = 0; i < array.Length; i++)
            {
                for (int j = 0; j < array[0].Length; j++)
                {
                    newArray[j][i] = array[i,j];
                }
            }
            return newArray;
        }

        /**
         * Check if all columns in a matrix contain the same values.
         * Return true if the number of distinct values in each column is 1.
         *
         * @param matrix  column-oriented matrix. A Row matrix should be transposed to column .
         * @return  true if all columns contain the same value
         */
        private static bool isAllColumnsSame(double[,] matrix)
        {
            if (matrix.Length == 0) return false;

            bool[] cols = new bool[matrix[0].Length];
            for (int j = 0; j < matrix[0].Length; j++)
            {
                double prev = Double.NaN;
                for (int i = 0; i < matrix.Length; i++)
                {
                    double v = matrix[i,j];
                    if (i > 0 && v != prev)
                    {
                        cols[j] = true;
                        break;
                    }
                    prev = v;
                }
            }
            bool allEquals = true;
            foreach (bool x in cols)
            {
                if (x)
                {
                    allEquals = false;
                    break;
                }
            };
            return allEquals;

        }

        private static TrendResults GetNewY(ValueEval[] args)
        {
            double[,]   xOrig;
            double[,]  x;
            double[,] yOrig;
            double[] y;
            double[,] newXOrig;
            double[,]
        newX;
            double[,]
        resultSize;
            bool passThroughOrigin = false;
            switch (args.Length)
            {
                case 1:
                    yOrig = evalToArray(args[0]);
                    xOrig = new double[0,0];
                    newXOrig = new double[0,0];
                    break;
                case 2:
                    yOrig = evalToArray(args[0]);
                    xOrig = evalToArray(args[1]);
                    newXOrig = new double[0][0];
                    break;
                case 3:
                    yOrig = evalToArray(args[0]);
                    xOrig = evalToArray(args[1]);
                    newXOrig = evalToArray(args[2]);
                    break;
                case 4:
                    yOrig = evalToArray(args[0]);
                    xOrig = evalToArray(args[1]);
                    newXOrig = evalToArray(args[2]);
                    if (!(args[3] is BoolEval))
                    {
                        throw new EvaluationException(ErrorEval.VALUE_INVALID);
                    }
                    // The argument in Excel is false when it *should* pass through the origin.
                    passThroughOrigin = !((BoolEval)args[3]).BooleanValue;
                    break;
                default:
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }

            if (yOrig.Length < 1)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            y = flattenArray(yOrig);
            newX = newXOrig;

            if (newXOrig.Length > 0)
            {
                resultSize = newXOrig;
            }
            else
            {
                resultSize = new double[1,1];
            }

            if (y.Length == 1)
            {
                /* See comment at top of file
                if (xOrig.Length > 0 && !(xOrig.Length == 1 || xOrig[0].Length == 1)) {
                    throw new EvaluationException(ErrorEval.REF_INVALID);
                } else if (xOrig.Length < 1) {
                    x = new double[1][1];
                    x[0][0] = 1;
                } else {
                    x = new double[1][];
                    x[0] = flattenArray(xOrig);
                    if (newXOrig.Length < 1) {
                        resultSize = xOrig;
                    }
                }*/
                throw new NotImplementedException("Sample size too small");
            }
            else if (yOrig.Length == 1 || yOrig[0].Length == 1)
            {
                if (xOrig.Length < 1)
                {
                    x = getDefaultArrayOneD(y.Length);
                    if (newXOrig.Length < 1)
                    {
                        resultSize = yOrig;
                    }
                }
                else
                {
                    x = xOrig;
                    if (xOrig[0].Length > 1 && yOrig.Length == 1)
                    {
                        x = switchRowsColumns(x);
                    }
                    if (newXOrig.Length < 1)
                    {
                        resultSize = xOrig;
                    }
                }
                if (newXOrig.Length > 0 && (x.Length == 1 || x[0].Length == 1))
                {
                    newX = flattenArrayToRow(newXOrig);
                }
            }
            else
            {
                if (xOrig.Length < 1)
                {
                    x = getDefaultArrayOneD(y.Length);
                    if (newXOrig.Length < 1)
                    {
                        resultSize = yOrig;
                    }
                }
                else
                {
                    x = flattenArrayToRow(xOrig);
                    if (newXOrig.Length < 1)
                    {
                        resultSize = xOrig;
                    }
                }
                if (newXOrig.Length > 0)
                {
                    newX = flattenArrayToRow(newXOrig);
                }
                if (y.Length != x.Length || yOrig.Length != xOrig.Length)
                {
                    throw new EvaluationException(ErrorEval.REF_INVALID);
                }
            }

            if (newXOrig.Length < 1)
            {
                newX = x;
            }
            else if (newXOrig.Length == 1 && newXOrig[0].Length > 1 && xOrig.Length > 1 && xOrig[0].Length == 1)
            {
                newX = switchRowsColumns(newXOrig);
            }

            if (newX[0].Length != x[0].Length)
            {
                throw new EvaluationException(ErrorEval.REF_INVALID);
            }

            if (x[0].Length >= x.Length)
            {
                /* See comment at top of file */
                throw new NotImplementedException("Sample size too small");
            }

            int resultHeight = resultSize.Length;
            int resultWidth = resultSize[0].Length;

            if (isAllColumnsSame(x))
            {
                double[] result = new double[newX.Length];
                double avg = Arrays.stream(y).average().orElse(0);
                for (int i = 0; i < result.Length; i++) result[i] = avg;
                return new TrendResults(result, resultWidth, resultHeight);
            }

            OLSMultipleLinearRegression reg = new OLSMultipleLinearRegression();
            if (passThroughOrigin)
            {
                reg.setNoIntercept(true);
            }

            try
            {
                reg.newSampleData(y, x);
            }
            catch (ArgumentException e)
            {
                throw new EvaluationException(ErrorEval.REF_INVALID);
            }
            double[] par;
            try
            {
                par = reg.estimateRegressionParameters();
            }
            catch (SingularMatrixException e)
            {
                throw new NotImplementedException("Singular matrix in input");
            }

            double[] result = new double[newX.Length];
            for (int i = 0; i < newX.Length; i++)
            {
                result[i] = 0;
                if (passThroughOrigin)
                {
                    for (int j = 0; j < par.Length; j++)
                    {
                        result[i] += par[j] * newX[i][j];
                    }
                }
                else
                {
                    result[i] = par[0];
                    for (int j = 1; j < par.Length; j++)
                    {
                        result[i] += par[j] * newX[i][j - 1];
                    }
                }
            }
            return new TrendResults(result, resultWidth, resultHeight);
        }
    }
}
