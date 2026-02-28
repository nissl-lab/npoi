using System;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Functions
{
    /// <summary>
    /// The Forecast class is a representation of the Excel FORECAST function.
    /// This function predicts a future value along a linear trend line based on existing historical data.
    /// The class inherits from the Fixed3ArgFunction class and overrides the Evaluate method.
    /// The Evaluate method takes three arguments: the x-value for which we want to forecast a y-value, 
    /// and two arrays of x-values and y-values representing historical data.
    /// The method calculates the slope and intercept of the line of best fit for the historical data 
    /// and uses these to calculate the forecast y-value.
    /// The class also includes methods for converting ValueEval objects to numeric arrays and for creating ValueVectors.
    /// </summary>
    public class Forecast : Fixed3ArgFunction
    {
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
            ValueEval arg2)
        {
            try
            {
                if(arg0 is ErrorEval arg0Error)
                {
                    return arg0Error;
                }

                if(arg1 is ErrorEval arg1Error)
                {
                    return arg1Error;
                }

                if(arg2 is ErrorEval arg2Error)
                {
                    return arg2Error;
                }

                double x = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double[] yValues = GetNumericArray(arg1);
                double[] xValues = GetNumericArray(arg2);

                if(yValues.Length != xValues.Length)
                {
                    return ErrorEval.NA;
                }

                double xSum = 0, ySum = 0, xySum = 0, xSquareSum = 0;
                int n = xValues.Length;

                for(int i = 0; i < n; i++)
                {
                    xSum += xValues[i];
                    ySum += yValues[i];
                    xySum += xValues[i] * yValues[i];
                    xSquareSum += Math.Pow(xValues[i], 2);
                }

                double slope = (n * xySum - xSum * ySum) / (n * xSquareSum - Math.Pow(xSum, 2));
                double intercept = (ySum - slope * xSum) / n;

                double forecastY = slope * x + intercept;

                return new NumberEval(forecastY);
            }
            catch(EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        private static double[] GetNumericArray(ValueEval arg)
        {
            ValueVector vv = CreateValueVector(arg);
            double[] result = new double[vv.Size];
            for(int i = 0; i < vv.Size; i++)
            {
                ValueEval v = vv.GetItem(i);
                if(v is ErrorEval errorEval)
                {
                    throw new EvaluationException(errorEval);
                }

                if(v is NumberEval numberEval)
                {
                    result[i] = numberEval.NumberValue;
                }
            }

            return result;
        }

        private static ValueVector CreateValueVector(ValueEval arg)
        {
            return arg switch {
                ErrorEval eval => throw new EvaluationException(eval),
                TwoDEval dEval => new AreaValueArray(dEval),
                RefEval refEval => new RefValueArray(refEval),
                _ => new SingleCellValueArray(arg)
            };
        }

        private abstract class ValueArray(int size) : ValueVector
        {
            public ValueEval GetItem(int index)
            {
                if(index < 0 || index > size)
                {
                    throw new ArgumentException($"Specified index {index} is outside range (0..{(size - 1)})");
                }

                return GetItemInternal(index);
            }

            protected abstract ValueEval GetItemInternal(int index);

            public int Size => size;
        }

        private sealed class SingleCellValueArray(ValueEval value) : ValueArray(1)
        {
            protected override ValueEval GetItemInternal(int index)
            {
                return value;
            }
        }

        private sealed class RefValueArray(RefEval ref1) : ValueArray(ref1.NumberOfSheets)
        {
            private readonly int _width = ref1.NumberOfSheets;

            protected override ValueEval GetItemInternal(int index)
            {
                int sIx = (index % _width) + ref1.FirstSheetIndex;
                return ref1.GetInnerValueEval(sIx);
            }
        }

        private sealed class AreaValueArray(TwoDEval ae) : ValueArray(ae.Width * ae.Height)
        {
            private readonly int _width = ae.Width;

            protected override ValueEval GetItemInternal(int index)
            {
                int rowIx = index / _width;
                int colIx = index % _width;
                return ae.GetValue(rowIx, colIx);
            }
        }
    }
}