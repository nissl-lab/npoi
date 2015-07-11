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
    using NPOI.SS.Formula;

    public interface Accumulator
    {
        double Accumulate(double x, double y);
    }
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    public abstract class XYNumericFunction : Fixed2ArgFunction
    {
        protected static int X = 0;
        protected static int Y = 1;

        private abstract class ValueArray : ValueVector
        {
            private int _size;
            protected ValueArray(int size)
            {
                _size = size;
            }
            public ValueEval GetItem(int index)
            {
                if (index < 0 || index > _size)
                {
                    throw new ArgumentException("Specified index " + index
                            + " is outside range (0.." + (_size - 1) + ")");
                }
                return GetItemInternal(index);
            }
            protected abstract ValueEval GetItemInternal(int index);
            public int Size
            {
                get
                {
                    return _size;
                }
            }
        }

        private class SingleCellValueArray : ValueArray
        {
            private ValueEval _value;
            public SingleCellValueArray(ValueEval value)
                : base(1)
            {

                _value = value;
            }
            protected override ValueEval GetItemInternal(int index)
            {
                return _value;
            }
        }
        private class RefValueArray : ValueArray
        {
            private RefEval _ref;
            private int _width;

            public RefValueArray(RefEval ref1)
                : base(ref1.NumberOfSheets)
            {
                _ref = ref1;
                _width = ref1.NumberOfSheets;
            }
            protected override ValueEval GetItemInternal(int index)
            {
                int sIx = (index % _width) + _ref.FirstSheetIndex;
                return _ref.GetInnerValueEval(sIx);
            }
        }
        private class AreaValueArray : ValueArray
        {
            private TwoDEval _ae;
            private int _width;

            public AreaValueArray(TwoDEval ae)
                : base(ae.Width * ae.Height)
            {
                _ae = ae;
                _width = ae.Width;
            }
            protected override ValueEval GetItemInternal(int index)
            {
                int rowIx = index / _width;
                int colIx = index % _width;
                return _ae.GetValue(rowIx, colIx);
            }
        }
        protected class DoubleArrayPair
        {

            private double[] _xArray;
            private double[] _yArray;

            public DoubleArrayPair(double[] xArray, double[] yArray)
            {
                _xArray = xArray;
                _yArray = yArray;
            }
            public double[] GetXArray()
            {
                return _xArray;
            }
            public double[] GetYArray()
            {
                return _yArray;
            }
        }


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                ValueVector vvX = CreateValueVector(arg0);
                ValueVector vvY = CreateValueVector(arg1);
                int size = vvX.Size;
                if (size == 0 || vvY.Size != size)
                {
                    return ErrorEval.NA;
                }
                result = EvaluateInternal(vvX, vvY, size);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            if (Double.IsNaN(result) || Double.IsInfinity(result))
            {
                return ErrorEval.NUM_ERROR;
            }
            return new NumberEval(result);
        }
        /**
 * Constructs a new instance of the Accumulator used to calculated this function
 */
        public abstract Accumulator CreateAccumulator();

        private double EvaluateInternal(ValueVector x, ValueVector y, int size)
        {
            Accumulator acc = CreateAccumulator();

            // error handling is as if the x is fully evaluated before y
            ErrorEval firstXerr = null;
            ErrorEval firstYerr = null;
            bool accumlatedSome = false;
            double result = 0.0;

            for (int i = 0; i < size; i++)
            {
                ValueEval vx = x.GetItem(i);
                ValueEval vy = y.GetItem(i);
                if (vx is ErrorEval)
                {
                    if (firstXerr == null)
                    {
                        firstXerr = (ErrorEval)vx;
                        continue;
                    }
                }
                if (vy is ErrorEval)
                {
                    if (firstYerr == null)
                    {
                        firstYerr = (ErrorEval)vy;
                        continue;
                    }
                }
                // only count pairs if both elements are numbers
                if (vx is NumberEval && vy is NumberEval)
                {
                    accumlatedSome = true;
                    NumberEval nx = (NumberEval)vx;
                    NumberEval ny = (NumberEval)vy;
                    result += acc.Accumulate(nx.NumberValue, ny.NumberValue);
                }
                else
                {
                    // all other combinations of value types are silently ignored
                }
            }
            if (firstXerr != null)
            {
                throw new EvaluationException(firstXerr);
            }
            if (firstYerr != null)
            {
                throw new EvaluationException(firstYerr);
            }
            if (!accumlatedSome)
            {
                throw new EvaluationException(ErrorEval.DIV_ZERO);
            }
            return result;
        }

        private static double[] TrimToSize(double[] arr, int len)
        {
            double[] tarr = arr;
            if (arr.Length > len)
            {
                tarr = new double[len];
                Array.Copy(arr, 0, tarr, 0, len);
            }
            return tarr;
        }

        private static bool IsNumberEval(Eval eval)
        {
            //bool retval = false;

            //if (eval is NumberEval)
            //{
            //    retval = true;
            //}
            //else if (eval is RefEval)
            //{
            //    RefEval re = (RefEval)eval;
            //    ValueEval ve = re.InnerValueEval;
            //    retval = (ve is NumberEval);
            //}

            //return retval;
            throw new InvalidOperationException("not found in poi");
        }

        private static double GetDoubleValue(Eval eval)
        {
            //double retval = 0;
            //if (eval is NumberEval)
            //{
            //    NumberEval ne = (NumberEval)eval;
            //    retval = ne.NumberValue;
            //}
            //else if (eval is RefEval)
            //{
            //    RefEval re = (RefEval)eval;
            //    ValueEval ve = re.InnerValueEval;
            //    retval = (ve is NumberEval)
            //        ? ((NumberEval)ve).NumberValue
            //        : double.NaN;
            //}
            //else if (eval is ErrorEval)
            //{
            //    retval = double.NaN;
            //}
            //return retval;
            throw new InvalidOperationException("not found in poi");
        }
        private static ValueVector CreateValueVector(ValueEval arg)
        {
            if (arg is ErrorEval)
            {
                throw new EvaluationException((ErrorEval)arg);
            }
            if (arg is TwoDEval)
            {
                return new AreaValueArray((TwoDEval)arg);
            }
            if (arg is RefEval)
            {
                return new RefValueArray((RefEval)arg);
            }
            return new SingleCellValueArray(arg);
        }
    }
}