/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using System;


    /**
     * Base class for linear regression functions.
     *
     * Calculates the linear regression line that is used to predict y values from x values<br/>
     * (http://introcs.cs.princeton.edu/java/97data/LinearRegression.java.html)
     * <b>Syntax</b>:<br/>
     * <b>INTERCEPT</b>(<b>arrayX</b>, <b>arrayY</b>)<p/>
     * or
     * <b>SLOPE</b>(<b>arrayX</b>, <b>arrayY</b>)<p/>
     *
     *
     * @author Johan Karlsteen
     */
    public class LinearRegressionFunction : Fixed2ArgFunction
    {

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

        public enum FUNCTION { INTERCEPT, SLOPE };
        public FUNCTION function;

        public LinearRegressionFunction(FUNCTION function)
        {
            this.function = function;
        }


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex,
                ValueEval arg0, ValueEval arg1)
        {
            double result;
            try
            {
                ValueVector vvY = CreateValueVector(arg0);
                ValueVector vvX = CreateValueVector(arg1);
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

        private double EvaluateInternal(ValueVector x, ValueVector y, int size)
        {

            // error handling is as if the x is fully Evaluated before y
            ErrorEval firstXerr = null;
            ErrorEval firstYerr = null;
            bool accumlatedSome = false;
            // first pass: read in data, compute xbar and ybar
            double sumx = 0.0, sumy = 0.0;

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
                    sumx += nx.NumberValue;
                    sumy += ny.NumberValue;
                }
                else
                {
                    // all other combinations of value types are silently ignored
                }
            }
            double xbar = sumx / size;
            double ybar = sumy / size;

            // second pass: compute summary statistics
            double xxbar = 0.0, xybar = 0.0;
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
                    NumberEval nx = (NumberEval)vx;
                    NumberEval ny = (NumberEval)vy;
                    xxbar += (nx.NumberValue - xbar) * (nx.NumberValue - xbar);
                    xybar += (nx.NumberValue - xbar) * (ny.NumberValue - ybar);
                }
                else
                {
                    // all other combinations of value types are silently ignored
                }
            }
            double beta1 = xybar / xxbar;
            double beta0 = ybar - beta1 * xbar;

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

            if (function == FUNCTION.INTERCEPT)
            {
                return beta0;
            }
            else
            {
                return beta1;
            }
        }

        private ValueVector CreateValueVector(ValueEval arg)
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