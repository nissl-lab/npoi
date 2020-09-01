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

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * This Is the base class for all excel function evaluator
     * classes that take variable number of operands, and
     * where the order of operands does not matter
     */
    public abstract class MultiOperandNumericFunction : Function
    {
        static double[] EMPTY_DOUBLE_ARRAY = { };
        private bool _isReferenceBoolCounted;
        private bool _isBlankCounted;

        protected MultiOperandNumericFunction(bool isReferenceBoolCounted, bool isBlankCounted)
        {
            _isReferenceBoolCounted = isReferenceBoolCounted;
            _isBlankCounted = isBlankCounted;
        }
        protected internal abstract double Evaluate(double[] values);

        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {

            double d;
            try
            {
                double[] values = GetNumberArray(args);
                d = Evaluate(values);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            if (Double.IsNaN(d) || Double.IsInfinity(d))
                return ErrorEval.NUM_ERROR;

            return new NumberEval(d);
        }

        private class DoubleList
        {
            private double[] _array;
            private int _Count;

            public DoubleList()
            {
                _array = new double[8];
                _Count = 0;
            }

            public double[] ToArray()
            {
                if (_Count < 1)
                {
                    return EMPTY_DOUBLE_ARRAY;
                }
                double[] result = new double[_Count];
                Array.Copy(_array, 0, result, 0, _Count);
                return result;
            }

            public void Add(double[] values)
            {
                int AddLen = values.Length;
                EnsureCapacity(_Count + AddLen);
                Array.Copy(values, 0, _array, _Count, AddLen);
                _Count += AddLen;
            }

            private void EnsureCapacity(int reqSize)
            {
                if (reqSize > _array.Length)
                {
                    int newSize = reqSize * 3 / 2; // grow with 50% extra
                    double[] newArr = new double[newSize];
                    Array.Copy(_array, 0, newArr, 0, _Count);
                    _array = newArr;
                }
            }

            public void Add(double value)
            {
                EnsureCapacity(_Count + 1);
                _array[_Count] = value;
                _Count++;
            }
        }

        private const int DEFAULT_MAX_NUM_OPERANDS = 30;

        /**
         * Maximum number of operands accepted by this function.
         * Subclasses may override to Change default value.
         */
        protected virtual int MaxNumOperands
        {
            get
            {
                return DEFAULT_MAX_NUM_OPERANDS;
            }
        }
        /**
     *  Whether to count nested subtotals.
     */
        public virtual bool IsSubtotalCounted
        {
            get
            {
                return true;
            }
        }
        /**
     * Collects values from a single argument
     */
        private void CollectValues(ValueEval operand, DoubleList temp)
        {
            if (operand is ThreeDEval)
            {
                ThreeDEval ae = (ThreeDEval)operand;
                for (int sIx = ae.FirstSheetIndex; sIx <= ae.LastSheetIndex; sIx++)
                {
                    int width = ae.Width;
                    int height = ae.Height;
                    for (int rrIx = 0; rrIx < height; rrIx++)
                    {
                        for (int rcIx = 0; rcIx < width; rcIx++)
                        {
                            ValueEval ve = ae.GetValue(sIx, rrIx, rcIx);
                            if (!IsSubtotalCounted && ae.IsSubTotal(rrIx, rcIx)) continue;
                            CollectValue(ve, true, temp);
                        }
                    }
                }
                return;
            }
            if (operand is TwoDEval)
            {
                TwoDEval ae = (TwoDEval)operand;
                int width = ae.Width;
                int height = ae.Height;
                for (int rrIx = 0; rrIx < height; rrIx++)
                {
                    for (int rcIx = 0; rcIx < width; rcIx++)
                    {
                        ValueEval ve = ae.GetValue(rrIx, rcIx);
                        if (!IsSubtotalCounted && ae.IsSubTotal(rrIx, rcIx)) continue;
                        CollectValue(ve, true, temp);
                    }
                }
                return;
            }
            if (operand is RefEval)
            {
                RefEval re = (RefEval)operand;
                for (int sIx = re.FirstSheetIndex; sIx <= re.LastSheetIndex; sIx++)
                {
                    CollectValue(re.GetInnerValueEval(sIx), true, temp);
                }
                return;
            }
            CollectValue((ValueEval)operand, false, temp);
        }
        private void CollectValue(ValueEval ve, bool isViaReference, DoubleList temp)
        {
            if (ve == null)
            {
                throw new ArgumentException("ve must not be null");
            }
            if (ve is BoolEval)
            {
                if (!isViaReference || _isReferenceBoolCounted)
                {
                    BoolEval boolEval = (BoolEval)ve;
                    temp.Add(boolEval.NumberValue);
                }
                return;
            }
            if (ve is NumberEval)
            {
                NumberEval ne = (NumberEval)ve;
                temp.Add(ne.NumberValue);
                return;
            }
            if (ve is StringEval)
            {
                if (isViaReference)
                {
                    // ignore all ref strings
                    return;
                }
                String s = ((StringEval)ve).StringValue;
                Double d = OperandResolver.ParseDouble(s);
                if (double.IsNaN(d))
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                temp.Add(d);
                return;
            }
            if (ve is ErrorEval)
            {
                throw new EvaluationException((ErrorEval)ve);
            }
            if (ve == BlankEval.instance)
            {
                if (_isBlankCounted)
                {
                    temp.Add(0.0);
                }
                return;
            }
            if (ve is NumberValueArrayEval nvae)
            {
                temp.Add(nvae.NumberValues);
                return;
            }
            throw new InvalidOperationException("Invalid ValueEval type passed for conversion: ("
                    + ve.GetType() + ")");
        }
        /**
         * Returns a double array that contains values for the numeric cells
         * from among the list of operands. Blanks and Blank equivalent cells
         * are ignored. Error operands or cells containing operands of type
         * that are considered invalid and would result in #VALUE! error in
         * excel cause this function to return <c>null</c>.
         *
         * @return never <c>null</c>
         */
        protected double[] GetNumberArray(ValueEval[] operands)
        {
            if (operands.Length > MaxNumOperands)
            {
                throw EvaluationException.InvalidValue();
            }
            DoubleList retval = new DoubleList();

            for (int i = 0, iSize = operands.Length; i < iSize; i++)
            {
                CollectValues(operands[i], retval);
            }
            return retval.ToArray();
        }
        /**
         * Ensures that a two dimensional array has all sub-arrays present and the same Length
         * @return <c>false</c> if any sub-array Is missing, or Is of different Length
         */
        protected static bool AreSubArraysConsistent(double[][] values)
        {

            if (values == null || values.Length < 1)
            {
                // TODO this doesn't seem right.  Fix or Add comment.
                return true;
            }

            if (values[0] == null)
            {
                return false;
            }
            int outerMax = values.Length;
            int innerMax = values[0].Length;
            for (int i = 1; i < outerMax; i++)
            { // note - 'i=1' start at second sub-array
                double[] subArr = values[i];
                if (subArr == null)
                {
                    return false;
                }
                if (innerMax != subArr.Length)
                {
                    return false;
                }
            }
            return true;
        }

    }
}