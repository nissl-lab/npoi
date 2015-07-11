/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace TestCases.SS.Formula.Functions
{

    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula;
    using System;

    /**
     * Test helper class for creating mock <code>Eval</code> objects
     *
     * @author Josh Micich
     */
    public class EvalFactory
    {

        private EvalFactory()
        {
            // no instances of this class
        }

        /**
         * Creates a dummy AreaEval
         * @param values empty (<code>null</code>) entries in this array will be Converted to NumberEval.ZERO
         */
        public static AreaEval CreateAreaEval(String areaRefStr, ValueEval[] values)
        {
            AreaPtg areaPtg = new AreaPtg(areaRefStr);
            return CreateAreaEval(areaPtg, values);
        }

        /**
         * Creates a dummy AreaEval
         * @param values empty (<code>null</code>) entries in this array will be Converted to NumberEval.ZERO
         */
        public static AreaEval CreateAreaEval(AreaPtg areaPtg, ValueEval[] values)
        {
            int nCols = areaPtg.LastColumn - areaPtg.FirstColumn + 1;
            int nRows = areaPtg.LastRow - areaPtg.FirstRow + 1;
            int nExpected = nRows * nCols;
            if (values.Length != nExpected)
            {
                throw new SystemException("Expected " + nExpected + " values but got " + values.Length);
            }
            for (int i = 0; i < nExpected; i++)
            {
                if (values[i] == null)
                {
                    values[i] = NumberEval.ZERO;
                }
            }
            return new MockAreaEval(areaPtg, values);
        }

        /**
         * Creates a single RefEval (with value zero)
         */
        public static RefEval CreateRefEval(String refStr)
        {
            return CreateRefEval(refStr, NumberEval.ZERO);
        }
        public static RefEval CreateRefEval(String refStr, ValueEval value)
        {
            return new MockRefEval(new RefPtg(refStr), value);
        }

        private class MockAreaEval : AreaEvalBase
        {
            private ValueEval[] _values;
            public MockAreaEval(AreaI areaPtg, ValueEval[] values)
                : base(areaPtg)
            {
                _values = values;
            }
            private MockAreaEval(int firstRow, int firstColumn, int lastRow, int lastColumn, ValueEval[] values)
                : base(firstRow, firstColumn, lastRow, lastColumn)
            {
                _values = values;
            }
            public override ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
            {
                return GetRelativeValue(-1, relativeRowIndex, relativeColumnIndex);
            }
            public override ValueEval GetRelativeValue(int sheetIndex, int relativeRowIndex, int relativeColumnIndex)
            {
                if (relativeRowIndex < 0 || relativeRowIndex >= Height)
                {
                    throw new ArgumentException("row index out of range");
                }
                int width = Width;
                if (relativeColumnIndex < 0 || relativeColumnIndex >= width)
                {
                    throw new ArgumentException("column index out of range");
                }
                int oneDimensionalIndex = relativeRowIndex * width + relativeColumnIndex;
                return _values[oneDimensionalIndex];
            }
            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
            {
                if (relFirstRowIx < 0 || relFirstColIx < 0
                        || relLastRowIx >= Height || relLastColIx >= Width)
                {
                    throw new SystemException("Operation not implemented on this mock object");
                }

                if (relFirstRowIx == 0 && relFirstColIx == 0
                        && relLastRowIx == Height - 1 && relLastColIx == Width - 1)
                {
                    return this;
                }
                ValueEval[] values = transpose(_values, Width, relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);
                return new MockAreaEval(FirstRow + relFirstRowIx, FirstColumn + relFirstColIx,
                     FirstRow + relLastRowIx, FirstColumn + relLastColIx, values);
            }
            private static ValueEval[] transpose(ValueEval[] srcValues, int srcWidth,
                    int relFirstRowIx, int relLastRowIx,
                    int relFirstColIx, int relLastColIx)
            {
                int height = relLastRowIx - relFirstRowIx + 1;
                int width = relLastColIx - relFirstColIx + 1;
                ValueEval[] result = new ValueEval[height * width];
                for (int r = 0; r < height; r++)
                {
                    int srcRowIx = r + relFirstRowIx;
                    for (int c = 0; c < width; c++)
                    {
                        int srcColIx = c + relFirstColIx;
                        int destIx = r * width + c;
                        int srcIx = srcRowIx * srcWidth + srcColIx;
                        result[destIx] = srcValues[srcIx];
                    }
                }
                return result;
            }
            public override TwoDEval GetRow(int rowIndex)
            {
                if (rowIndex >= Height)
                {
                    throw new ArgumentException("Invalid rowIndex " + rowIndex
                            + ".  Allowable range is (0.." + Height + ").");
                }
                ValueEval[] values = new ValueEval[Width];
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = GetRelativeValue(rowIndex, i);
                }
                return new MockAreaEval(rowIndex, FirstColumn, rowIndex, LastColumn, values);
            }
            public override TwoDEval GetColumn(int columnIndex)
            {
                if (columnIndex >= Width)
                {
                    throw new ArgumentException("Invalid columnIndex " + columnIndex
                            + ".  Allowable range is (0.." + Width + ").");
                }
                ValueEval[] values = new ValueEval[Height];
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = GetRelativeValue(i, columnIndex);
                }
                return new MockAreaEval(FirstRow, columnIndex, LastRow, columnIndex, values);
            }
        }

        private class MockRefEval : RefEvalBase
        {
            private ValueEval _value;
            public MockRefEval(RefPtg ptg, ValueEval value)
                : base(-1, -1, ptg.Row, ptg.Column)
            {
                _value = value;
            }
            public MockRefEval(Ref3DPtg ptg, ValueEval value)
                : base(-1, -1, ptg.Row, ptg.Column)
            {
                _value = value;
            }
            public override ValueEval GetInnerValueEval(int sheetIndex)
            {
                return _value;
            }
            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
            {
                throw new SystemException("Operation not implemented on this mock object");
            }
        }
    }

}