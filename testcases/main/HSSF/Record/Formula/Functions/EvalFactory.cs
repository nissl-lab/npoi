/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


namespace TestCases.HSSF.Record.Formula.Functions
{
    using System;
    using System.Collections;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Record.Formula.Eval;

    /**
     * Test helper class for creating mock <c>Eval</c> objects
     * 
     * @author Josh Micich
     */
    class EvalFactory
    {
        /**
         * Creates a dummy AreaEval 
         * @param values empty (<c>null</c>) entries in this array will be converted to NumberEval.ZERO
         */
        public static AreaEval CreateAreaEval(String areaRefStr, ValueEval[] values)
        {
            AreaPtg areaPtg = new AreaPtg(areaRefStr);
            return CreateAreaEval(areaPtg, values);
        }

        /**
         * Creates a dummy AreaEval 
         * @param values empty (<c>null</c>) entries in this array will be converted to NumberEval.ZERO
         */
        public static AreaEval CreateAreaEval(AreaPtg areaPtg, ValueEval[] values)
        {
            int nCols = areaPtg.LastColumn - areaPtg.FirstColumn + 1;
            int nRows = areaPtg.LastRow - areaPtg.FirstRow + 1;
            int nExpected = nRows * nCols;
            if (values.Length != nExpected)
            {
                throw new InvalidOperationException("Expected " + nExpected + " values but got " + values.Length);
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
            public MockAreaEval(AreaPtg areaPtg, ValueEval[] values)
                : base(areaPtg)
            {

                _values = values;
            }
            public override ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
            {
                if (relativeRowIndex < 0 || relativeRowIndex >= this.Height)
                {
                    throw new ArgumentException("row index out of range");
                }
                int width = this.Width;
                if (relativeColumnIndex < 0 || relativeColumnIndex >= width)
                {
                    throw new ArgumentException("column index out of range");
                }
                int oneDimensionalIndex = relativeRowIndex * width + relativeColumnIndex;
                return _values[oneDimensionalIndex];
            }
            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
            {
                throw new InvalidOperationException("Operation not implemented on this mock object");
            }
        }

        private class MockRefEval : RefEvalBase
        {
            private ValueEval _value;
            public MockRefEval(RefPtg ptg, ValueEval value)
                : base(ptg.Row, ptg.Column)
            {

                _value = value;
            }
            public MockRefEval(Ref3DPtg ptg, ValueEval value)
                : base(ptg.Row, ptg.Column)
            {
                _value = value;
            }
            public override ValueEval InnerValueEval
            {
                get
                {
                    return _value;
                }
            }
            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
            {
                throw new InvalidOperationException("Operation not implemented on this mock object");
            }
        }
    }
}
