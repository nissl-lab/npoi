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

    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NUnit.Framework;
    using System;
    /**
     * Tests for Excel functions SUMX2MY2(), SUMX2PY2(), SUMXMY2()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestXYNumericFunction
    {
        private static Function SUM_SQUARES = new Sumx2py2();
        private static Function DIFF_SQUARES = new Sumx2my2();
        private static Function SUM_SQUARES_OF_DIFFS = new Sumxmy2();

        private static ValueEval invoke(Function function, ValueEval xArray, ValueEval yArray)
        {
            ValueEval[] args = new ValueEval[] { xArray, yArray, };
            return function.Evaluate(args, -1, (short)-1);
        }

        private void Confirm(Function function, ValueEval xArray, ValueEval yArray, double expected)
        {
            ValueEval result = invoke(function, xArray, yArray);
            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0);
        }
        private void ConfirmError(Function function, ValueEval xArray, ValueEval yArray, ErrorEval expectedError)
        {
            ValueEval result = invoke(function, xArray, yArray);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(expectedError.ErrorCode, ((ErrorEval)result).ErrorCode);
        }

        private void ConfirmError(ValueEval xArray, ValueEval yArray, ErrorEval expectedError)
        {
            ConfirmError(SUM_SQUARES, xArray, yArray, expectedError);
            ConfirmError(DIFF_SQUARES, xArray, yArray, expectedError);
            ConfirmError(SUM_SQUARES_OF_DIFFS, xArray, yArray, expectedError);
        }
        [Test]
        public void TestBasic()
        {
            ValueEval[] xValues = {
			new NumberEval(1),
			new NumberEval(2),
		};
            ValueEval areaEvalX = CreateAreaEval(xValues);
            Confirm(SUM_SQUARES, areaEvalX, areaEvalX, 10.0);
            Confirm(DIFF_SQUARES, areaEvalX, areaEvalX, 0.0);
            Confirm(SUM_SQUARES_OF_DIFFS, areaEvalX, areaEvalX, 0.0);

            ValueEval[] yValues = {
			new NumberEval(3),
			new NumberEval(4),
		};
            ValueEval areaEvalY = CreateAreaEval(yValues);
            Confirm(SUM_SQUARES, areaEvalX, areaEvalY, 30.0);
            Confirm(DIFF_SQUARES, areaEvalX, areaEvalY, -20.0);
            Confirm(SUM_SQUARES_OF_DIFFS, areaEvalX, areaEvalY, 8.0);
        }

        /**
         * number of items in array is not limited to 30
         */
        [Test]
        public void TestLargeArrays()
        {
            ValueEval[] xValues = CreateMockNumberArray(100, 3);
            ValueEval[] yValues = CreateMockNumberArray(100, 2);

            Confirm(SUM_SQUARES, CreateAreaEval(xValues), CreateAreaEval(yValues), 1300.0);
            Confirm(DIFF_SQUARES, CreateAreaEval(xValues), CreateAreaEval(yValues), 500.0);
            Confirm(SUM_SQUARES_OF_DIFFS, CreateAreaEval(xValues), CreateAreaEval(yValues), 100.0);
        }


        private ValueEval[] CreateMockNumberArray(int size, double value)
        {
            ValueEval[] result = new ValueEval[size];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = new NumberEval(value);
            }
            return result;
        }

        private static ValueEval CreateAreaEval(ValueEval[] values)
        {
            String refStr = "A1:A" + values.Length;
            return EvalFactory.CreateAreaEval(refStr, values);
        }
        [Test]
        public void TestErrors()
        {
            ValueEval[] xValues = {
				ErrorEval.REF_INVALID,
				new NumberEval(2),
		};
            ValueEval areaEvalX = CreateAreaEval(xValues);
            ValueEval[] yValues = {
				new NumberEval(2),
				ErrorEval.NULL_INTERSECTION,
		};
            ValueEval areaEvalY = CreateAreaEval(yValues);
            ValueEval[] zValues = { // wrong size
				new NumberEval(2),
		};
            ValueEval areaEvalZ = CreateAreaEval(zValues);

            // if either arg is an error, that error propagates
            ConfirmError(ErrorEval.REF_INVALID, ErrorEval.NAME_INVALID, ErrorEval.REF_INVALID);
            ConfirmError(areaEvalX, ErrorEval.NAME_INVALID, ErrorEval.NAME_INVALID);
            ConfirmError(ErrorEval.NAME_INVALID, areaEvalX, ErrorEval.NAME_INVALID);

            // array sizes must match
            ConfirmError(areaEvalX, areaEvalZ, ErrorEval.NA);
            ConfirmError(areaEvalZ, areaEvalY, ErrorEval.NA);

            // any error in an array item propagates up
            ConfirmError(areaEvalX, areaEvalX, ErrorEval.REF_INVALID);

            // search for errors array by array, not pair by pair
            ConfirmError(areaEvalX, areaEvalY, ErrorEval.REF_INVALID);
            ConfirmError(areaEvalY, areaEvalX, ErrorEval.NULL_INTERSECTION);
        }
    }

}