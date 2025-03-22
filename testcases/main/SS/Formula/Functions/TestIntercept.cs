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

namespace TestCases.SS.Formula.Functions
{

    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System;
    using TestCases.HSSF;
    using NPOI.SS.Formula.Eval;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Formula.Functions;

    /**
     * Test for Excel function INTERCEPT()
     *
     * @author Johan Karlsteen
     */
    [TestFixture]
    public class TestIntercept
    {
        private static Function INTERCEPT = new Intercept();

        private static ValueEval invoke(Function function, ValueEval xArray, ValueEval yArray)
        {
            ValueEval[] args = new ValueEval[] { xArray, yArray, };
            return function.Evaluate(args, -1, (short)-1);
        }

        private void Confirm(Function function, ValueEval xArray, ValueEval yArray, double expected)
        {
            ValueEval result = invoke(function, xArray, yArray);
            ClassicAssert.AreEqual(typeof(NumberEval), result.GetType());
            ClassicAssert.AreEqual(expected, ((NumberEval)result).NumberValue, 0);
        }
        private void ConfirmError(Function function, ValueEval xArray, ValueEval yArray, ErrorEval expectedError)
        {
            ValueEval result = invoke(function, xArray, yArray);
            ClassicAssert.AreEqual(typeof(ErrorEval), result.GetType());
            ClassicAssert.AreEqual(expectedError.ErrorCode, ((ErrorEval)result).ErrorCode);
        }

        private void ConfirmError(ValueEval xArray, ValueEval yArray, ErrorEval expectedError)
        {
            ConfirmError(INTERCEPT, xArray, yArray, expectedError);
        }
        [Test]
        public void TestBasic()
        {
            Double exp = Math.Pow(10, 7.5);
            ValueEval[] yValues = {
            new NumberEval(3+exp),
            new NumberEval(4+exp),
            new NumberEval(2+exp),
            new NumberEval(5+exp),
            new NumberEval(4+exp),
            new NumberEval(7+exp),
            };
            ValueEval areaEvalY = CreateAreaEval(yValues);

            ValueEval[] xValues = {
            new NumberEval(1),
            new NumberEval(2),
            new NumberEval(3),
            new NumberEval(4),
            new NumberEval(5),
            new NumberEval(6),
            };
            ValueEval areaEvalX = CreateAreaEval(xValues);
            Confirm(INTERCEPT, areaEvalX, areaEvalY, -24516534.39905822);
            // Excel 2010 gives -24516534.3990583
        }

        /**
         * number of items in array is not limited to 30
         */
        [Test]
        public void TestLargeArrays()
        {
            ValueEval[] yValues = CreateMockNumberArray(100, 3); // [1,2,0,1,2,0,...,0,1]
            yValues[0] = new NumberEval(2.0); // Changes first element to 2
            ValueEval[] xValues = CreateMockNumberArray(100, 101); // [1,2,3,4,...,99,100]

            Confirm(INTERCEPT, CreateAreaEval(xValues), CreateAreaEval(yValues), 51.74384236453202);
            // Excel 2010 gives 51.74384236453200
        }

        private ValueEval[] CreateMockNumberArray(int size, double value)
        {
            ValueEval[] result = new ValueEval[size];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = new NumberEval((i + 1) % value);
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
            ConfirmError(areaEvalX, areaEvalY, ErrorEval.NULL_INTERSECTION);
            ConfirmError(areaEvalY, areaEvalX, ErrorEval.REF_INVALID);
        }

        /**
         *  Example from
         *  http://office.microsoft.com/en-us/excel-help/intercept-function-HP010062512.aspx?CTT=5&origin=HA010277524
         */
        [Test]
        public void TestFromFile()
        {

            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("intercept.xls");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            ISheet example1 = wb.GetSheet("Example 1");
            ICell a8 = example1.GetRow(7).GetCell(0);
            ClassicAssert.AreEqual("INTERCEPT(A2:A6,B2:B6)", a8.CellFormula);
            fe.Evaluate(a8);
            ClassicAssert.AreEqual(0.048387097, a8.NumericCellValue, 0.000000001);

        }
    }
}