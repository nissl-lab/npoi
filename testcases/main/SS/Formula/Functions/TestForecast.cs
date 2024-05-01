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
    using NUnit.Framework;
    using System;
    using HSSF;
    using NPOI.SS.Formula.Eval;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Formula.Functions;

    /**
     * Test for Excel function FORECAST()
     *
     * @author Ken Smith
     */
    [TestFixture]
    public class TestForecast
    {
        private static readonly Function FORECAST = new Forecast();

        /// <summary>
        /// This test is replicated in the "TestBasic" tab of the "Forecast.xls" file.
        /// </summary>
        [Test]
        public void TestBasic()
        {
            ValueEval x = new NumberEval(100);
            ValueEval[] yValues = [
                new NumberEval(1),
                new NumberEval(2),
                new NumberEval(3),
                new NumberEval(4),
                new NumberEval(5),
                new NumberEval(6)
            ];

            ValueEval[] xValues = [
                new NumberEval(2),
                new NumberEval(4),
                new NumberEval(6),
                new NumberEval(8),
                new NumberEval(10),
                new NumberEval(12)
            ];
            Confirm(x, CreateAreaEval(yValues), CreateAreaEval(xValues), 50.0);
            // Excel 365 build 2402 gives 50.0
        }

        /// <summary>
        /// This test is replicated in the "TestLargeNumbers" tab of the "Forecast.xls" file.
        /// </summary>
        [Test]
        public void TestLargeNumbers()
        {
            double exp = Math.Pow(10, 7.5);
            ValueEval x = new NumberEval(100);
            ValueEval[] yValues = [
                new NumberEval(3 + exp),
                new NumberEval(4 + exp),
                new NumberEval(2 + exp),
                new NumberEval(5 + exp),
                new NumberEval(4 + exp),
                new NumberEval(7 + exp)
            ];

            ValueEval[] xValues = [
                new NumberEval(1),
                new NumberEval(2),
                new NumberEval(3),
                new NumberEval(4),
                new NumberEval(5),
                new NumberEval(6)
            ];
            Confirm(x, CreateAreaEval(yValues), CreateAreaEval(xValues), 31622844.1826363);
            // Excel 365 build 2402 gives 31622844.1826363
        }

        /// <summary>
        /// This test is replicated in the "TestLargeArrays" tab of the "Forecast.xls" file.
        /// </summary>
        [Test]
        public void TestLargeArrays()
        {
            ValueEval x = new NumberEval(100);
            ValueEval[] yValues = CreateMockNumberArray(100, 3); // [2,2,0,1,2,0,...,0,1]
            yValues[0] = new NumberEval(2.0); // Changes first element to 2
            ValueEval[] xValues = CreateMockNumberArray(100, 101); // [1,2,3,4,...,99,100]

            Confirm(x, CreateAreaEval(yValues), CreateAreaEval(xValues), 0.960990099);
            // Excel 365 build 2402 gives 0.960990099
        }

        [Test]
        public void TestErrors()
        {
            NumberEval x = new(100);

            ValueEval areaEval1 = CreateAreaEval([new NumberEval(2)]);
            ValueEval areaEval2 = CreateAreaEval([new NumberEval(2), new NumberEval(2)]); // different size

            ValueEval areaEvalWithNullError = CreateAreaEval([new NumberEval(2), ErrorEval.NULL_INTERSECTION]);
            ValueEval areaEvalWithRefError = CreateAreaEval([ErrorEval.REF_INVALID, new NumberEval(2)]);

            // if either arg is an error, that error propagates
            ConfirmError(x, ErrorEval.REF_INVALID, ErrorEval.NAME_INVALID, ErrorEval.REF_INVALID);
            ConfirmError(x, areaEvalWithRefError, ErrorEval.NAME_INVALID, ErrorEval.NAME_INVALID);
            ConfirmError(x, ErrorEval.NAME_INVALID, areaEvalWithRefError, ErrorEval.NAME_INVALID);

            // array sizes must match
            ConfirmError(x, areaEval1, areaEval2, ErrorEval.NA);

            // any error in an array item propagates up
            ConfirmError(x, areaEvalWithRefError, areaEvalWithRefError, ErrorEval.REF_INVALID);

            // search for errors array by array, not pair by pair
            ConfirmError(x, areaEvalWithRefError, areaEvalWithNullError, ErrorEval.REF_INVALID);
            ConfirmError(x, areaEvalWithNullError, areaEvalWithRefError, ErrorEval.NULL_INTERSECTION);
        }

        /**
         *  Example from
         *  https://support.microsoft.com/en-us/office/forecast-and-forecast-linear-functions-50ca49c9-7b40-4892-94e4-7ad38bbeda99
         */
        [Test]
        public void TestFromFile()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("Forecast.xls");
            HSSFFormulaEvaluator fe = new(wb);

            ISheet example1 = wb.GetSheet("TestFromFile");
            ICell a8 = example1.GetRow(7).GetCell(0);
            Assert.AreEqual("FORECAST(30,A2:A6,B2:B6)", a8.CellFormula);
            fe.Evaluate(a8);
            Assert.AreEqual(10.60725309, a8.NumericCellValue, 0.00000001);
        }

        private static ValueEval Invoke(ValueEval x, ValueEval yArray, ValueEval xArray)
        {
            ValueEval[] args = [x, yArray, xArray];
            return FORECAST.Evaluate(args, -1, (short) -1);
        }

        private static void Confirm(ValueEval x, ValueEval yArray, ValueEval xArray, double expected)
        {
            ValueEval result = Invoke(x, yArray, xArray);
            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(expected, ((NumberEval) result).NumberValue, expected * .000000001);
        }

        private static void ConfirmError(ValueEval x, ValueEval yArray, ValueEval xArray, ErrorEval expectedError)
        {
            ValueEval result = Invoke(x, yArray, xArray);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(expectedError, (ErrorEval) result);
        }

        private static ValueEval[] CreateMockNumberArray(int size, double value)
        {
            ValueEval[] result = new ValueEval[size];
            for(int i = 0; i < result.Length; i++)
            {
                result[i] = new NumberEval((i + 1) % value);
            }

            return result;
        }

        private static ValueEval CreateAreaEval(ValueEval[] values)
        {
            string refStr = "A1:A" + values.Length;
            return EvalFactory.CreateAreaEval(refStr, values);
        }
    }
}