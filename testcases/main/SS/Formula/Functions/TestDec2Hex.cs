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

using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using NUnit.Framework;
using NPOI.SS.Formula;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
namespace TestCases.SS.Formula.Functions
{

    /**
     * Tests for {@link Dec2Hex}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestDec2Hex
    {
        public TestDec2Hex()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        private static ValueEval invokeValue(String number1, String number2)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), new StringEval(number2), };
            return new Dec2Hex().Evaluate(args, -1, -1);
        }
        private static ValueEval invokeBack(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1) };
            return new Hex2Dec().Evaluate(args, -1, -1);
        }
        private static ValueEval invokeValue(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), };
            return new Dec2Hex().Evaluate(args, -1, -1);
        }

        private static void ConfirmValue(String msg, String number1, String number2, String expected)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual(expected, ((StringEval)result).StringValue, msg);
        }

        private static void ConfirmValue(String msg, String number1, String expected)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual(expected, ((StringEval)result).StringValue, msg);
        }

        private static void ConfirmValueError(String msg, String number1, String number2, ErrorEval numError)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(numError, result, msg);
        }
        [Test]
        public void TestBasic()
        {
            ConfirmValue("Converts decimal 100 to hexadecimal with 0 characters (64)", "100", "0", "64");
            ConfirmValue("Converts decimal 100 to hexadecimal with 4 characters (0064)", "100", "4", "0064");
            ConfirmValue("Converts decimal 100 to hexadecimal with 5 characters (0064)", "100", "5", "00064");
            ConfirmValue("Converts decimal 100 to hexadecimal with 10 (default) characters", "100", "10", "0000000064");
            ConfirmValue("If argument places Contains a decimal value, dec2hex ignores the numbers to the right side of the decimal point.", "100", "10.0", "0000000064");

            ConfirmValue("Converts decimal -54 to hexadecimal, 2 is ignored", "-54", "2", "FFFFFFFFCA");
            ConfirmValue("places is optionnal", "-54", "FFFFFFFFCA");


            ConfirmValue("Converts normal decimal number to hexadecimal", "100", "64");

            String maxInt = Int32.MaxValue.ToString();
            Assert.AreEqual("2147483647", maxInt);
            ConfirmValue("Converts INT_MAX to hexadecimal", maxInt, "7FFFFFFF");

            String minInt = Int32.MinValue.ToString();
            Assert.AreEqual("-2147483648", minInt);
            ConfirmValue("Converts INT_MIN to hexadecimal", minInt, "FF80000000");

            String maxIntPlusOne = (((long)Int32.MaxValue) + 1).ToString();
            Assert.AreEqual("2147483648", maxIntPlusOne);
            ConfirmValue("Converts INT_MAX + 1 to hexadecimal", maxIntPlusOne, "80000000");

            String maxLong = (549755813887).ToString();
            Assert.AreEqual("549755813887", maxLong);
            ConfirmValue("Converts the max supported value to hexadecimal", maxLong, "7FFFFFFFFF");

            String minLong = (-549755813888L).ToString();
            Assert.AreEqual("-549755813888", minLong);
            ConfirmValue("Converts the min supported value to hexadecimal", minLong, "FF80000000");

        }
        [Test]
        public void TestErrors()
        {
            ConfirmValueError("Out of range min number", "-549755813889", "0", ErrorEval.NUM_ERROR);
            ConfirmValueError("Out of range max number", "549755813888", "0", ErrorEval.NUM_ERROR);

            ConfirmValueError("negative places not allowed", "549755813888", "-10", ErrorEval.NUM_ERROR);
            ConfirmValueError("non number places not allowed", "ABCDEF", "0", ErrorEval.VALUE_INVALID);
        }

        [Test]
        public void TestEvalOperationEvaluationContextFails()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ErrorEval.VALUE_INVALID };
            ValueEval result = new Dec2Hex().Evaluate(args, ctx);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        private OperationEvaluationContext CreateContext()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("123.43");
            cell = row.CreateCell(1);
            cell.SetCellValue("8");
            cell = row.CreateCell(2);
            cell.SetCellValue("-8");

            HSSFEvaluationWorkbook workbook = HSSFEvaluationWorkbook.Create(wb);
            WorkbookEvaluator workbookEvaluator = new WorkbookEvaluator(workbook, new IStabilityClassifier1(), null);
            OperationEvaluationContext ctx = new OperationEvaluationContext(workbookEvaluator,
                    workbook, 0, 0, 0, null);
            return ctx;
        }
        class IStabilityClassifier1 : IStabilityClassifier
        {

            public override bool IsCellFinal(int sheetIndex, int rowIndex, int columnIndex)
            {
                return true;
            }
        }
        [Test]
        public void TestRefs()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0) };
            ValueEval result = new Dec2Hex().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(StringEval), result.GetType(), "Had: " + result.ToString());
            Assert.AreEqual("7B", ((StringEval)result).StringValue);
        }

        [Test]
        public void TestWithPlacesIntInt()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 1) };
            ValueEval result = new Dec2Hex().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(StringEval), result.GetType(), "Had: " + result.ToString());
            Assert.AreEqual("0000007B", ((StringEval)result).StringValue);
        }

        [Test]
        public void TestWithPlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 1) };
            ValueEval result = new Dec2Hex().Evaluate(args, ctx);

            Assert.AreEqual(typeof(StringEval), result.GetType(), "Had: " + result.ToString());
            Assert.AreEqual("0000007B", ((StringEval)result).StringValue);
        }

        [Test]
        public void TestWithTooManyParamsIntInt()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 1), ctx.GetRefEval(0, 1) };
            ValueEval result = new Dec2Hex().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        [Test]
        public void TestWithTooManyParams()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 1), ctx.GetRefEval(0, 1) };
            ValueEval result = new Dec2Hex().Evaluate(args, ctx);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        [Test]
        public void TestWithErrorPlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ErrorEval.NULL_INTERSECTION };
            ValueEval result = new Dec2Hex().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.NULL_INTERSECTION, result);
        }

        [Test]
        public void TestWithNegativePlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 2) };
            ValueEval result = new Dec2Hex().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.NUM_ERROR, result);
        }

        [Test]
        public void TestWithEmptyPlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(1, 0) };
            ValueEval result = new Dec2Hex().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        [Test]
        public void TestBackAndForth()
        {
            for (int i = -512; i < 512; i++)
            {
                ValueEval result = invokeValue(i.ToString());
                Assert.AreEqual(typeof(StringEval), result.GetType(), "Had: " + result.ToString());

                ValueEval back = invokeBack(((StringEval)result).StringValue);
                Assert.AreEqual(typeof(NumberEval), back.GetType(), "Had: " + back.ToString());

                Assert.AreEqual(i.ToString(),
                        ((NumberEval)back).StringValue);
            }
        }

    }
}