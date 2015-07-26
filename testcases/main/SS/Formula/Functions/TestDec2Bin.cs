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
    using System;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;

    /**
     * Tests for {@link Dec2Bin}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestDec2Bin
    {

        private static ValueEval invokeValue(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1) };
            return new Dec2Bin().Evaluate(args, -1, -1);
        }

        private static ValueEval invokeBack(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1) };
            return new Bin2Dec().Evaluate(args, -1, -1);
        }

        private static void ConfirmValue(String msg, String number1, String expected)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(StringEval), result.GetType(), "Had: " + result.ToString());
            Assert.AreEqual(expected, ((StringEval)result).StringValue, msg);
        }

        private static void ConfirmValueError(String msg, String number1, ErrorEval numError)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(numError, result, msg);
        }

        [Test]
        public void TestBasic()
        {
            ConfirmValue("Converts binary '00101' from binary (5)", "5", "101");
            ConfirmValue("Converts binary '1111111111' from binary (-1)", "-1", "1111111111");
            ConfirmValue("Converts binary '1111111110' from binary (-2)", "-2", "1111111110");
            ConfirmValue("Converts binary '0111111111' from binary (511)", "511", "111111111");
            ConfirmValue("Converts binary '1000000000' from binary (511)", "-512", "1000000000");
        }

        [Test]
        public void TestErrors()
        {
            ConfirmValueError("fails for >= 512 or < -512", "512", ErrorEval.NUM_ERROR);
            ConfirmValueError("fails for >= 512 or < -512", "-513", ErrorEval.NUM_ERROR);
            ConfirmValueError("not a valid decimal number", "GGGGGGG", ErrorEval.VALUE_INVALID);
            ConfirmValueError("not a valid decimal number", "3.14159a", ErrorEval.VALUE_INVALID);
        }

        [Test]
        public void TestEvalOperationEvaluationContext()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0) };
            ValueEval result = new Dec2Bin().Evaluate(args, ctx);

            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual("1101", ((StringEval)result).StringValue);
        }

        [Test]
        public void TestEvalOperationEvaluationContextFails()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ErrorEval.VALUE_INVALID };
            ValueEval result = new Dec2Bin().Evaluate(args, ctx);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        private OperationEvaluationContext CreateContext()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("13.43");
            cell = row.CreateCell(1);
            cell.SetCellValue("8");
            cell = row.CreateCell(2);
            cell.SetCellValue("-8");
            cell = row.CreateCell(3);
            cell.SetCellValue("1");

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
            ValueEval result = new Dec2Bin().Evaluate(args, -1, -1);

            Assert.AreEqual(result.GetType(), typeof(StringEval), "Had: " + result.ToString());
            Assert.AreEqual("1101", ((StringEval)result).StringValue);
        }

        [Test]
        public void TestWithPlacesIntInt()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 1) };
            ValueEval result = new Dec2Bin().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(StringEval), result.GetType(), "Had: " + result.ToString());
            // TODO: documentation and behavior do not match here!
            Assert.AreEqual("1101", ((StringEval)result).StringValue);
        }

        [Test]
        public void TestWithPlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 1) };
            ValueEval result = new Dec2Bin().Evaluate(args, ctx);

            Assert.AreEqual(typeof(StringEval), result.GetType(), "Had: " + result.ToString());
            // TODO: documentation and behavior do not match here!
            Assert.AreEqual("1101", ((StringEval)result).StringValue);
        }

        [Test]
        public void TestWithToshortPlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 3) };
            ValueEval result = new Dec2Bin().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.NUM_ERROR, result);
        }

        [Test]
        public void TestWithTooManyParamsIntInt()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 1), ctx.GetRefEval(0, 1) };
            ValueEval result = new Dec2Bin().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        [Test]
        public void TestWithTooManyParams()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 1), ctx.GetRefEval(0, 1) };
            ValueEval result = new Dec2Bin().Evaluate(args, ctx);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        [Test]
        public void TestWithErrorPlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ErrorEval.NULL_INTERSECTION };
            ValueEval result = new Dec2Bin().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.NULL_INTERSECTION, result);
        }

        [Test]
        public void TestWithNegativePlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 2) };
            ValueEval result = new Dec2Bin().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.NUM_ERROR, result);
        }

        [Test]
        public void TestWithZeroPlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), new NumberEval(0.0) };
            ValueEval result = new Dec2Bin().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.NUM_ERROR, result);
        }

        [Test]
        public void TestWithEmptyPlaces()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(1, 0) };
            ValueEval result = new Dec2Bin().Evaluate(args, -1, -1);

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

                Assert.AreEqual(i.ToString(), ((NumberEval)back).StringValue);
            }
        }
    }
}
