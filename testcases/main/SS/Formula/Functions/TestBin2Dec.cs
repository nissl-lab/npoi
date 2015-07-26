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
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;

    /**
     * Tests for {@link Bin2Dec}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestBin2Dec
    {

        private static ValueEval invokeValue(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1) };
            return new Bin2Dec().Evaluate(args, -1, -1);
        }

        private static void ConfirmValue(String msg, String number1, String expected)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(NumberEval), result.GetType(), "Had: " + result.ToString());
            Assert.AreEqual(expected, ((NumberEval)result).StringValue, msg);
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
            ConfirmValue("Converts binary '00101' to decimal (5)", "00101", "5");
            ConfirmValue("Converts binary '1111111111' to decimal (-1)", "1111111111", "-1");
            ConfirmValue("Converts binary '1111111110' to decimal (-2)", "1111111110", "-2");
            ConfirmValue("Converts binary '0111111111' to decimal (511)", "0111111111", "511");
        }

        [Test]
        public void TestErrors()
        {
            ConfirmValueError("does not support more than 10 digits", "01010101010", ErrorEval.NUM_ERROR);
            ConfirmValueError("not a valid binary number", "GGGGGGG", ErrorEval.NUM_ERROR);
            ConfirmValueError("not a valid binary number", "3.14159", ErrorEval.NUM_ERROR);
        }

        [Test]
        public void TestEvalOperationEvaluationContext()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0) };
            ValueEval result = new Bin2Dec().Evaluate(args, ctx);

            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual("0", ((NumberEval)result).StringValue);
        }

        [Test]
        public void TestEvalOperationEvaluationContextFails()
        {
            OperationEvaluationContext ctx = CreateContext();

            ValueEval[] args = new ValueEval[] { ctx.GetRefEval(0, 0), ctx.GetRefEval(0, 0) };
            ValueEval result = new Bin2Dec().Evaluate(args, ctx);

            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        private OperationEvaluationContext CreateContext()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet();
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
            ValueEval result = new Bin2Dec().Evaluate(args, -1, -1);

            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual("0", ((NumberEval)result).StringValue);
        }
    }
}
