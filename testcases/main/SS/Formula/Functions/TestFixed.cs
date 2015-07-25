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
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;

    [TestFixture]
    public class TestFixed
    {

        private HSSFCell cell11;
        private HSSFFormulaEvaluator Evaluator;

        [SetUp]
        public void SetUp()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            try
            {
                HSSFSheet sheet = wb.CreateSheet("new sheet") as HSSFSheet;
                cell11 = sheet.CreateRow(0).CreateCell(0) as HSSFCell;
                cell11.SetCellType(CellType.Formula);
                Evaluator = new HSSFFormulaEvaluator(wb);
            }
            finally
            {
                //wb.Close();
            }
        }

        [Test]
        public void TestValid()
        {
            // thousands separator
            Confirm("FIXED(1234.56789,2,TRUE)", "1234.57");
            Confirm("FIXED(1234.56789,2,FALSE)", "1,234.57");
            // rounding
            Confirm("FIXED(1.8,0,TRUE)", "2");
            Confirm("FIXED(1.2,0,TRUE)", "1");
            Confirm("FIXED(1.5,0,TRUE)", "2");
            Confirm("FIXED(1,0,TRUE)", "1");
            // fractional digits
            Confirm("FIXED(1234.56789,7,TRUE)", "1234.5678900");
            Confirm("FIXED(1234.56789,0,TRUE)", "1235");
            Confirm("FIXED(1234.56789,-1,TRUE)", "1230");
            // less than three arguments
            Confirm("FIXED(1234.56789)", "1,234.57");
            Confirm("FIXED(1234.56789,3)", "1,234.568");
            // invalid arguments
            ConfirmValueError("FIXED(\"invalid\")");
            ConfirmValueError("FIXED(1,\"invalid\")");
            ConfirmValueError("FIXED(1,2,\"invalid\")");
            // strange arguments
            Confirm("FIXED(1000,2,8)", "1000.00");
            Confirm("FIXED(1000,2,0)", "1,000.00");
            // corner cases
            Confirm("FIXED(1.23456789012345,15,TRUE)", "1.234567890123450");
            // Seems POI accepts longer numbers than Excel does, excel Trims the
            // number to 15 digits and Removes the "9" in the formula itself.
            // Not the fault of FIXED though.
            // Confirm("FIXED(1.234567890123459,15,TRUE)", "1.234567890123450");
            Confirm("FIXED(60,-2,TRUE)", "100");
            Confirm("FIXED(10,-2,TRUE)", "0");
            // rounding propagation
            Confirm("FIXED(99.9,0,TRUE)", "100");
        }

        [Test]
        public void TestOptionalParams()
        {
            Fixed fixedFunc = new Fixed();
            ValueEval Evaluate = fixedFunc.Evaluate(0, 0, new NumberEval(1234.56789));
            Assert.IsTrue(Evaluate is StringEval);
            Assert.AreEqual("1,234.57", ((StringEval)Evaluate).StringValue);

            Evaluate = fixedFunc.Evaluate(0, 0, new NumberEval(1234.56789), new NumberEval(1));
            Assert.IsTrue(Evaluate is StringEval);
            Assert.AreEqual("1,234.6", ((StringEval)Evaluate).StringValue);

            Evaluate = fixedFunc.Evaluate(0, 0, new NumberEval(1234.56789), new NumberEval(1), BoolEval.TRUE);
            Assert.IsTrue(Evaluate is StringEval);
            Assert.AreEqual("1234.6", ((StringEval)Evaluate).StringValue);

            Evaluate = fixedFunc.Evaluate(new ValueEval[] { }, 1, 1);
            Assert.IsTrue(Evaluate is ErrorEval);

            Evaluate = fixedFunc.Evaluate(new ValueEval[] { new NumberEval(1), new NumberEval(1), new NumberEval(1), new NumberEval(1) }, 1, 1);
            Assert.IsTrue(Evaluate is ErrorEval);
        }

        private void Confirm(String formulaText, String expectedResult)
        {
            cell11.CellFormula = (/*setter*/formulaText);
            Evaluator.ClearAllCachedResultValues();
            CellValue cv = Evaluator.Evaluate(cell11);
            Assert.AreEqual(CellType.String, cv.CellType, "Wrong result type: " + cv.FormatAsString());
            String actualValue = cv.StringValue;
            Assert.AreEqual(expectedResult, actualValue);
        }

        private void ConfirmValueError(String formulaText)
        {
            cell11.CellFormula = (/*setter*/formulaText);
            Evaluator.ClearAllCachedResultValues();
            CellValue cv = Evaluator.Evaluate(cell11);
            Assert.IsTrue(cv.CellType == CellType.Error
                    && cv.ErrorValue == ErrorConstants.ERROR_VALUE, "Wrong result type: " + cv.FormatAsString());
        }
    }
}

