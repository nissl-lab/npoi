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

namespace TestCases.SS.Formula
{
    using System;
    using System.IO;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;

    using TestCases.HSSF;

    /**
     * Tests {@link WorkbookEvaluator}.
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestWorkbookEvaluator
    {
        private static readonly double EPSILON = 0.0000001;

        private static ValueEval EvaluateFormula(Ptg[] ptgs)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet().CreateRow(0).CreateCell(0);
            IEvaluationWorkbook ewb = HSSFEvaluationWorkbook.Create(wb);
            OperationEvaluationContext ec = new OperationEvaluationContext(null, ewb, 0, 0, 0, null);
            return new WorkbookEvaluator(null, null, null).EvaluateFormula(ec, ptgs);
        }

        /**
         * Make sure that the Evaluator can directly handle tAttrSum (instead of relying on re-parsing
         * the whole formula which Converts tAttrSum to tFuncVar("SUM") )
         */
        [Test]
        public void TestAttrSum()
        {
            Ptg[] ptgs = { new IntPtg(42), AttrPtg.SUM, };

            ValueEval result = EvaluateFormula(ptgs);
            Assert.AreEqual(42, ((NumberEval)result).NumberValue, 0.0);
        }

        /**
         * Make sure that the Evaluator can directly handle (deleted) ref error tokens
         * (instead of relying on re-parsing the whole formula which Converts these
         * to the error constant #REF! )
         */
        [Test]
        public void TestRefErr()
        {

            ConfirmRefErr(new RefErrorPtg());
            ConfirmRefErr(new AreaErrPtg());
            ConfirmRefErr(new DeletedRef3DPtg(0));
            ConfirmRefErr(new DeletedArea3DPtg(0));
        }
        private static void ConfirmRefErr(Ptg ptg)
        {
            Ptg[] ptgs = { ptg };

            ValueEval result = EvaluateFormula(ptgs);
            Assert.AreEqual(ErrorEval.REF_INVALID, result);
        }

        /**
         * Make sure that the Evaluator can directly handle tAttrSum (instead of relying on re-parsing
         * the whole formula which Converts tAttrSum to tFuncVar("SUM") )
         */
        [Test]
        public void TestMemFunc()
        {
            Ptg[] ptgs = { new IntPtg(42), AttrPtg.SUM, };

            ValueEval result = EvaluateFormula(ptgs);
            Assert.AreEqual(42, ((NumberEval)result).NumberValue, 0.0);
        }

        [Test]
        public void TestEvaluateMultipleWorkbooks()
        {
            HSSFWorkbook wbA = HSSFTestDataSamples.OpenSampleWorkbook("multibookFormulaA.xls");
            HSSFWorkbook wbB = HSSFTestDataSamples.OpenSampleWorkbook("multibookFormulaB.xls");

            HSSFFormulaEvaluator EvaluatorA = new HSSFFormulaEvaluator(wbA);
            HSSFFormulaEvaluator EvaluatorB = new HSSFFormulaEvaluator(wbB);

            // Hook up the IWorkbook Evaluators to enable Evaluation of formulas across books
            String[] bookNames = { "multibookFormulaA.xls", "multibookFormulaB.xls", };
            HSSFFormulaEvaluator[] Evaluators = { EvaluatorA, EvaluatorB, };
            HSSFFormulaEvaluator.SetupEnvironment(bookNames, Evaluators);

            ISheet aSheet1 = wbA.GetSheetAt(0);
            ISheet bSheet1 = wbB.GetSheetAt(0);

            // Simple case - single link from wbA to wbB
            ConfirmFormula(wbA, 0, 0, 0, "[multibookFormulaB.xls]BSheet1!B1");
            ICell cell = aSheet1.GetRow(0).GetCell(0);
            ConfirmEvaluation(35, EvaluatorA, cell);

            // more complex case - back link into wbA
            // [wbA]ASheet1!A2 references (among other things) [wbB]BSheet1!B2
            ConfirmFormula(wbA, 0, 1, 0, "[multibookFormulaB.xls]BSheet1!$B$2+2*A3");

            // [wbB]BSheet1!B2 references (among other things) [wbA]AnotherSheet!A1:B2
            ConfirmFormula(wbB, 0, 1, 1, "SUM([multibookFormulaA.xls]AnotherSheet!$A$1:$B$2)+B3");

            cell = aSheet1.GetRow(1).GetCell(0);
            ConfirmEvaluation(264, EvaluatorA, cell);

            // change [wbB]BSheet1!B3 (from 50 to 60)
            ICell cellB3 = bSheet1.GetRow(2).GetCell(1);
            cellB3.SetCellValue(60);
            EvaluatorB.NotifyUpdateCell(cellB3);
            ConfirmEvaluation(274, EvaluatorA, cell);

            // change [wbA]ASheet1!A3 (from 100 to 80)
            ICell cellA3 = aSheet1.GetRow(2).GetCell(0);
            cellA3.SetCellValue(80);
            EvaluatorA.NotifyUpdateCell(cellA3);
            ConfirmEvaluation(234, EvaluatorA, cell);

            // change [wbA]AnotherSheet!A1 (from 2 to 3)
            ICell cellA1 = wbA.GetSheetAt(1).GetRow(0).GetCell(0);
            cellA1.SetCellValue(3);
            EvaluatorA.NotifyUpdateCell(cellA1);
            ConfirmEvaluation(235, EvaluatorA, cell);
        }

        private static void ConfirmEvaluation(double expectedValue, HSSFFormulaEvaluator fe, ICell cell)
        {
            Assert.AreEqual(expectedValue, fe.Evaluate(cell).NumberValue, 0.0);
        }

        private static void ConfirmFormula(HSSFWorkbook wb, int sheetIndex, int rowIndex, int columnIndex,
                String expectedFormula)
        {
            ICell cell = wb.GetSheetAt(sheetIndex).GetRow(rowIndex).GetCell(columnIndex);
            Assert.AreEqual(expectedFormula, cell.CellFormula);
        }

        /**
         * This Test Makes sure that any {@link MissingArgEval} that propagates to
         * the result of a function Gets translated to {@link BlankEval}.
         */
        [Test]
        public void TestMissingArg()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.CellFormula = "1+IF(1,,)";
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = null;
            try
            {
                cv = fe.Evaluate(cell);
            }
            catch (Exception)
            {
                Assert.Fail("Missing arg result not being handled correctly.");
            }

            Assert.AreEqual(CellType.Numeric, cv.CellType);

            // Adding blank to 1.0 gives 1.0
            Assert.AreEqual(1.0, cv.NumberValue, 0.0);

            // check with string operand
            cell.CellFormula = "\"abc\"&IF(1,,)";
            fe.NotifySetFormula(cell);
            cv = fe.Evaluate(cell);
            Assert.AreEqual(CellType.String, cv.CellType);

            // Adding blank to "abc" gives "abc"
            Assert.AreEqual("abc", cv.StringValue);

            // check CHOOSE()
            cell.CellFormula = "\"abc\"&CHOOSE(2,5,,9)";
            fe.NotifySetFormula(cell);
            cv = fe.Evaluate(cell);
            Assert.AreEqual(CellType.String, cv.CellType);

            // Adding blank to "abc" gives "abc"
            Assert.AreEqual("abc", cv.StringValue);
        }

        /**
         * Functions like IF, INDIRECT, INDEX, OFFSET etc can return AreaEvals which
         * should be dereferenced by the Evaluator
         */
        [Test]
        public void TestResultOutsideRange()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet("Sheet1").CreateRow(0).CreateCell(0);
            cell.CellFormula = "D2:D5"; // IF(TRUE,D2:D5,D2) or  OFFSET(D2:D5,0,0) would work too
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            CellValue cv;
            try
            {
                cv = fe.Evaluate(cell);
            }
            catch (ArgumentException e)
            {
                if ("Specified row index (0) is outside the allowed range (1..4)".Equals(e.Message))
                {
                    Assert.Fail("Identified bug in result dereferencing");
                }
                throw;
            }

            Assert.AreEqual(CellType.Error, cv.CellType);
            Assert.AreEqual(ErrorEval.VALUE_INVALID.ErrorCode, cv.ErrorValue);

            // verify circular refs are still detected properly
            fe.ClearAllCachedResultValues();
            cell.CellFormula = "OFFSET(A1,0,0)";
            cv = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Error, cv.CellType);
            Assert.AreEqual(ErrorEval.CIRCULAR_REF_ERROR.ErrorCode, cv.ErrorValue);
        }

        /**
         * formulas with defined names.
         */
        [Test]
        public void TestNamesInFormulas()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");

            IName name1 = wb.CreateName();
            name1.NameName = "aConstant";
            name1.RefersToFormula = "3.14";

            IName name2 = wb.CreateName();
            name2.NameName = "aFormula";
            name2.RefersToFormula = "SUM(Sheet1!$A$1:$A$3)";

            IName name3 = wb.CreateName();
            name3.NameName = "aSet";
            name3.RefersToFormula = "Sheet1!$A$2:$A$4";

            IRow row0 = sheet.CreateRow(0);
            IRow row1 = sheet.CreateRow(1);
            IRow row2 = sheet.CreateRow(2);
            IRow row3 = sheet.CreateRow(3);
            row0.CreateCell(0).SetCellValue(2);
            row1.CreateCell(0).SetCellValue(5);
            row2.CreateCell(0).SetCellValue(3);
            row3.CreateCell(0).SetCellValue(7);

            row0.CreateCell(2).SetCellFormula("aConstant");
            row1.CreateCell(2).SetCellFormula("aFormula");
            row2.CreateCell(2).SetCellFormula("SUM(aSet)");
            row3.CreateCell(2).SetCellFormula("aConstant+aFormula+SUM(aSet)");

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual(3.14, fe.Evaluate(row0.GetCell(2)).NumberValue);
            Assert.AreEqual(10.0, fe.Evaluate(row1.GetCell(2)).NumberValue);
            Assert.AreEqual(15.0, fe.Evaluate(row2.GetCell(2)).NumberValue);
            Assert.AreEqual(28.14, fe.Evaluate(row3.GetCell(2)).NumberValue);
        }


        // Test IF-Equals Formula Evaluation (bug 58591)
        
        private IWorkbook TestIFEqualsFormulaEvaluation_setup(String formula, CellType a1CellType)
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("IFEquals");
            IRow row = sheet.CreateRow(0);
            ICell A1 = row.CreateCell(0);
            ICell B1 = row.CreateCell(1);
            ICell C1 = row.CreateCell(2);
            ICell D1 = row.CreateCell(3);

            switch (a1CellType)
            {
                case CellType.Numeric:
                    A1.SetCellValue(1.0);
                    // "A1=1" should return true
                    break;
                case CellType.String:
                    A1.SetCellValue("1");
                    // "A1=1" should return false
                    // "A1=\"1\"" should return true
                    break;
                case CellType.Boolean:
                    A1.SetCellValue(true);
                    // "A1=1" should return true
                    break;
                case CellType.Formula:
                    A1.SetCellFormula("1");
                    // "A1=1" should return true
                    break;
                case CellType.Blank:
                    A1.SetCellValue((String)null);
                    // "A1=1" should return false
                    break;
            }
            B1.SetCellValue(2.0);
            C1.SetCellValue(3.0);
            D1.CellFormula = (formula);

            return wb;
        }

        private void TestIFEqualsFormulaEvaluation_teardown(IWorkbook wb)
        {
            try
            {
                wb.Close();
            }
            catch (IOException)
            {
                Assert.Fail("Unable to close workbook");
            }
        }


        private void TestIFEqualsFormulaEvaluation_evaluate(
            String formula, CellType cellType, String expectedFormula, double expectedResult)
        {
            IWorkbook wb = TestIFEqualsFormulaEvaluation_setup(formula, cellType);
            ICell D1 = wb.GetSheet("IFEquals").GetRow(0).GetCell(3);

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            CellValue result = eval.Evaluate(D1);

            // Call should not modify the contents
            Assert.AreEqual(CellType.Formula, D1.CellType);
            Assert.AreEqual(expectedFormula, D1.CellFormula);

            Assert.AreEqual(CellType.Numeric, result.CellType);
            Assert.AreEqual(expectedResult, result.NumberValue, EPSILON);

            TestIFEqualsFormulaEvaluation_teardown(wb);
        }

        private void TestIFEqualsFormulaEvaluation_eval(
                String formula, CellType cellType, String expectedFormula, double expectedValue)
        {
            TestIFEqualsFormulaEvaluation_evaluate(formula, cellType, expectedFormula, expectedValue);
            TestIFEqualsFormulaEvaluation_evaluateFormulaCell(formula, cellType, expectedFormula, expectedValue);
            TestIFEqualsFormulaEvaluation_evaluateInCell(formula, cellType, expectedFormula, expectedValue);
            TestIFEqualsFormulaEvaluation_evaluateAll(formula, cellType, expectedFormula, expectedValue);
            TestIFEqualsFormulaEvaluation_evaluateAllFormulaCells(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_NumericLiteral()
        {
            String formula = "IF(A1=1, 2, 3)";
            CellType cellType = CellType.Numeric;
            String expectedFormula = "IF(A1=1,2,3)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_Numeric()
        {
            String formula = "IF(A1=1, B1, C1)";
            CellType cellType = CellType.Numeric;
            String expectedFormula = "IF(A1=1,B1,C1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_NumericCoerceToString()
        {
            String formula = "IF(A1&\"\"=\"1\", B1, C1)";
            CellType cellType = CellType.Numeric;
            String expectedFormula = "IF(A1&\"\"=\"1\",B1,C1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_String()
        {
            String formula = "IF(A1=1, B1, C1)";
            CellType cellType = CellType.String;
            String expectedFormula = "IF(A1=1,B1,C1)";
            double expectedValue = 3.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_StringCompareToString()
        {
            String formula = "IF(A1=\"1\", B1, C1)";
            CellType cellType = CellType.String;
            String expectedFormula = "IF(A1=\"1\",B1,C1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_StringCoerceToNumeric()
        {
            String formula = "IF(A1+0=1, B1, C1)";
            CellType cellType = CellType.String;
            String expectedFormula = "IF(A1+0=1,B1,C1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Ignore("Bug 58591: this test currently Assert.Fails")]
        [Test]
        public void TestIFEqualsFormulaEvaluation_Boolean()
        {
            String formula = "IF(A1=1, B1, C1)";
            CellType cellType = CellType.Boolean;
            String expectedFormula = "IF(A1=1,B1,C1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Ignore("Bug 58591: this test currently Assert.Fails")]
        [Test]
        public void TestIFEqualsFormulaEvaluation_BooleanSimple()
        {
            String formula = "3-(A1=1)";
            CellType cellType = CellType.Boolean;
            String expectedFormula = "3-(A1=1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_Formula()
        {
            String formula = "IF(A1=1, B1, C1)";
            CellType cellType = CellType.Formula;
            String expectedFormula = "IF(A1=1,B1,C1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_Blank()
        {
            String formula = "IF(A1=1, B1, C1)";
            CellType cellType = CellType.Blank;
            String expectedFormula = "IF(A1=1,B1,C1)";
            double expectedValue = 3.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Test]
        public void TestIFEqualsFormulaEvaluation_BlankCompareToZero()
        {
            String formula = "IF(A1=0, B1, C1)";
            CellType cellType = CellType.Blank;
            String expectedFormula = "IF(A1=0,B1,C1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Ignore("Bug 58591: this test currently Assert.Fails")]
        [Test]
        public void TestIFEqualsFormulaEvaluation_BlankInverted()
        {
            String formula = "IF(NOT(A1)=1, B1, C1)";
            CellType cellType = CellType.Blank;
            String expectedFormula = "IF(NOT(A1)=1,B1,C1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }

        [Ignore("Bug 58591: this test currently Assert.Fails")]
        [Test]
        public void TestIFEqualsFormulaEvaluation_BlankInvertedSimple()
        {
            String formula = "3-(NOT(A1)=1)";
            CellType cellType = CellType.Blank;
            String expectedFormula = "3-(NOT(A1)=1)";
            double expectedValue = 2.0;
            TestIFEqualsFormulaEvaluation_eval(formula, cellType, expectedFormula, expectedValue);
        }


        private void TestIFEqualsFormulaEvaluation_evaluateFormulaCell(
                String formula, CellType cellType, String expectedFormula, double expectedResult)
        {
            IWorkbook wb = TestIFEqualsFormulaEvaluation_setup(formula, cellType);
            ICell D1 = wb.GetSheet("IFEquals").GetRow(0).GetCell(3);

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            CellType resultCellType = eval.EvaluateFormulaCell(D1);

            // Call should modify the contents, but leave the formula intact
            Assert.AreEqual(CellType.Formula, D1.CellType);
            Assert.AreEqual(expectedFormula, D1.CellFormula);
            Assert.AreEqual(CellType.Numeric, resultCellType);
            Assert.AreEqual(CellType.Numeric, D1.CachedFormulaResultType);
            Assert.AreEqual(expectedResult, D1.NumericCellValue, EPSILON);

            TestIFEqualsFormulaEvaluation_teardown(wb);
        }

        private void TestIFEqualsFormulaEvaluation_evaluateInCell(
                String formula, CellType cellType, String expectedFormula, double expectedResult)
        {
            IWorkbook wb = TestIFEqualsFormulaEvaluation_setup(formula, cellType);
            ICell D1 = wb.GetSheet("IFEquals").GetRow(0).GetCell(3);

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            ICell result = eval.EvaluateInCell(D1);

            // Call should modify the contents and replace the formula with the result
            Assert.AreSame(D1, result); // returns the same cell that was provided as an argument so that calls can be chained.
            try
            {
                string tmp = D1.CellFormula;
                Assert.Fail("cell formula should be overwritten with formula result");
            }
            catch (InvalidOperationException)
            {
            }
            Assert.AreEqual(CellType.Numeric, D1.CellType);
            Assert.AreEqual(expectedResult, D1.NumericCellValue, EPSILON);

            TestIFEqualsFormulaEvaluation_teardown(wb);
        }

        private void TestIFEqualsFormulaEvaluation_evaluateAll(
                String formula, CellType cellType, String expectedFormula, double expectedResult)
        {
            IWorkbook wb = TestIFEqualsFormulaEvaluation_setup(formula, cellType);
            ICell D1 = wb.GetSheet("IFEquals").GetRow(0).GetCell(3);

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            eval.EvaluateAll();

            // Call should modify the contents
            Assert.AreEqual(CellType.Formula, D1.CellType);
            Assert.AreEqual(expectedFormula, D1.CellFormula);

            Assert.AreEqual(CellType.Numeric, D1.CachedFormulaResultType);
            Assert.AreEqual(expectedResult, D1.NumericCellValue, EPSILON);

            TestIFEqualsFormulaEvaluation_teardown(wb);
        }

        private void TestIFEqualsFormulaEvaluation_evaluateAllFormulaCells(
                String formula, CellType cellType, String expectedFormula, double expectedResult)
        {
            IWorkbook wb = TestIFEqualsFormulaEvaluation_setup(formula, cellType);
            ICell D1 = wb.GetSheet("IFEquals").GetRow(0).GetCell(3);

            HSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);

            // Call should modify the contents
            Assert.AreEqual(CellType.Formula, D1.CellType);
            // whitespace gets deleted because formula is parsed and re-rendered
            Assert.AreEqual(expectedFormula, D1.CellFormula);

            Assert.AreEqual(CellType.Numeric, D1.CachedFormulaResultType);
            Assert.AreEqual(expectedResult, D1.NumericCellValue, EPSILON);

            TestIFEqualsFormulaEvaluation_teardown(wb);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestRefToBlankCellInArrayFormula()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cellA1 = row.CreateCell(0);
            ICell cellB1 = row.CreateCell(1);
            ICell cellC1 = row.CreateCell(2);
            IRow row2 = sheet.CreateRow(1);
            ICell cellA2 = row2.CreateCell(0);
            ICell cellB2 = row2.CreateCell(1);
            ICell cellC2 = row2.CreateCell(2);
            IRow row3 = sheet.CreateRow(2);
            ICell cellA3 = row3.CreateCell(0);
            ICell cellB3 = row3.CreateCell(1);
            ICell cellC3 = row3.CreateCell(2);

            cellA1.SetCellValue("1");
            // cell B1 intentionally left blank
            cellC1.SetCellValue("3");

            cellA2.SetCellFormula("A1");
            cellB2.SetCellFormula("B1");
            cellC2.SetCellFormula("C1");

            sheet.SetArrayFormula("A1:C1", CellRangeAddress.ValueOf("A3:C3"));

            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            Assert.AreEqual(cellA2.StringCellValue, "1");
            Assert.AreEqual(cellB2.NumericCellValue, 0, 0.00001);
            Assert.AreEqual(cellC2.StringCellValue, "3");

            Assert.AreEqual(cellA3.StringCellValue, "1");
            Assert.AreEqual(cellB3.NumericCellValue, 0, 0.00001);
            Assert.AreEqual(cellC3.StringCellValue, "3");
        }
    }
}