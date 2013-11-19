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

    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;

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
        private static ValueEval EvaluateFormula(Ptg[] ptgs)
        {
            OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);
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

            // Hook up the workbook Evaluators to enable Evaluation of formulas across books
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
            CellValue cv;
            try
            {
                cv = fe.Evaluate(cell);
            }
            catch (Exception)
            {
                throw new AssertionException("Missing arg result not being handled correctly.");
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
                    throw new AssertionException("Identified bug in result dereferencing");
                }
                throw;
            }

            Assert.AreEqual(CellType.Error, cv.CellType);
            Assert.AreEqual(ErrorConstants.ERROR_VALUE, cv.ErrorValue);

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
    }
}