/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
    using TestCases.HSSF;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.HSSF.Record.Formula.Eval;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.UserModel;

    /**
     * Tests {@link WorkbookEvaluator}.
     *
     * @author Josh Micich
     */
    [TestClass]
    public class TestWorkbookEvaluator
    {
        private static ValueEval EvaluateFormula(Ptg[] ptgs)
        {
            OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);
            return new WorkbookEvaluator(null, null, null).EvaluateFormula(ec, ptgs);
        }
        /**
         * Make sure that the evaluator can directly handle tAttrSum (instead of relying on re-parsing
         * the whole formula which converts tAttrSum to tFuncVar("SUM") )
         */
        [TestMethod]
        public void TestAttrSum()
        {

            Ptg[] ptgs = {
			new IntPtg(42),
			AttrPtg.SUM,
		};

            ValueEval result = EvaluateFormula(ptgs);
            Assert.AreEqual(42, ((NumberEval)result).NumberValue, 0.0);
        }

        /**
         * Make sure that the evaluator can directly handle (deleted) ref error tokens
         * (instead of relying on re-parsing the whole formula which converts these
         * to the error constant #REF! )
         */
        [TestMethod]
        public void TestRefErr()
        {

            ConfirmRefErr(new RefErrorPtg());
            ConfirmRefErr(new AreaErrPtg());
            ConfirmRefErr(new DeletedRef3DPtg(0));
            ConfirmRefErr(new DeletedArea3DPtg(0));
        }
        private static void ConfirmRefErr(Ptg ptg)
        {
            Ptg[] ptgs = {
			ptg,
		};

            ValueEval result = EvaluateFormula(ptgs);
            Assert.AreEqual(ErrorEval.REF_INVALID, result);
        }

        /**
         * Make sure that the evaluator can directly handle tAttrSum (instead of relying on re-parsing
         * the whole formula which converts tAttrSum to tFuncVar("SUM") )
         */
        [TestMethod]
        public void TestMemFunc()
        {

            Ptg[] ptgs = {
			    new IntPtg(42),
			    AttrPtg.SUM,
		    };

            ValueEval result = EvaluateFormula(ptgs);
            Assert.AreEqual(42, ((NumberEval)result).NumberValue, 0.0);
        }

        [TestMethod]
        public void TestEvaluateMultipleWorkbooks()
        {
            HSSFWorkbook wbA = HSSFTestDataSamples.OpenSampleWorkbook("multibookFormulaA.xls");
            HSSFWorkbook wbB = HSSFTestDataSamples.OpenSampleWorkbook("multibookFormulaB.xls");

            HSSFFormulaEvaluator evaluatorA = new HSSFFormulaEvaluator(wbA);
            HSSFFormulaEvaluator evaluatorB = new HSSFFormulaEvaluator(wbB);

            // Hook up the workbook evaluators to enable evaluation of formulas across books
            String[] bookNames = { "multibookFormulaA.xls", "multibookFormulaB.xls", };
            HSSFFormulaEvaluator[] evaluators = { evaluatorA, evaluatorB, };
            HSSFFormulaEvaluator.SetupEnvironment(bookNames, evaluators);

            Cell cell;

            NPOI.SS.UserModel.Sheet aSheet1 = wbA.GetSheetAt(0);
            NPOI.SS.UserModel.Sheet bSheet1 = wbB.GetSheetAt(0);

            // Simple case - single link from wbA to wbB
            ConfirmFormula(wbA, 0, 0, 0, "[multibookFormulaB.xls]BSheet1!B1");
            cell = aSheet1.GetRow(0).GetCell(0);
            ConfirmEvaluation(35, evaluatorA, cell);


            // more complex case - back link into wbA
            // [wbA]ASheet1!A2 references (among other things) [wbB]BSheet1!B2
            ConfirmFormula(wbA, 0, 1, 0, "[multibookFormulaB.xls]BSheet1!$B$2+2*A3");
            // [wbB]BSheet1!B2 references (among other things) [wbA]AnotherSheet!A1:B2
            ConfirmFormula(wbB, 0, 1, 1, "SUM([multibookFormulaA.xls]AnotherSheet!$A$1:$B$2)+B3");

            cell = aSheet1.GetRow(1).GetCell(0);
            ConfirmEvaluation(264, evaluatorA, cell);

            // change [wbB]BSheet1!B3 (from 50 to 60)
            Cell cellB3 = bSheet1.GetRow(2).GetCell(1);
            cellB3.SetCellValue(60);
            evaluatorB.NotifyUpdateCell(cellB3);
            ConfirmEvaluation(274, evaluatorA, cell);

            // change [wbA]ASheet1!A3 (from 100 to 80)
            Cell cellA3 = aSheet1.GetRow(2).GetCell(0);
            cellA3.SetCellValue(80);
            evaluatorA.NotifyUpdateCell(cellA3);
            ConfirmEvaluation(234, evaluatorA, cell);

            // change [wbA]AnotherSheet!A1 (from 2 to 3)
            Cell cellA1 = wbA.GetSheetAt(1).GetRow(0).GetCell(0);
            cellA1.SetCellValue(3);
            evaluatorA.NotifyUpdateCell(cellA1);
            ConfirmEvaluation(235, evaluatorA, cell);
        }

        private static void ConfirmEvaluation(double expectedValue, HSSFFormulaEvaluator fe, Cell cell)
        {
            Assert.AreEqual(expectedValue, fe.Evaluate(cell).NumberValue, 0.0);
        }

        private static void ConfirmFormula(HSSFWorkbook wb, int sheetIndex, int rowIndex, int columnIndex,
                String expectedFormula)
        {
            Cell cell = wb.GetSheetAt(sheetIndex).GetRow(rowIndex).GetCell(columnIndex);
            Assert.AreEqual(expectedFormula, cell.CellFormula);
        }
    }
}