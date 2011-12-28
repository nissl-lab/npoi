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

namespace NPOI.SS.Formula;

using junit.framework.AssertionFailedError;
using junit.framework.TestCase;

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.SS.Formula.PTG.AreaErrPtg;
using NPOI.SS.Formula.PTG.AttrPtg;
using NPOI.SS.Formula.PTG.DeletedArea3DPtg;
using NPOI.SS.Formula.PTG.DeletedRef3DPtg;
using NPOI.SS.Formula.PTG.IntPtg;
using NPOI.SS.Formula.PTG.Ptg;
using NPOI.SS.Formula.PTG.RefErrorPtg;
using NPOI.SS.Formula.Eval.BlankEval;
using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.SS.Formula.Eval.MissingArgEval;
using NPOI.SS.Formula.Eval.NumberEval;
using NPOI.SS.Formula.Eval.ValueEval;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFRow;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.Cell;
using NPOI.SS.UserModel.CellValue;
using NPOI.SS.UserModel.ErrorConstants;
using NPOI.SS.UserModel.FormulaEvaluator;
using NPOI.SS.UserModel.Workbook;

/**
 * Tests {@link WorkbookEvaluator}.
 *
 * @author Josh Micich
 */
public class TestWorkbookEvaluator  {

	private static ValueEval EvaluateFormula(Ptg[] ptgs) {
		OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);
		return new WorkbookEvaluator(null, null, null).EvaluateFormula(ec, ptgs);
	}

	/**
	 * Make sure that the Evaluator can directly handle tAttrSum (instead of relying on re-parsing
	 * the whole formula which Converts tAttrSum to tFuncVar("SUM") )
	 */
	public void TestAttrSum() {

		Ptg[] ptgs = {
			new IntPtg(42),
			AttrPtg.SUM,
		};

		ValueEval result = EvaluateFormula(ptgs);
		Assert.AreEqual(42, ((NumberEval)result).GetNumberValue(), 0.0);
	}

	/**
	 * Make sure that the Evaluator can directly handle (deleted) ref error tokens
	 * (instead of relying on re-parsing the whole formula which Converts these
	 * to the error constant #REF! )
	 */
	public void TestRefErr() {

		ConfirmRefErr(new RefErrorPtg());
		ConfirmRefErr(new AreaErrPtg());
		ConfirmRefErr(new DeletedRef3DPtg(0));
		ConfirmRefErr(new DeletedArea3DPtg(0));
	}
	private static void ConfirmRefErr(Ptg ptg) {
		Ptg[] ptgs = {
			ptg,
		};

		ValueEval result = EvaluateFormula(ptgs);
		Assert.AreEqual(ErrorEval.REF_INVALID, result);
	}

	/**
	 * Make sure that the Evaluator can directly handle tAttrSum (instead of relying on re-parsing
	 * the whole formula which Converts tAttrSum to tFuncVar("SUM") )
	 */
	public void TestMemFunc() {

		Ptg[] ptgs = {
			new IntPtg(42),
			AttrPtg.SUM,
		};

		ValueEval result = EvaluateFormula(ptgs);
		Assert.AreEqual(42, ((NumberEval)result).GetNumberValue(), 0.0);
	}


	public void TestEvaluateMultipleWorkbooks() {
		HSSFWorkbook wbA = HSSFTestDataSamples.OpenSampleWorkbook("multibookFormulaA.xls");
		HSSFWorkbook wbB = HSSFTestDataSamples.OpenSampleWorkbook("multibookFormulaB.xls");

		HSSFFormulaEvaluator EvaluatorA = new HSSFFormulaEvaluator(wbA);
		HSSFFormulaEvaluator EvaluatorB = new HSSFFormulaEvaluator(wbB);

		// Hook up the workbook Evaluators to enable Evaluation of formulas across books
		String[] bookNames = { "multibookFormulaA.xls", "multibookFormulaB.xls", };
		HSSFFormulaEvaluator[] Evaluators = { EvaluatorA, EvaluatorB, };
		HSSFFormulaEvaluator.SetupEnvironment(bookNames, Evaluators);

		HSSFCell cell;

		HSSFSheet aSheet1 = wbA.GetSheetAt(0);
		HSSFSheet bSheet1 = wbB.GetSheetAt(0);

		// Simple case - single link from wbA to wbB
		ConfirmFormula(wbA, 0, 0, 0, "[multibookFormulaB.xls]BSheet1!B1");
		cell = aSheet1.GetRow(0).GetCell(0);
		ConfirmEvaluation(35, EvaluatorA, cell);


		// more complex case - back link into wbA
		// [wbA]ASheet1!A2 references (among other things) [wbB]BSheet1!B2
		ConfirmFormula(wbA, 0, 1, 0, "[multibookFormulaB.xls]BSheet1!$B$2+2*A3");
		// [wbB]BSheet1!B2 references (among other things) [wbA]AnotherSheet!A1:B2
		ConfirmFormula(wbB, 0, 1, 1, "SUM([multibookFormulaA.xls]AnotherSheet!$A$1:$B$2)+B3");

		cell = aSheet1.GetRow(1).GetCell(0);
		ConfirmEvaluation(264, EvaluatorA, cell);

		// change [wbB]BSheet1!B3 (from 50 to 60)
		HSSFCell cellB3 = bSheet1.GetRow(2).GetCell(1);
		cellB3.SetCellValue(60);
		EvaluatorB.notifyUpdateCell(cellB3);
		ConfirmEvaluation(274, EvaluatorA, cell);

		// change [wbA]ASheet1!A3 (from 100 to 80)
		HSSFCell cellA3 = aSheet1.GetRow(2).GetCell(0);
		cellA3.SetCellValue(80);
		EvaluatorA.notifyUpdateCell(cellA3);
		ConfirmEvaluation(234, EvaluatorA, cell);

		// change [wbA]AnotherSheet!A1 (from 2 to 3)
		HSSFCell cellA1 = wbA.GetSheetAt(1).GetRow(0).GetCell(0);
		cellA1.SetCellValue(3);
		EvaluatorA.notifyUpdateCell(cellA1);
		ConfirmEvaluation(235, EvaluatorA, cell);
	}

	private static void ConfirmEvaluation(double expectedValue, HSSFFormulaEvaluator fe, HSSFCell cell) {
		Assert.AreEqual(expectedValue, fe.Evaluate(cell).GetNumberValue(), 0.0);
	}

	private static void ConfirmFormula(HSSFWorkbook wb, int sheetIndex, int rowIndex, int columnIndex,
			String expectedFormula) {
		HSSFCell cell = wb.GetSheetAt(sheetIndex).GetRow(rowIndex).GetCell(columnIndex);
		Assert.AreEqual(expectedFormula, cell.GetCellFormula());
	}

	/**
	 * This Test Makes sure that any {@link MissingArgEval} that propagates to
	 * the result of a function Gets translated to {@link BlankEval}.
	 */
	public void TestMissingArg() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.CreateSheet("Sheet1");
		HSSFRow row = sheet.CreateRow(0);
		HSSFCell cell = row.CreateCell(0);
		cell.SetCellFormula("1+IF(1,,)");
		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		CellValue cv;
		try {
			cv = fe.Evaluate(cell);
		} catch (RuntimeException e) {
			throw new AssertionFailedError("Missing arg result not being handled correctly.");
		}
		Assert.AreEqual(HSSFCell.CELL_TYPE_NUMERIC, cv.GetCellType());
		// Adding blank to 1.0 gives 1.0
		Assert.AreEqual(1.0, cv.GetNumberValue(), 0.0);

		// check with string operand
		cell.SetCellFormula("\"abc\"&IF(1,,)");
		fe.NotifySetFormula(cell);
		cv = fe.Evaluate(cell);
		Assert.AreEqual(HSSFCell.CELL_TYPE_STRING, cv.GetCellType());
		// Adding blank to "abc" gives "abc"
		Assert.AreEqual("abc", cv.StringValue);

		// check CHOOSE()
		cell.SetCellFormula("\"abc\"&CHOOSE(2,5,,9)");
		fe.NotifySetFormula(cell);
		cv = fe.Evaluate(cell);
		Assert.AreEqual(HSSFCell.CELL_TYPE_STRING, cv.GetCellType());
		// Adding blank to "abc" gives "abc"
		Assert.AreEqual("abc", cv.StringValue);
	}

	/**
	 * Functions like IF, INDIRECT, INDEX, OFFSET etc can return AreaEvals which
	 * should be dereferenced by the Evaluator
	 */
	public void TestResultOutsideRange() {
		Workbook wb = new HSSFWorkbook();
		Cell cell = wb.CreateSheet("Sheet1").createRow(0).createCell(0);
		cell.SetCellFormula("D2:D5"); // IF(TRUE,D2:D5,D2) or  OFFSET(D2:D5,0,0) would work too
		FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
		CellValue cv;
		try {
			cv = fe.Evaluate(cell);
		} catch (ArgumentException e) {
			if ("Specified row index (0) is outside the allowed range (1..4)".Equals(e.GetMessage())) {
				throw new AssertionFailedError("Identified bug in result dereferencing");
			}
			throw new RuntimeException(e);
		}
		Assert.AreEqual(Cell.CELL_TYPE_ERROR, cv.GetCellType());
		Assert.AreEqual(ErrorConstants.ERROR_VALUE, cv.GetErrorValue());

		// verify circular refs are still detected properly
		fe.ClearAllCachedResultValues();
		cell.SetCellFormula("OFFSET(A1,0,0)");
		cv = fe.Evaluate(cell);
		Assert.AreEqual(Cell.CELL_TYPE_ERROR, cv.GetCellType());
		Assert.AreEqual(ErrorEval.CIRCULAR_REF_ERROR.GetErrorCode(), cv.GetErrorValue());
	}
}

