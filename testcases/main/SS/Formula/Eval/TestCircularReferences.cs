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

namespace NPOI.SS.Formula.Eval;

using junit.framework.AssertionFailedError;
using junit.framework.TestCase;

using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFRow;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.Cell;
using NPOI.SS.UserModel.CellValue;
/**
 * Tests HSSFFormulaEvaluator for its handling of cell formula circular references.
 *
 * @author Josh Micich
 */
public class TestCircularReferences  {
	/**
	 * Translates StackOverflowError into AssertionFailedError
	 */
	private static CellValue EvaluateWithCycles(HSSFWorkbook wb, HSSFCell TestCell)
			throws AssertionFailedError {
		HSSFFormulaEvaluator Evaluator = new HSSFFormulaEvaluator(wb);
		try {
			return Evaluator.evaluate(testCell);
		} catch (StackOverflowError e) {
			throw new AssertionFailedError( "circular reference caused stack overflow error");
		}
	}
	/**
	 * Makes sure that the specified Evaluated cell value represents a circular reference error.
	 */
	private static void ConfirmCycleErrorCode(CellValue cellValue) {
		Assert.IsTrue(cellValue.GetCellType() == HSSFCell.CELL_TYPE_ERROR);
		Assert.AreEqual(ErrorEval.CIRCULAR_REF_ERROR.GetErrorCode(), cellValue.GetErrorValue());
	}


	/**
	 * ASF Bugzilla Bug 44413
	 * "INDEX() formula cannot contain its own location in the data array range"
	 */
	public void TestIndexFormula() {

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.CreateSheet("Sheet1");

		int colB = 1;
		sheet.CreateRow(0).createCell(colB).SetCellValue(1);
		sheet.CreateRow(1).createCell(colB).SetCellValue(2);
		sheet.CreateRow(2).createCell(colB).SetCellValue(3);
		HSSFRow row4 = sheet.CreateRow(3);
		HSSFCell TestCell = row4.CreateCell(0);
		// This formula should Evaluate to the contents of B2,
	 TestCell.SetCellFormula("INDEX(A1:B4,2,2)");
		// However the range A1:B4 also includes the current cell A4.  If the other parameters
		// were 4 and 1, this would represent a circular reference.  Prior to v3.2 POI would
		// 'fully' Evaluate ref arguments before invoking operators, which raised the possibility of
		// cycles / StackOverflowErrors.


		CellValue cellValue = EvaluateWithCycles(wb, TestCell);

		Assert.IsTrue(cellValue.GetCellType() == HSSFCell.CELL_TYPE_NUMERIC);
		Assert.AreEqual(2, cellValue.GetNumberValue(), 0);
	}

	/**
	 * Cell A1 has formula "=A1"
	 */
	public void TestSimpleCircularReference() {

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.CreateSheet("Sheet1");

		HSSFRow row = sheet.CreateRow(0);
		HSSFCell TestCell = row.CreateCell(0);
	 TestCell.SetCellFormula("A1");

		CellValue cellValue = EvaluateWithCycles(wb, TestCell);

		ConfirmCycleErrorCode(cellValue);
	}

	/**
	 * A1=B1, B1=C1, C1=D1, D1=A1
	 */
	public void TestMultiLevelCircularReference() {

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.CreateSheet("Sheet1");

		HSSFRow row = sheet.CreateRow(0);
		row.CreateCell(0).SetCellFormula("B1");
		row.CreateCell(1).SetCellFormula("C1");
		row.CreateCell(2).SetCellFormula("D1");
		HSSFCell TestCell = row.CreateCell(3);
	 TestCell.SetCellFormula("A1");

		CellValue cellValue = EvaluateWithCycles(wb, TestCell);

		ConfirmCycleErrorCode(cellValue);
	}

	public void TestIntermediateCircularReferenceResults_bug46898() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.CreateSheet("Sheet1");

		HSSFRow row = sheet.CreateRow(0);

		HSSFCell cellA1 = row.CreateCell(0);
		HSSFCell cellB1 = row.CreateCell(1);
		HSSFCell cellC1 = row.CreateCell(2);
		HSSFCell cellD1 = row.CreateCell(3);
		HSSFCell cellE1 = row.CreateCell(4);

		cellA1.SetCellFormula("IF(FALSE, 1+B1, 42)");
		cellB1.SetCellFormula("1+C1");
		cellC1.SetCellFormula("1+D1");
		cellD1.SetCellFormula("1+E1");
		cellE1.SetCellFormula("1+A1");

		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		CellValue cv;

		// Happy day flow - Evaluate A1 first
		cv = fe.Evaluate(cellA1);
		Assert.AreEqual(Cell.CELL_TYPE_NUMERIC, cv.GetCellType());
		Assert.AreEqual(42.0, cv.GetNumberValue(), 0.0);
		cv = fe.Evaluate(cellB1); // no circ-ref-error because A1 result is cached
		Assert.AreEqual(Cell.CELL_TYPE_NUMERIC, cv.GetCellType());
		Assert.AreEqual(46.0, cv.GetNumberValue(), 0.0);

		// Show the bug - Evaluate another cell from the loop first
		fe.ClearAllCachedResultValues();
		cv = fe.Evaluate(cellB1);
		if (cv.GetCellType() == ErrorEval.CIRCULAR_REF_ERROR.GetErrorCode()) {
			throw new AssertionFailedError("Identified bug 46898");
		}
		Assert.AreEqual(Cell.CELL_TYPE_NUMERIC, cv.GetCellType());
		Assert.AreEqual(46.0, cv.GetNumberValue(), 0.0);

		// start Evaluation on another cell
		fe.ClearAllCachedResultValues();
		cv = fe.Evaluate(cellE1);
		Assert.AreEqual(Cell.CELL_TYPE_NUMERIC, cv.GetCellType());
		Assert.AreEqual(43.0, cv.GetNumberValue(), 0.0);


	}
}

