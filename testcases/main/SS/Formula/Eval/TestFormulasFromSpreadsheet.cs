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



using junit.framework.Assert;
using junit.framework.AssertionFailedError;
using junit.framework.TestCase;

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.SS.Formula.functions.TestMathX;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.Cell;
using NPOI.SS.UserModel.CellValue;
using NPOI.SS.UserModel.Row;
using NPOI.SS.UserModel.Sheet;

/**
 * Tests formulas and operators as loaded from a Test data spreadsheet.<p/>
 * This class does not Test implementors of <tt>Function</tt> and <tt>OperationEval</tt> in
 * isolation.  Much of the Evaluation engine (i.e. <tt>HSSFFormulaEvaluator</tt>, ...) Gets
 * exercised as well.  Tests for bug fixes and specific/tricky behaviour can be found in the
 * corresponding Test class (<tt>TestXxxx</tt>) of the target (<tt>Xxxx</tt>) implementor,
 * where execution can be observed more easily.
 *
 * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
 */
public class TestFormulasFromSpreadsheet  {

	private static class Result {
		public static int SOME_EVALUATIONS_FAILED = -1;
		public static int ALL_EVALUATIONS_SUCCEEDED = +1;
		public static int NO_EVALUATIONS_FOUND = 0;
	}

	/**
	 * This class defines constants for navigating around the Test data spreadsheet used for these Tests.
	 */
	private static class SS {

		/**
		 * Name of the Test spreadsheet (found in the standard Test data folder)
		 */
		public static String FILENAME = "FormulaEvalTestData.xls";
		/**
		 * Row (zero-based) in the Test spreadsheet where the operator examples start.
		 */
		public static int START_OPERATORS_ROW_INDEX = 22; // Row '23'
		/**
		 * Row (zero-based) in the Test spreadsheet where the function examples start.
		 */
		public static int START_FUNCTIONS_ROW_INDEX = 95; // Row '96'
		/**
		 * Index of the column that Contains the function names
		 */
		public static int COLUMN_INDEX_FUNCTION_NAME = 1; // Column 'B'

		/**
		 * Used to indicate when there are no more functions left
		 */
		public static String FUNCTION_NAMES_END_SENTINEL = "<END-OF-FUNCTIONS>";

		/**
		 * Index of the column where the Test values start (for each function)
		 */
		public static short COLUMN_INDEX_FIRST_TEST_VALUE = 3; // Column 'D'

		/**
		 * Each function takes 4 rows in the Test spreadsheet
		 */
		public static int NUMBER_OF_ROWS_PER_FUNCTION = 4;
	}

	private HSSFWorkbook workbook;
	private Sheet sheet;
	// Note - multiple failures are aggregated before ending.
	// If one or more functions fail, a single AssertionFailedError is thrown at the end
	private int _functionFailureCount;
	private int _functionSuccessCount;
	private int _EvaluationFailureCount;
	private int _EvaluationSuccessCount;

	private static Cell GetExpectedValueCell(Row row, int columnIndex) {
		if (row == null) {
			return null;
		}
		return row.GetCell(columnIndex);
	}


	private static void ConfirmExpectedResult(String msg, Cell expected, CellValue actual) {
		if (expected == null) {
			throw new AssertionFailedError(msg + " - Bad Setup data expected value is null");
		}
		if(actual == null) {
			throw new AssertionFailedError(msg + " - actual value was null");
		}

		switch (expected.GetCellType()) {
			case Cell.CELL_TYPE_BLANK:
				Assert.AreEqual(msg, Cell.CELL_TYPE_BLANK, actual.GetCellType());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				Assert.AreEqual(msg, Cell.CELL_TYPE_BOOLEAN, actual.GetCellType());
				Assert.AreEqual(msg, expected.GetBooleanCellValue(), actual.GetBooleanValue());
				break;
			case Cell.CELL_TYPE_ERROR:
				Assert.AreEqual(msg, Cell.CELL_TYPE_ERROR, actual.GetCellType());
				Assert.AreEqual(msg, ErrorEval.GetText(expected.GetErrorCellValue()), ErrorEval.GetText(actual.GetErrorValue()));
				break;
			case Cell.CELL_TYPE_FORMULA: // will never be used, since we will call method After formula Evaluation
				throw new AssertionFailedError("Cannot expect formula as result of formula Evaluation: " + msg);
			case Cell.CELL_TYPE_NUMERIC:
				Assert.AreEqual(msg, Cell.CELL_TYPE_NUMERIC, actual.GetCellType());
				TestMathX.Assert.AreEqual(msg, expected.GetNumericCellValue(), actual.GetNumberValue(), TestMathX.POS_ZERO, TestMathX.DIFF_TOLERANCE_FACTOR);
				break;
			case Cell.CELL_TYPE_STRING:
				Assert.AreEqual(msg, Cell.CELL_TYPE_STRING, actual.GetCellType());
				Assert.AreEqual(msg, expected.GetRichStringCellValue().GetString(), actual.StringValue);
				break;
		}
	}


	protected void SetUp() {
		if (workbook == null) {
			workbook = HSSFTestDataSamples.OpenSampleWorkbook(SS.FILENAME);
			sheet = workbook.GetSheetAt( 0 );
		}
		_functionFailureCount = 0;
		_functionSuccessCount = 0;
		_EvaluationFailureCount = 0;
		_EvaluationSuccessCount = 0;
	}

	public void TestFunctionsFromTestSpreadsheet() {

		ProcessFunctionGroup(SS.START_OPERATORS_ROW_INDEX, null);
		ProcessFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, null);
		// example for debugging individual functions/operators:
//		ProcessFunctionGroup(SS.START_OPERATORS_ROW_INDEX, "ConcatEval");
//		ProcessFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, "AVERAGE");

		// confirm results
		String successMsg = "There were "
				+ _EvaluationSuccessCount + " successful Evaluation(s) and "
				+ _functionSuccessCount + " function(s) without error";
		if(_functionFailureCount > 0) {
			String msg = _functionFailureCount + " function(s) failed in "
			+ _EvaluationFailureCount + " Evaluation(s).  " + successMsg;
			throw new AssertionFailedError(msg);
		}
		if(false) { // normally no output for successful Tests
			Console.WriteLine(getClass().GetName() + ": " + successMsg);
		}
	}

	/**
	 * @param startRowIndex row index in the spreadsheet where the first function/operator is found
	 * @param TestFocusFunctionName name of a single function/operator to Test alone.
	 * Typically pass <code>null</code> to Test all functions
	 */
	private void ProcessFunctionGroup(int startRowIndex, String TestFocusFunctionName) {
		HSSFFormulaEvaluator Evaluator = new HSSFFormulaEvaluator(workbook);

		int rowIndex = startRowIndex;
		while (true) {
			Row r = sheet.GetRow(rowIndex);
			String targetFunctionName = GetTargetFunctionName(r);
			if(targetFunctionName == null) {
				throw new AssertionFailedError("Test spreadsheet cell empty on row ("
						+ (rowIndex+1) + "). Expected function name or '"
						+ SS.FUNCTION_NAMES_END_SENTINEL + "'");
			}
			if(targetFunctionName.Equals(SS.FUNCTION_NAMES_END_SENTINEL)) {
				// found end of functions list
				break;
			}
			if(testFocusFunctionName == null || targetFunctionName.EqualsIgnoreCase(testFocusFunctionName)) {

				// expected results are on the row below
				Row expectedValuesRow = sheet.GetRow(rowIndex + 1);
				if(expectedValuesRow == null) {
					int missingRowNum = rowIndex + 2; //+1 for 1-based, +1 for next row
					throw new AssertionFailedError("Missing expected values row for function '"
							+ targetFunctionName + " (row " + missingRowNum + ")");
				}
				switch(ProcessFunctionRow(evaluator, targetFunctionName, r, expectedValuesRow)) {
					case Result.ALL_EVALUATIONS_SUCCEEDED: _functionSuccessCount++; break;
					case Result.SOME_EVALUATIONS_FAILED: _functionFailureCount++; break;
					default:
						throw new RuntimeException("unexpected result");
					case Result.NO_EVALUATIONS_FOUND: // do nothing
				}
			}
			rowIndex += SS.NUMBER_OF_ROWS_PER_FUNCTION;
		}
	}

	/**
	 *
	 * @return a constant from the local Result class denoting whether there were any Evaluation
	 * cases, and whether they all succeeded.
	 */
	private int ProcessFunctionRow(HSSFFormulaEvaluator Evaluator, String targetFunctionName,
			Row formulasRow, Row expectedValuesRow) {

		int result = Result.NO_EVALUATIONS_FOUND; // so far
		short endcolnum = formulasRow.GetLastCellNum();

		// iterate across the row for all the Evaluation cases
		for (int colnum=SS.COLUMN_INDEX_FIRST_TEST_VALUE; colnum < endcolnum; colnum++) {
			Cell c = formulasRow.GetCell(colnum);
			if (c == null || c.GetCellType() != Cell.CELL_TYPE_FORMULA) {
				continue;
			}

			CellValue actualValue = Evaluator.evaluate(c);

			Cell expectedValueCell = GetExpectedValueCell(expectedValuesRow, colnum);
			try {
				ConfirmExpectedResult("Function '" + targetFunctionName + "': Formula: " + c.GetCellFormula() + " @ " + formulasRow.GetRowNum() + ":" + colnum,
						expectedValueCell, actualValue);
				_EvaluationSuccessCount ++;
				if(result != Result.SOME_EVALUATIONS_FAILED) {
					result = Result.ALL_EVALUATIONS_SUCCEEDED;
				}
			} catch (AssertionFailedError e) {
				_EvaluationFailureCount ++;
				printshortStackTrace(System.err, e);
				result = Result.SOME_EVALUATIONS_FAILED;
			}
		}
		return result;
	}

	/**
	 * Useful to keep output concise when expecting many failures to be reported by this Test case
	 */
	private static void printshortStackTrace(PrintStream ps, AssertionFailedError e) {
		StackTraceElement[] stes = e.GetStackTrace();

		int startIx = 0;
		// skip any top frames inside junit.framework.Assert
		while(startIx<stes.Length) {
			if(!stes[startIx].GetClassName().Equals(Assert.class.GetName())) {
				break;
			}
			startIx++;
		}
		// skip bottom frames (part of junit framework)
		int endIx = startIx+1;
		while(endIx < stes.Length) {
			if(stes[endIx].GetClassName().Equals(TestCase.class.GetName())) {
				break;
			}
			endIx++;
		}
		if(startIx >= endIx) {
			// something went wrong. just print the whole stack trace
			e.printStackTrace(ps);
		}
		endIx -= 4; // skip 4 frames of reflection invocation
		ps.println(e.ToString());
		for(int i=startIx; i<endIx; i++) {
			ps.println("\tat " + stes[i].ToString());
		}
	}

	/**
	 * @return <code>null</code> if cell is missing, empty or blank
	 */
	private static String GetTargetFunctionName(Row r) {
		if(r == null) {
			System.err.println("Warning - given null row, can't figure out function name");
			return null;
		}
		Cell cell = r.GetCell(SS.COLUMN_INDEX_FUNCTION_NAME);
		if(cell == null) {
			System.err.println("Warning - Row " + r.GetRowNum() + " has no cell " + SS.COLUMN_INDEX_FUNCTION_NAME + ", can't figure out function name");
			return null;
		}
		if(cell.GetCellType() == Cell.CELL_TYPE_BLANK) {
			return null;
		}
		if(cell.GetCellType() == Cell.CELL_TYPE_STRING) {
			return cell.GetRichStringCellValue().GetString();
		}

		throw new AssertionFailedError("Bad cell type for 'function name' column: ("
				+ cell.GetCellType() + ") row (" + (r.GetRowNum() +1) + ")");
	}
}

