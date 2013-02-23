/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/

namespace NPOI.XSSF.UserModel;




using junit.framework.Assert;
using junit.framework.AssertionFailedError;
using junit.framework.TestCase;

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.ss.formula.Eval.TestFormulasFromSpreadsheet;
using NPOI.ss.formula.functions.TestMathX;
using NPOI.ss.UserModel.Cell;
using NPOI.ss.UserModel.CellValue;
using NPOI.ss.UserModel.FormulaEvaluator;
using NPOI.ss.UserModel.Row;
using NPOI.ss.UserModel.Sheet;
using NPOI.Openxml4j.opc.OPCPackage;

/**
 * Performs much the same role as {@link TestFormulasFromSpreadsheet},
 *  except for a XSSF spreadsheet, not a HSSF one.
 * This allows us to check that all our Formula Evaluation code
 *  is able to work for XSSF, as well as for HSSF.
 * 
 * Periodically, you should open FormulaEvalTestData.xls in
 *  Excel 2007, and re-save it as FormulaEvalTestData_Copy.xlsx
 *  
 */
public class TestFormulaEvaluatorOnXSSF  {
	
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
		public static String FILENAME = "FormulaEvalTestData_Copy.xlsx";
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

	private XSSFWorkbook workbook;
	private Sheet sheet;
	// Note - multiple failures are aggregated before ending.  
	// If one or more functions fail, a single AssertionFailedError is thrown at the end
	private int _functionFailureCount;
	private int _functionSuccessCount;
	private int _EvaluationFailureCount;
	private int _EvaluationSuccessCount;

	private static Cell GetExpectedValueCell(Row row, short columnIndex) {
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
				if(false) { // TODO: fix ~45 functions which are currently returning incorrect error values
					Assert.AreEqual(msg, expected.GetErrorCellValue(), actual.GetErrorValue());
				}
				break;
			case Cell.CELL_TYPE_FORMULA: // will never be used, since we will call method After formula Evaluation
				throw new AssertionFailedError("Cannot expect formula as result of formula Evaluation: " + msg);
			case Cell.CELL_TYPE_NUMERIC:
				Assert.AreEqual(msg, Cell.CELL_TYPE_NUMERIC, actual.GetCellType());
				TestMathX.Assert.AreEqual(msg, expected.GetNumericCellValue(), actual.GetNumberValue(), TestMathX.POS_ZERO, TestMathX.DIFF_TOLERANCE_FACTOR);
//				double delta = Math.abs(expected.GetNumericCellValue()-actual.GetNumberValue());
//				double pctExpected = Math.abs(0.00001*expected.GetNumericCellValue());
//				Assert.IsTrue(msg, delta <= pctExpected);
				break;
			case Cell.CELL_TYPE_STRING:
				Assert.AreEqual(msg, Cell.CELL_TYPE_STRING, actual.GetCellType());
				Assert.AreEqual(msg, expected.GetRichStringCellValue().GetString(), actual.StringValue);
				break;
		}
	}


	protected void SetUp() throws Exception {
		if (workbook == null) {
			InputStream is = HSSFTestDataSamples.OpenSampleFileStream(SS.FILENAME);
			OPCPackage pkg = OPCPackage.Open(is);
			workbook = new XSSFWorkbook( pkg );
			sheet = workbook.GetSheetAt( 0 );
		  }
		_functionFailureCount = 0;
		_functionSuccessCount = 0;
		_EvaluationFailureCount = 0;
		_EvaluationSuccessCount = 0;
	}
	
	/**
	 * Checks that we can actually open the file
	 */
	public void TestOpen() {
		Assert.IsNotNull(workbook);
	}
	
	/**
	 * Disabled for now, as many things seem to break
	 *  for XSSF, which is a shame
	 */
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
 
		FormulaEvaluator Evaluator = new XSSFFormulaEvaluator(workbook);

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
	private int ProcessFunctionRow(FormulaEvaluator Evaluator, String targetFunctionName, 
			Row formulasRow, Row expectedValuesRow) {
		
		int result = Result.NO_EVALUATIONS_FOUND; // so far
		short endcolnum = formulasRow.GetLastCellNum();

		// iterate across the row for all the Evaluation cases
		for (short colnum=SS.COLUMN_INDEX_FIRST_TEST_VALUE; colnum < endcolnum; colnum++) {
			Cell c = formulasRow.GetCell(colnum);
			if (c == null || c.GetCellType() != Cell.CELL_TYPE_FORMULA) {
				continue;
			}
			if(isIgnoredFormulaTestCase(c.GetCellFormula())) {
				continue;
			}

			CellValue actualValue;
			try {
				actualValue = Evaluator.evaluate(c);
			} catch (RuntimeException e) {
				_EvaluationFailureCount ++;
				printshortStackTrace(System.err, e);
				result = Result.SOME_EVALUATIONS_FAILED;
				continue;
			}

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

	/*
	 * TODO - these are all formulas which currently (Apr-2008) break on ooxml 
	 */
	private static bool IsIgnoredFormulaTestCase(String cellFormula) {
		if ("COLUMN(1:2)".Equals(cellFormula) || "ROW(2:3)".Equals(cellFormula)) {
			// full row ranges are not Parsed properly yet.
			// These cases currently work in svn tRunk because of another bug which causes the 
			// formula to Get rendered as COLUMN($A$1:$IV$2) or ROW($A$2:$IV$3) 
			return true;
		}
		if ("ISREF(currentcell())".Equals(cellFormula)) {
			// currently throws NPE because unknown function "currentcell" causes name lookup 
			// Name lookup requires some equivalent object of the Workbook within xSSFWorkbook.
			return true;
		}
		return false;
	}


	/**
	 * Useful to keep output concise when expecting many failures to be reported by this Test case
	 */
	private static void printshortStackTrace(PrintStream ps, Throwable e) {
		StackTraceElement[] stes = e.GetStackTrace();
		
		int startIx = 0;
		// Skip any top frames inside junit.framework.Assert
		while(startIx<stes.Length) {
			if(!stes[startIx].GetClassName().Equals(Assert.class.GetName())) {
				break;
			}
			startIx++;
		}
		// Skip bottom frames (part of junit framework)
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
		endIx -= 4; // Skip 4 frames of reflection invocation
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

