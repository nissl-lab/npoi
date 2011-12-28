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

namespace NPOI.SS.Formula.functions;



using junit.framework.Assert;
using junit.framework.AssertionFailedError;
using junit.framework.TestCase;

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFRow;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.hssf.Util.CellReference;
using NPOI.SS.UserModel.CellValue;

/**
 * Tests INDEX() as loaded from a Test data spreadsheet.<p/>
 *
 * @author Josh Micich
 */
public class TestIndexFunctionFromSpreadsheet  {

	private static class Result {
		public static int SOME_EVALUATIONS_FAILED = -1;
		public static int ALL_EVALUATIONS_SUCCEEDED = +1;
		public static int NO_EVALUATIONS_FOUND = 0;
	}

	/**
	 * This class defines constants for navigating around the Test data spreadsheet used for these Tests.
	 */
	private static class SS {

		/** Name of the Test spreadsheet (found in the standard Test data folder) */
		public static String FILENAME = "IndexFunctionTestCaseData.xls";

		public static int COLUMN_INDEX_EVALUATION = 2; // Column 'C'
		public static int COLUMN_INDEX_EXPECTED_RESULT = 3; // Column 'D'

	}

	// Note - multiple failures are aggregated before ending.
	// If one or more functions fail, a single AssertionFailedError is thrown at the end
	private int _EvaluationFailureCount;
	private int _EvaluationSuccessCount;



	private static void ConfirmExpectedResult(String msg, HSSFCell expected, CellValue actual) {
		if (expected == null) {
			throw new AssertionFailedError(msg + " - Bad Setup data expected value is null");
		}
		if(actual == null) {
			throw new AssertionFailedError(msg + " - actual value was null");
		}
		if(expected.GetCellType() == HSSFCell.CELL_TYPE_ERROR) {
			ConfirmErrorResult(msg, expected.GetErrorCellValue(), actual);
			return;
		}
		if(actual.GetCellType() == HSSFCell.CELL_TYPE_ERROR) {
			throw unexpectedError(msg, expected, actual.GetErrorValue());
		}
		if(actual.GetCellType() != expected.GetCellType()) {
			throw wrongTypeError(msg, expected, actual);
		}


		switch (expected.GetCellType()) {
			case HSSFCell.CELL_TYPE_BOOLEAN:
				Assert.AreEqual(msg, expected.GetBooleanCellValue(), actual.GetBooleanValue());
				break;
			case HSSFCell.CELL_TYPE_FORMULA: // will never be used, since we will call method After formula Evaluation
				throw new AssertionFailedError("Cannot expect formula as result of formula Evaluation: " + msg);
			case HSSFCell.CELL_TYPE_NUMERIC:
				Assert.AreEqual(msg, expected.GetNumericCellValue(), actual.GetNumberValue(), 0.0);
				break;
			case HSSFCell.CELL_TYPE_STRING:
				Assert.AreEqual(msg, expected.GetRichStringCellValue().GetString(), actual.StringValue);
				break;
		}
	}


	private static AssertionFailedError wrongTypeError(String msgPrefix, HSSFCell expectedCell, CellValue actualValue) {
		return new AssertionFailedError(msgPrefix + " Result type mismatch. Evaluated result was "
				+ actualValue.formatAsString()
				+ " but the expected result was "
				+ formatValue(expectedCell)
				);
	}
	private static AssertionFailedError unexpectedError(String msgPrefix, HSSFCell expected, int actualErrorCode) {
		return new AssertionFailedError(msgPrefix + " Error code ("
				+ ErrorEval.GetText(actualErrorCode)
				+ ") was Evaluated, but the expected result was "
				+ formatValue(expected)
				);
	}


	private static void ConfirmErrorResult(String msgPrefix, int expectedErrorCode, CellValue actual) {
		if(actual.GetCellType() != HSSFCell.CELL_TYPE_ERROR) {
			throw new AssertionFailedError(msgPrefix + " Expected cell error ("
					+ ErrorEval.GetText(expectedErrorCode) + ") but actual value was "
					+ actual.formatAsString());
		}
		if(expectedErrorCode != actual.GetErrorValue()) {
			throw new AssertionFailedError(msgPrefix + " Expected cell error code ("
					+ ErrorEval.GetText(expectedErrorCode)
					+ ") but actual error code was ("
					+ ErrorEval.GetText(actual.GetErrorValue())
					+ ")");
		}
	}


	private static String formatValue(HSSFCell expecedCell) {
		switch (expecedCell.GetCellType()) {
			case HSSFCell.CELL_TYPE_BLANK: return "<blank>";
			case HSSFCell.CELL_TYPE_BOOLEAN: return String.ValueOf(expecedCell.GetBooleanCellValue());
			case HSSFCell.CELL_TYPE_NUMERIC: return String.ValueOf(expecedCell.GetNumericCellValue());
			case HSSFCell.CELL_TYPE_STRING: return expecedCell.GetRichStringCellValue().GetString();
		}
		throw new RuntimeException("Unexpected cell type of expected value (" + expecedCell.GetCellType() + ")");
	}


	protected void SetUp() {
		_EvaluationFailureCount = 0;
		_EvaluationSuccessCount = 0;
	}

	public void TestFunctionsFromTestSpreadsheet() {
		HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook(SS.FILENAME);

		ProcessTestSheet(workbook, workbook.GetSheetName(0));

		// confirm results
		String successMsg = "There were "
				+ _EvaluationSuccessCount + " function(s) without error";
		if(_EvaluationFailureCount > 0) {
			String msg = _EvaluationFailureCount + " Evaluation(s) failed.  " + successMsg;
			throw new AssertionFailedError(msg);
		}
		if(false) { // normally no output for successful Tests
			Console.WriteLine(getClass().GetName() + ": " + successMsg);
		}
	}

	private void ProcessTestSheet(HSSFWorkbook workbook, String sheetName) {
		HSSFSheet sheet = workbook.GetSheetAt(0);
		HSSFFormulaEvaluator Evaluator = new HSSFFormulaEvaluator(workbook);
		int maxRows = sheet.GetLastRowNum()+1;
		int result = Result.NO_EVALUATIONS_FOUND; // so far

		for(int rowIndex=0; rowIndex<maxRows; rowIndex++) {
			HSSFRow r = sheet.GetRow(rowIndex);
			if(r == null) {
				continue;
			}
			HSSFCell c = r.GetCell(SS.COLUMN_INDEX_EVALUATION);
			if (c == null || c.GetCellType() != HSSFCell.CELL_TYPE_FORMULA) {
				continue;
			}
			HSSFCell expectedValueCell = r.GetCell(SS.COLUMN_INDEX_EXPECTED_RESULT);

			String msgPrefix = formatTestCaseDetails(sheetName, r.GetRowNum(), c);
			try {
				CellValue actualValue = Evaluator.evaluate(c);
				ConfirmExpectedResult(msgPrefix, expectedValueCell, actualValue);
				_EvaluationSuccessCount ++;
				if(result != Result.SOME_EVALUATIONS_FAILED) {
					result = Result.ALL_EVALUATIONS_SUCCEEDED;
				}
			} catch (RuntimeException e) {
				_EvaluationFailureCount ++;
				printshortStackTrace(System.err, e, msgPrefix);
				result = Result.SOME_EVALUATIONS_FAILED;
			} catch (AssertionFailedError e) {
				_EvaluationFailureCount ++;
				printshortStackTrace(System.err, e, msgPrefix);
				result = Result.SOME_EVALUATIONS_FAILED;
			}
		}
	}


	private static String formatTestCaseDetails(String sheetName, int rowIndex, HSSFCell c) {

		StringBuilder sb = new StringBuilder();
		CellReference cr = new CellReference(sheetName, rowIndex, c.ColumnIndex, false, false);
		sb.Append(cr.formatAsString());
		sb.Append(" [formula: ").Append(c.GetCellFormula()).append(" ]");
		return sb.ToString();
	}

	/**
	 * Useful to keep output concise when expecting many failures to be reported by this Test case
	 */
	private static void printshortStackTrace(PrintStream ps, Throwable e, String msgPrefix) {
		System.err.println("Problem with " + msgPrefix);
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
}


