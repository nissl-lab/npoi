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

namespace NPOI.SS.Formula.atp;






using junit.framework.Assert;
using junit.framework.AssertionFailedError;
using junit.framework.ComparisonFailure;
using junit.framework.TestCase;

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.SS.Formula.Eval.EvaluationException;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFDateUtil;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFRow;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;

/**
 * Tests YearFracCalculator using Test-cases listed in a sample spreadsheet
 * 
 * @author Josh Micich
 */
public class TestYearFracCalculatorFromSpreadsheet  {
	
	private static class SS {

		public static int BASIS_COLUMN = 1; // "B"
		public static int START_YEAR_COLUMN = 2; // "C"
		public static int END_YEAR_COLUMN = 5; // "F"
		public static int YEARFRAC_FORMULA_COLUMN = 11; // "L"
		public static int EXPECTED_RESULT_COLUMN = 13; // "N"
	}

	public void TestAll() {
		
		HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("yearfracExamples.xls");
		HSSFSheet sheet = wb.GetSheetAt(0);
		HSSFFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(wb);
		int nSuccess = 0;
		int nFailures = 0;
		int nUnexpectedErrors = 0;
		Iterator rowIterator = sheet.rowIterator();
		while(rowIterator.HasNext()) {
			HSSFRow row = (HSSFRow) rowIterator.next();
			
			HSSFCell cell = row.GetCell(SS.YEARFRAC_FORMULA_COLUMN);
			if (cell == null || cell.GetCellType() != HSSFCell.CELL_TYPE_FORMULA) {
				continue;
			}
			try {
				ProcessRow(row, cell, formulaEvaluator);
				nSuccess++;
			} catch (RuntimeException e) {
				nUnexpectedErrors ++;
				printshortStackTrace(System.err, e);
			} catch (AssertionFailedError e) {
				nFailures ++;
				printshortStackTrace(System.err, e);
			}
		}
		if (nUnexpectedErrors + nFailures > 0) {
			String msg = nFailures + " failures(s) and " + nUnexpectedErrors 
				+ " unexpected errors(s) occurred. See stderr for details";
			throw new AssertionFailedError(msg);
		}
		if (nSuccess < 1) {
			throw new RuntimeException("No Test sample cases found");
		}
	}
	
	private static void ProcessRow(HSSFRow row, HSSFCell cell, HSSFFormulaEvaluator formulaEvaluator) {
		
		double startDate = MakeDate(row, SS.START_YEAR_COLUMN);
		double endDate = MakeDate(row, SS.END_YEAR_COLUMN);
		
		int basis = GetIntCell(row, SS.BASIS_COLUMN);
		
		double expectedValue = GetDoubleCell(row, SS.EXPECTED_RESULT_COLUMN);
		
		double actualValue;
		try {
			actualValue = YearFracCalculator.calculate(startDate, endDate, basis);
		} catch (EvaluationException e) {
			throw new RuntimeException(e);
		}
		if (expectedValue != actualValue) {
			throw new ComparisonFailure("Direct calculate failed - row " + (row.GetRowNum()+1), 
					String.ValueOf(expectedValue), String.ValueOf(actualValue));
		}
		actualValue = formulaEvaluator.Evaluate(cell).GetNumberValue();
		if (expectedValue != actualValue) {
			throw new ComparisonFailure("Formula Evaluate failed - row " + (row.GetRowNum()+1), 
					String.ValueOf(expectedValue), String.ValueOf(actualValue));
		}
	}

	private static double MakeDate(HSSFRow row, int yearColumn) {
		int year = GetIntCell(row, yearColumn + 0);
		int month = GetIntCell(row, yearColumn + 1);
		int day = GetIntCell(row, yearColumn + 2);
		Calendar c = new GregorianCalendar(year, month-1, day, 0, 0, 0);
		c.Set(Calendar.MILLISECOND, 0);
		return HSSFDateUtil.GetExcelDate(c.GetTime());
	}

	private static int GetIntCell(HSSFRow row, int colIx) {
		double dVal = GetDoubleCell(row, colIx);
		if (Math.floor(dVal) != dVal) {
			throw new RuntimeException("Non integer value (" + dVal 
					+ ") cell found at column " + (char)('A' + colIx));
		}
		return (int)dVal;
	}

	private static double GetDoubleCell(HSSFRow row, int colIx) {
		HSSFCell cell = row.GetCell(colIx);
		if (cell == null) {
			throw new RuntimeException("No cell found at column " + colIx);
		}
		double dVal = cell.GetNumericCellValue();
		return dVal;
	}

	/**
	 * Useful to keep output concise when expecting many failures to be reported by this Test case
	 * TODO - refactor duplicates in other Test~FromSpreadsheet classes
	 */
	private static void printshortStackTrace(PrintStream ps, Throwable e) {
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

