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

using junit.framework.AssertionFailedError;
using junit.framework.TestCase;

using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.hssf.UserModel.*;
using NPOI.SS.UserModel.Cell;
using NPOI.SS.UserModel.CellValue;
using NPOI.SS.UserModel.FormulaEvaluator;

/**
 * Tests for the INDIRECT() function.</p>
 *
 * @author Josh Micich
 */
public class TestIndirect  {
	// convenient access to namespace
	private static ErrorEval EE = null;

	private static void CreateDataRow(HSSFSheet sheet, int rowIndex, double... vals) {
		HSSFRow row = sheet.CreateRow(rowIndex);
		for (int i = 0; i < vals.Length; i++) {
			row.CreateCell(i).SetCellValue(vals[i]);
		}
	}

	private static HSSFWorkbook CreateWBA() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet1 = wb.CreateSheet("Sheet1");
		HSSFSheet sheet2 = wb.CreateSheet("Sheet2");
		HSSFSheet sheet3 = wb.CreateSheet("John's sales");

		CreateDataRow(sheet1, 0, 11, 12, 13, 14);
		CreateDataRow(sheet1, 1, 21, 22, 23, 24);
		CreateDataRow(sheet1, 2, 31, 32, 33, 34);

		CreateDataRow(sheet2, 0, 50, 55, 60, 65);
		CreateDataRow(sheet2, 1, 51, 56, 61, 66);
		CreateDataRow(sheet2, 2, 52, 57, 62, 67);

		CreateDataRow(sheet3, 0, 30, 31, 32);
		CreateDataRow(sheet3, 1, 33, 34, 35);

        HSSFName name1 = wb.CreateName();
        name1.SetNameName("sales1");
        name1.SetRefersToFormula("Sheet1!A1:D1");

        HSSFName name2 = wb.CreateName();
        name2.SetNameName("sales2");
        name2.SetRefersToFormula("Sheet2!B1:C3");

        HSSFRow row = sheet1.CreateRow(3);
        row.CreateCell(0).SetCellValue("sales1");  //A4
        row.CreateCell(1).SetCellValue("sales2");  //B4

		return wb;
	}

	private static HSSFWorkbook CreateWBB() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet1 = wb.CreateSheet("Sheet1");
		HSSFSheet sheet2 = wb.CreateSheet("Sheet2");
		HSSFSheet sheet3 = wb.CreateSheet("## Look here!");

		CreateDataRow(sheet1, 0, 400, 440, 480, 520);
		CreateDataRow(sheet1, 1, 420, 460, 500, 540);

		CreateDataRow(sheet2, 0, 50, 55, 60, 65);
		CreateDataRow(sheet2, 1, 51, 56, 61, 66);

		CreateDataRow(sheet3, 0, 42);

		return wb;
	}

	public void TestBasic() {

		HSSFWorkbook wbA = CreateWBA();
		HSSFCell c = wbA.GetSheetAt(0).CreateRow(5).createCell(2);
		HSSFFormulaEvaluator feA = new HSSFFormulaEvaluator(wbA);

		// non-error cases
		Confirm(feA, c, "INDIRECT(\"C2\")", 23);
		Confirm(feA, c, "INDIRECT(\"$C2\")", 23);
		Confirm(feA, c, "INDIRECT(\"C$2\")", 23);
		Confirm(feA, c, "SUM(INDIRECT(\"Sheet2!B1:C3\"))", 351); // area ref
		Confirm(feA, c, "SUM(INDIRECT(\"Sheet2! B1 : C3 \"))", 351); // spaces in area ref
		Confirm(feA, c, "SUM(INDIRECT(\"'John''s sales'!A1:C1\"))", 93); // special chars in sheet name
		Confirm(feA, c, "INDIRECT(\"'Sheet1'!B3\")", 32); // redundant sheet name quotes
		Confirm(feA, c, "INDIRECT(\"sHeet1!B3\")", 32); // case-insensitive sheet name
		Confirm(feA, c, "INDIRECT(\" D3 \")", 34); // spaces around cell ref
		Confirm(feA, c, "INDIRECT(\"Sheet1! D3 \")", 34); // spaces around cell ref
		Confirm(feA, c, "INDIRECT(\"A1\", TRUE)", 11); // explicit arg1. only TRUE supported so far

		Confirm(feA, c, "INDIRECT(\"A1:G1\")", 13); // de-reference area ref (note formula is in C4)

        Confirm(feA, c, "SUM(INDIRECT(A4))", 50); // indirect defined name
        Confirm(feA, c, "SUM(INDIRECT(B4))", 351); // indirect defined name pointinh to other sheet

		// simple error propagation:

		// arg0 is Evaluated to text first
		Confirm(feA, c, "INDIRECT(#DIV/0!)", EE.DIV_ZERO);
		Confirm(feA, c, "INDIRECT(#DIV/0!)", EE.DIV_ZERO);
		Confirm(feA, c, "INDIRECT(#NAME?, \"x\")", EE.NAME_INVALID);
		Confirm(feA, c, "INDIRECT(#NUM!, #N/A)", EE.NUM_ERROR);

		// arg1 is Evaluated to bool before arg0 is decoded
		Confirm(feA, c, "INDIRECT(\"garbage\", #N/A)", EE.NA);
		Confirm(feA, c, "INDIRECT(\"garbage\", \"\")", EE.VALUE_INVALID); // empty string is not valid bool
		Confirm(feA, c, "INDIRECT(\"garbage\", \"flase\")", EE.VALUE_INVALID); // must be "TRUE" or "FALSE"


		// spaces around sheet name (with or without quotes Makes no difference)
		Confirm(feA, c, "INDIRECT(\"'Sheet1 '!D3\")", EE.REF_INVALID);
		Confirm(feA, c, "INDIRECT(\" Sheet1!D3\")", EE.REF_INVALID);
		Confirm(feA, c, "INDIRECT(\"'Sheet1' !D3\")", EE.REF_INVALID);


		Confirm(feA, c, "SUM(INDIRECT(\"'John's sales'!A1:C1\"))", EE.REF_INVALID); // bad quote escaping
		Confirm(feA, c, "INDIRECT(\"[Book1]Sheet1!A1\")", EE.REF_INVALID); // unknown external workbook
		Confirm(feA, c, "INDIRECT(\"Sheet3!A1\")", EE.REF_INVALID); // unknown sheet
		if (false) { // TODO - support Evaluation of defined names
			Confirm(feA, c, "INDIRECT(\"Sheet1!IW1\")", EE.REF_INVALID); // bad column
			Confirm(feA, c, "INDIRECT(\"Sheet1!A65537\")", EE.REF_INVALID); // bad row
		}
		Confirm(feA, c, "INDIRECT(\"Sheet1!A 1\")", EE.REF_INVALID); // space in cell ref
	}

	public void TestMultipleWorkbooks() {
		HSSFWorkbook wbA = CreateWBA();
		HSSFCell cellA = wbA.GetSheetAt(0).CreateRow(10).createCell(0);
		HSSFFormulaEvaluator feA = new HSSFFormulaEvaluator(wbA);

		HSSFWorkbook wbB = CreateWBB();
		HSSFCell cellB = wbB.GetSheetAt(0).CreateRow(10).createCell(0);
		HSSFFormulaEvaluator feB = new HSSFFormulaEvaluator(wbB);

		String[] workbookNames = { "MyBook", "Figures for January", };
		HSSFFormulaEvaluator[] Evaluators = { feA, feB, };
		HSSFFormulaEvaluator.SetupEnvironment(workbookNames, Evaluators);

		Confirm(feB, cellB, "INDIRECT(\"'[Figures for January]## Look here!'!A1\")", 42); // same wb
		Confirm(feA, cellA, "INDIRECT(\"'[Figures for January]## Look here!'!A1\")", 42); // across workbooks

		// 2 level recursion
		Confirm(feB, cellB, "INDIRECT(\"[MyBook]Sheet2!A1\")", 50); // Set up (and Check) first level
		Confirm(feA, cellA, "INDIRECT(\"'[Figures for January]Sheet1'!A11\")", 50); // points to cellB
	}

	private static void Confirm(FormulaEvaluator fe, Cell cell, String formula,
			double expectedResult) {
		fe.ClearAllCachedResultValues();
		cell.SetCellFormula(formula);
		CellValue cv = fe.Evaluate(cell);
		if (cv.GetCellType() != Cell.CELL_TYPE_NUMERIC) {
			throw new AssertionFailedError("expected numeric cell type but got " + cv.formatAsString());
		}
		Assert.AreEqual(expectedResult, cv.GetNumberValue(), 0.0);
	}
	private static void Confirm(FormulaEvaluator fe, Cell cell, String formula,
			ErrorEval expectedResult) {
		fe.ClearAllCachedResultValues();
		cell.SetCellFormula(formula);
		CellValue cv = fe.Evaluate(cell);
		if (cv.GetCellType() != Cell.CELL_TYPE_ERROR) {
			throw new AssertionFailedError("expected error cell type but got " + cv.formatAsString());
		}
		int expCode = expectedResult.GetErrorCode();
		if (cv.GetErrorValue() != expCode) {
			throw new AssertionFailedError("Expected error '" + EE.GetText(expCode)
					+ "' but got '" + cv.formatAsString() + "'.");
		}
	}
}

