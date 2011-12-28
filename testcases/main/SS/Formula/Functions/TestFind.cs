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

using junit.framework.TestCase;

using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFErrorConstants;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.CellValue;

/**
 * Tests for {@link Financed}
 * 
 * @author Torstein Svendsen (torstei@officenet.no)
 */
public class TestFind  {

	public void TestFind() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFCell cell = wb.CreateSheet().createRow(0).createCell(0);

		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

		ConfirmResult(fe, cell, "Find(\"h\", \"haystack\")", 1);
		ConfirmResult(fe, cell, "Find(\"a\", \"haystack\",2)", 2);
		ConfirmResult(fe, cell, "Find(\"a\", \"haystack\",3)", 6);

		// number args Converted to text
		ConfirmResult(fe, cell, "Find(7, 32768)", 3);
		ConfirmResult(fe, cell, "Find(\"34\", 1341235233412, 3)", 10);
		ConfirmResult(fe, cell, "Find(5, 87654)", 4);

		// Errors
		ConfirmError(fe, cell, "Find(\"n\", \"haystack\")", HSSFErrorConstants.ERROR_VALUE);
		ConfirmError(fe, cell, "Find(\"k\", \"haystack\",9)", HSSFErrorConstants.ERROR_VALUE);
		ConfirmError(fe, cell, "Find(\"k\", \"haystack\",#REF!)", HSSFErrorConstants.ERROR_REF);
		ConfirmError(fe, cell, "Find(\"k\", \"haystack\",0)", HSSFErrorConstants.ERROR_VALUE);
		ConfirmError(fe, cell, "Find(#DIV/0!, #N/A, #REF!)", HSSFErrorConstants.ERROR_DIV_0);
		ConfirmError(fe, cell, "Find(2, #N/A, #REF!)", HSSFErrorConstants.ERROR_NA);
	}

	private static void ConfirmResult(HSSFFormulaEvaluator fe, HSSFCell cell, String formulaText,
			int expectedResult) {
		cell.SetCellFormula(formulaText);
		fe.NotifyUpdateCell(cell);
		CellValue result = fe.Evaluate(cell);
		Assert.AreEqual(result.GetCellType(), HSSFCell.CELL_TYPE_NUMERIC);
		Assert.AreEqual(expectedResult, result.GetNumberValue(), 0.0);
	}

	private static void ConfirmError(HSSFFormulaEvaluator fe, HSSFCell cell, String formulaText,
			int expectedErrorCode) {
		cell.SetCellFormula(formulaText);
		fe.NotifyUpdateCell(cell);
		CellValue result = fe.Evaluate(cell);
		Assert.AreEqual(result.GetCellType(), HSSFCell.CELL_TYPE_ERROR);
		Assert.AreEqual(expectedErrorCode, result.GetErrorValue());
	}
}

