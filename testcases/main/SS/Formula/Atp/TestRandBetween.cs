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

using junit.framework.TestCase;

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.SS.UserModel.Cell;
using NPOI.SS.UserModel.FormulaEvaluator;
using NPOI.SS.UserModel.Row;
using NPOI.SS.UserModel.Sheet;
using NPOI.SS.UserModel.Workbook;

/**
 * Testcase for 'Analysis Toolpak' function RANDBETWEEN()
 * 
 * @author Brendan Nolan
 */
public class TestRandBetween  {

	private Workbook wb;
	private FormulaEvaluator Evaluator;
	private Cell bottomValueCell;
	private Cell topValueCell;
	private Cell formulaCell;
	
	
	protected void SetUp() throws Exception {
		super.SetUp();
		wb = HSSFTestDataSamples.OpenSampleWorkbook("TestRandBetween.xls");
		Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
		
		Sheet sheet = wb.CreateSheet("RandBetweenSheet");
		Row row = sheet.CreateRow(0);
		bottomValueCell = row.CreateCell(0);
		topValueCell = row.CreateCell(1);
		formulaCell = row.CreateCell(2, Cell.CELL_TYPE_FORMULA);
	}
	
	
	protected void tearDown() throws Exception {
		// TODO Auto-generated method stub
		super.tearDown();
	}
	
	/**
	 * Check where values are the same
	 */
	public void TestRandBetweenSameValues() {
		
		Evaluator.clearAllCachedResultValues();
		formulaCell.SetCellFormula("RANDBETWEEN(1,1)");
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(1, formulaCell.GetNumericCellValue(), 0);
		Evaluator.clearAllCachedResultValues();
		formulaCell.SetCellFormula("RANDBETWEEN(-1,-1)");
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(-1, formulaCell.GetNumericCellValue(), 0);

	}
	
	/**
	 * Check special case where rounded up bottom value is greater than 
	 * top value.
	 */
	public void TestRandBetweenSpecialCase() {
		

		bottomValueCell.SetCellValue(0.05);		
		topValueCell.SetCellValue(0.1);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(1, formulaCell.GetNumericCellValue(), 0);
		bottomValueCell.SetCellValue(-0.1);		
		topValueCell.SetCellValue(-0.05);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(0, formulaCell.GetNumericCellValue(), 0);
		bottomValueCell.SetCellValue(-1.1);		
		topValueCell.SetCellValue(-1.05);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(-1, formulaCell.GetNumericCellValue(), 0);
		bottomValueCell.SetCellValue(-1.1);		
		topValueCell.SetCellValue(-1.1);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(-1, formulaCell.GetNumericCellValue(), 0);
	}
	
	/**
	 * Check top value of BLANK which Excel will Evaluate as 0
	 */
	public void TestRandBetweenTopBlank() {

		bottomValueCell.SetCellValue(-1);		
		topValueCell.SetCellType(Cell.CELL_TYPE_BLANK);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.IsTrue(formulaCell.GetNumericCellValue() == 0 || formulaCell.GetNumericCellValue() == -1);
	
	}
	/**
	 * Check where input values are of wrong type
	 */
	public void TestRandBetweenWrongInputTypes() {
		// Check case where bottom input is of the wrong type
		bottomValueCell.SetCellValue("STRING");		
		topValueCell.SetCellValue(1);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(Cell.CELL_TYPE_ERROR, formulaCell.GetCachedFormulaResultType());
		Assert.AreEqual(ErrorEval.VALUE_INVALID.GetErrorCode(), formulaCell.GetErrorCellValue());
		
		
		// Check case where top input is of the wrong type
		bottomValueCell.SetCellValue(1);
		topValueCell.SetCellValue("STRING");		
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(Cell.CELL_TYPE_ERROR, formulaCell.GetCachedFormulaResultType());
		Assert.AreEqual(ErrorEval.VALUE_INVALID.GetErrorCode(), formulaCell.GetErrorCellValue());

		// Check case where both inputs are of wrong type
		bottomValueCell.SetCellValue("STRING");
		topValueCell.SetCellValue("STRING");		
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(Cell.CELL_TYPE_ERROR, formulaCell.GetCachedFormulaResultType());
		Assert.AreEqual(ErrorEval.VALUE_INVALID.GetErrorCode(), formulaCell.GetErrorCellValue());
	
	}
	
	/**
	 * Check case where bottom is greater than top
	 */
	public void TestRandBetweenBottomGreaterThanTop() {

		// Check case where bottom is greater than top
		bottomValueCell.SetCellValue(1);		
		topValueCell.SetCellValue(0);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(Cell.CELL_TYPE_ERROR, formulaCell.GetCachedFormulaResultType());
		Assert.AreEqual(ErrorEval.NUM_ERROR.GetErrorCode(), formulaCell.GetErrorCellValue());		
		bottomValueCell.SetCellValue(1);		
		topValueCell.SetCellType(Cell.CELL_TYPE_BLANK);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.AreEqual(Cell.CELL_TYPE_ERROR, formulaCell.GetCachedFormulaResultType());
		Assert.AreEqual(ErrorEval.NUM_ERROR.GetErrorCode(), formulaCell.GetErrorCellValue());
	}
	
	/**
	 * Boundary check of Double MIN and MAX values
	 */
	public void TestRandBetweenBoundaryCheck() {

		bottomValueCell.SetCellValue(Double.MinValue);		
		topValueCell.SetCellValue(Double.MaxValue);
		formulaCell.SetCellFormula("RANDBETWEEN($A$1,$B$1)");
		Evaluator.clearAllCachedResultValues();
		Evaluator.evaluateFormulaCell(formulaCell);
		Assert.IsTrue(formulaCell.GetNumericCellValue() >= Double.MinValue && formulaCell.GetNumericCellValue() <= Double.MaxValue);		
		
	}
	
}

