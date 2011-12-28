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
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.CellValue;

/**
 * Tests for {@link MissingArgEval}
 *
 * @author Josh Micich
 */
public class TestMissingArgEval  {
	
	public void TestEvaluateMissingArgs() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		HSSFSheet sheet = wb.CreateSheet("Sheet1");
		HSSFCell cell = sheet.CreateRow(0).createCell(0);
		
		cell.SetCellFormula("if(true,)"); 
		fe.ClearAllCachedResultValues();
		CellValue cv;
		try {
			cv = fe.Evaluate(cell);
		} catch (EmptyStackException e) {
			throw new AssertionFailedError("Missing args Evaluation not implemented (bug 43354");
		}
		// MissingArg -> BlankEval -> zero (as formula result)
		Assert.AreEqual(0.0, cv.GetNumberValue(), 0.0);
		
		// MissingArg -> BlankEval -> empty string (in concatenation)
		cell.SetCellFormula("\"abc\"&if(true,)"); 
		fe.ClearAllCachedResultValues();
		Assert.AreEqual("abc", fe.Evaluate(cell).StringValue);
	}
	
	public void TestCountFuncs() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		HSSFSheet sheet = wb.CreateSheet("Sheet1");
		HSSFCell cell = sheet.CreateRow(0).createCell(0);
		
		cell.SetCellFormula("COUNT(C5,,,,)"); // 4 missing args, C5 is blank 
		Assert.AreEqual(4.0, fe.Evaluate(cell).GetNumberValue(), 0.0);

		cell.SetCellFormula("COUNTA(C5,,)"); // 2 missing args, C5 is blank 
		fe.ClearAllCachedResultValues();
		Assert.AreEqual(2.0, fe.Evaluate(cell).GetNumberValue(), 0.0);
	}
}

