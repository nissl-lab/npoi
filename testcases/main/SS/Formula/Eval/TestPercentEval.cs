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

using NPOI.SS.Formula.functions.EvalFactory;
using NPOI.SS.Formula.functions.NumericFunctionInvoker;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFRow;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.CellValue;

/**
 * Test for percent operator Evaluator.
 *
 * @author Josh Micich
 */
public class TestPercentEval  {

	private static void Confirm(ValueEval arg, double expectedResult) {
		ValueEval[] args = {
			arg,
		};

		double result = NumericFunctionInvoker.invoke(PercentEval.instance, args, 0, 0);

		Assert.AreEqual(expectedResult, result, 0);
	}

	public void TestBasic() {
		Confirm(new NumberEval(5), 0.05);
		Confirm(new NumberEval(3000), 30.0);
		Confirm(new NumberEval(-150), -1.5);
		Confirm(new StringEval("0.2"), 0.002);
		Confirm(BoolEval.TRUE, 0.01);
	}

	public void Test1x1Area() {
		AreaEval ae = EvalFactory.CreateAreaEval("B2:B2", new ValueEval[] { new NumberEval(50), });
		Confirm(ae, 0.5);
	}
	public void TestInSpreadSheet() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.CreateSheet("Sheet1");
		HSSFRow row = sheet.CreateRow(0);
		HSSFCell cell = row.CreateCell(0);
		cell.SetCellFormula("B1%");
		row.CreateCell(1).SetCellValue(50.0);

		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		CellValue cv;
		try {
			cv = fe.Evaluate(cell);
		} catch (RuntimeException e) {
			if(e.GetCause() is NullPointerException) {
				throw new AssertionFailedError("Identified bug 44608");
			}
			// else some other unexpected error
			throw e;
		}
		Assert.AreEqual(HSSFCell.CELL_TYPE_NUMERIC, cv.GetCellType());
		Assert.AreEqual(0.5, cv.GetNumberValue(), 0.0);
	}
}

