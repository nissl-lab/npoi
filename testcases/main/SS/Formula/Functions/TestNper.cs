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

using NPOI.SS.Formula.Eval.NumberEval;
using NPOI.SS.Formula.Eval.ValueEval;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFErrorConstants;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;

/**
 * Tests for {@link FinanceFunction#NPER}
 *
 * @author Josh Micich
 */
public class TestNper  {
	public void TestSimpleEvaluate() {

		ValueEval[] args = {
			new NumberEval(0.05),
			new NumberEval(250),
			new NumberEval(-1000),
		};
		ValueEval result = FinanceFunction.NPER.Evaluate(args, 0, (short)0);

		Assert.AreEqual(NumberEval.class, result.GetType());
		Assert.AreEqual(4.57353557, ((NumberEval)result).GetNumberValue(), 0.00000001);
	}

	public void TestEvaluate_bug_45732() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.CreateSheet("Sheet1");
		HSSFCell cell = sheet.CreateRow(0).createCell(0);

		cell.SetCellFormula("NPER(12,4500,100000,100000)");
		cell.SetCellValue(15.0);
		Assert.AreEqual("NPER(12,4500,100000,100000)", cell.GetCellFormula());
		Assert.AreEqual(HSSFCell.CELL_TYPE_NUMERIC, cell.GetCachedFormulaResultType());
		Assert.AreEqual(15.0, cell.GetNumericCellValue(), 0.0);

		HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		fe.EvaluateFormulaCell(cell);
		Assert.AreEqual(HSSFCell.CELL_TYPE_ERROR, cell.GetCachedFormulaResultType());
		Assert.AreEqual(HSSFErrorConstants.ERROR_NUM, cell.GetErrorCellValue());
	}
}

