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

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.Formula.OperationEvaluationContext;
using NPOI.SS.Formula.Eval.AreaEval;
using NPOI.SS.Formula.Eval.NumberEval;
using NPOI.SS.Formula.Eval.ValueEval;

using junit.framework.TestCase;
using NPOI.SS.UserModel.*;
using NPOI.SS.Util.SheetBuilder;



/**
 * Tests for {@link Subtotal}
 *
 * @author Paul Tomlin
 */
public class TestSubtotal  {
	private static int FUNCTION_AVERAGE = 1;
	private static int FUNCTION_COUNT = 2;
	private static int FUNCTION_MAX = 4;
	private static int FUNCTION_MIN = 5;
	private static int FUNCTION_PRODUCT = 6;
	private static int FUNCTION_STDEV = 7;
	private static int FUNCTION_SUM = 9;

	private static double[] TEST_VALUES0 = {
		1, 2,
		3, 4,
		5, 6,
		7, 8,
		9, 10
	};

	private static void ConfirmSubtotal(int function, double expected) {
		ValueEval[] values = new ValueEval[TEST_VALUES0.Length];
		for (int i = 0; i < TEST_VALUES0.Length; i++) {
			values[i] = new NumberEval(TEST_VALUES0[i]);
		}

		AreaEval arg1 = EvalFactory.CreateAreaEval("C1:D5", values);
		ValueEval args[] = { new NumberEval(function), arg1 };

		ValueEval result = new Subtotal().Evaluate(args, 0, 0);

		Assert.AreEqual(NumberEval.class, result.GetType());
		Assert.AreEqual(expected, ((NumberEval) result).GetNumberValue(), 0.0);
	}

	public void TestBasics() {
		ConfirmSubtotal(FUNCTION_SUM, 55.0);
		ConfirmSubtotal(FUNCTION_AVERAGE, 5.5);
		ConfirmSubtotal(FUNCTION_COUNT, 10.0);
		ConfirmSubtotal(FUNCTION_MAX, 10.0);
		ConfirmSubtotal(FUNCTION_MIN, 1.0);
		ConfirmSubtotal(FUNCTION_PRODUCT, 3628800.0);
		ConfirmSubtotal(FUNCTION_STDEV, 3.0276503540974917);
	}

     public void TestAvg(){

        Workbook wb = new HSSFWorkbook();

        FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

        Sheet sh = wb.CreateSheet();
        Cell a1 = sh.CreateRow(0).createCell(0);
        a1.SetCellValue(1);
        Cell a2 = sh.CreateRow(1).createCell(0);
        a2.SetCellValue(3);
        Cell a3 = sh.CreateRow(2).createCell(0);
        a3.SetCellFormula("SUBTOTAL(1,A1:A2)");
        Cell a4 = sh.CreateRow(3).createCell(0);
        a4.SetCellValue(1);
        Cell a5 = sh.CreateRow(4).createCell(0);
        a5.SetCellValue(7);
        Cell a6 = sh.CreateRow(5).createCell(0);
        a6.SetCellFormula("SUBTOTAL(1,A1:A5)*2 + 2");
        Cell a7 = sh.CreateRow(6).createCell(0);
        a7.SetCellFormula("SUBTOTAL(1,A1:A6)");

        fe.EvaluateAll();

        Assert.AreEqual(2.0, a3.GetNumericCellValue());
        Assert.AreEqual(8.0, a6.GetNumericCellValue());
        Assert.AreEqual(3.0, a7.GetNumericCellValue());
    }

    public void TestSum(){

        Workbook wb = new HSSFWorkbook();

        FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

        Sheet sh = wb.CreateSheet();
        Cell a1 = sh.CreateRow(0).createCell(0);
        a1.SetCellValue(1);
        Cell a2 = sh.CreateRow(1).createCell(0);
        a2.SetCellValue(3);
        Cell a3 = sh.CreateRow(2).createCell(0);
        a3.SetCellFormula("SUBTOTAL(9,A1:A2)");
        Cell a4 = sh.CreateRow(3).createCell(0);
        a4.SetCellValue(1);
        Cell a5 = sh.CreateRow(4).createCell(0);
        a5.SetCellValue(7);
        Cell a6 = sh.CreateRow(5).createCell(0);
        a6.SetCellFormula("SUBTOTAL(9,A1:A5)*2 + 2");
        Cell a7 = sh.CreateRow(6).createCell(0);
        a7.SetCellFormula("SUBTOTAL(9,A1:A6)");

        fe.EvaluateAll();

        Assert.AreEqual(4.0, a3.GetNumericCellValue());
        Assert.AreEqual(26.0, a6.GetNumericCellValue());
        Assert.AreEqual(12.0, a7.GetNumericCellValue());
    }

    public void TestCount(){

        Workbook wb = new HSSFWorkbook();

        FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

        Sheet sh = wb.CreateSheet();
        Cell a1 = sh.CreateRow(0).createCell(0);
        a1.SetCellValue(1);
        Cell a2 = sh.CreateRow(1).createCell(0);
        a2.SetCellValue(3);
        Cell a3 = sh.CreateRow(2).createCell(0);
        a3.SetCellFormula("SUBTOTAL(2,A1:A2)");
        Cell a4 = sh.CreateRow(3).createCell(0);
        a4.SetCellValue("POI");                  // A4 is string and not counted
        Cell a5 = sh.CreateRow(4).createCell(0); // A5 is blank and not counted

        Cell a6 = sh.CreateRow(5).createCell(0);
        a6.SetCellFormula("SUBTOTAL(2,A1:A5)*2 + 2");
        Cell a7 = sh.CreateRow(6).createCell(0);
        a7.SetCellFormula("SUBTOTAL(2,A1:A6)");

        fe.EvaluateAll();

        Assert.AreEqual(2.0, a3.GetNumericCellValue());
        Assert.AreEqual(6.0, a6.GetNumericCellValue());
        Assert.AreEqual(2.0, a7.GetNumericCellValue());
    }

    public void TestCounta(){

        Workbook wb = new HSSFWorkbook();

        FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

        Sheet sh = wb.CreateSheet();
        Cell a1 = sh.CreateRow(0).createCell(0);
        a1.SetCellValue(1);
        Cell a2 = sh.CreateRow(1).createCell(0);
        a2.SetCellValue(3);
        Cell a3 = sh.CreateRow(2).createCell(0);
        a3.SetCellFormula("SUBTOTAL(3,A1:A2)");
        Cell a4 = sh.CreateRow(3).createCell(0);
        a4.SetCellValue("POI");                  // A4 is string and not counted
        Cell a5 = sh.CreateRow(4).createCell(0); // A5 is blank and not counted

        Cell a6 = sh.CreateRow(5).createCell(0);
        a6.SetCellFormula("SUBTOTAL(3,A1:A5)*2 + 2");
        Cell a7 = sh.CreateRow(6).createCell(0);
        a7.SetCellFormula("SUBTOTAL(3,A1:A6)");

        fe.EvaluateAll();

        Assert.AreEqual(2.0, a3.GetNumericCellValue());
        Assert.AreEqual(8.0, a6.GetNumericCellValue());
        Assert.AreEqual(3.0, a7.GetNumericCellValue());
    }

    public void TestMax(){

        Workbook wb = new HSSFWorkbook();

        FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

        Sheet sh = wb.CreateSheet();
        Cell a1 = sh.CreateRow(0).createCell(0);
        a1.SetCellValue(1);
        Cell a2 = sh.CreateRow(1).createCell(0);
        a2.SetCellValue(3);
        Cell a3 = sh.CreateRow(2).createCell(0);
        a3.SetCellFormula("SUBTOTAL(4,A1:A2)");
        Cell a4 = sh.CreateRow(3).createCell(0);
        a4.SetCellValue(1);
        Cell a5 = sh.CreateRow(4).createCell(0);
        a5.SetCellValue(7);
        Cell a6 = sh.CreateRow(5).createCell(0);
        a6.SetCellFormula("SUBTOTAL(4,A1:A5)*2 + 2");
        Cell a7 = sh.CreateRow(6).createCell(0);
        a7.SetCellFormula("SUBTOTAL(4,A1:A6)");

        fe.EvaluateAll();

        Assert.AreEqual(3.0, a3.GetNumericCellValue());
        Assert.AreEqual(16.0, a6.GetNumericCellValue());
        Assert.AreEqual(7.0, a7.GetNumericCellValue());
    }

    public void TestMin(){

        Workbook wb = new HSSFWorkbook();

        FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

        Sheet sh = wb.CreateSheet();
        Cell a1 = sh.CreateRow(0).createCell(0);
        a1.SetCellValue(1);
        Cell a2 = sh.CreateRow(1).createCell(0);
        a2.SetCellValue(3);
        Cell a3 = sh.CreateRow(2).createCell(0);
        a3.SetCellFormula("SUBTOTAL(5,A1:A2)");
        Cell a4 = sh.CreateRow(3).createCell(0);
        a4.SetCellValue(1);
        Cell a5 = sh.CreateRow(4).createCell(0);
        a5.SetCellValue(7);
        Cell a6 = sh.CreateRow(5).createCell(0);
        a6.SetCellFormula("SUBTOTAL(5,A1:A5)*2 + 2");
        Cell a7 = sh.CreateRow(6).createCell(0);
        a7.SetCellFormula("SUBTOTAL(5,A1:A6)");

        fe.EvaluateAll();

        Assert.AreEqual(1.0, a3.GetNumericCellValue());
        Assert.AreEqual(4.0, a6.GetNumericCellValue());
        Assert.AreEqual(1.0, a7.GetNumericCellValue());
    }

    public void TestStdev(){

        Workbook wb = new HSSFWorkbook();

        FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

        Sheet sh = wb.CreateSheet();
        Cell a1 = sh.CreateRow(0).createCell(0);
        a1.SetCellValue(1);
        Cell a2 = sh.CreateRow(1).createCell(0);
        a2.SetCellValue(3);
        Cell a3 = sh.CreateRow(2).createCell(0);
        a3.SetCellFormula("SUBTOTAL(7,A1:A2)");
        Cell a4 = sh.CreateRow(3).createCell(0);
        a4.SetCellValue(1);
        Cell a5 = sh.CreateRow(4).createCell(0);
        a5.SetCellValue(7);
        Cell a6 = sh.CreateRow(5).createCell(0);
        a6.SetCellFormula("SUBTOTAL(7,A1:A5)*2 + 2");
        Cell a7 = sh.CreateRow(6).createCell(0);
        a7.SetCellFormula("SUBTOTAL(7,A1:A6)");

        fe.EvaluateAll();

        Assert.AreEqual(1.41421, a3.GetNumericCellValue(), 0.0001);
        Assert.AreEqual(7.65685, a6.GetNumericCellValue(), 0.0001);
        Assert.AreEqual(2.82842, a7.GetNumericCellValue(), 0.0001);
    }

    public void Test50209(){
        Workbook wb = new HSSFWorkbook();
        Sheet sh = wb.CreateSheet();
        Cell a1 = sh.CreateRow(0).createCell(0);
        a1.SetCellValue(1);
        Cell a2 = sh.CreateRow(1).createCell(0);
        a2.SetCellFormula("SUBTOTAL(9,A1)");
        Cell a3 = sh.CreateRow(2).createCell(0);
        a3.SetCellFormula("SUBTOTAL(9,A1:A2)");

        FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
        fe.EvaluateAll();
        Assert.AreEqual(1.0, a2.GetNumericCellValue());
        Assert.AreEqual(1.0, a3.GetNumericCellValue());
    }

    private static void ConfirmExpectedResult(FormulaEvaluator Evaluator, String msg, Cell cell, double expected) {

        CellValue value = Evaluator.evaluate(cell);
        if (value.GetErrorValue() != 0)
            throw new RuntimeException(msg + ": " + value.formatAsString());
        Assert.AreEqual(msg, expected, value.GetNumberValue());
    }

    public void TestFunctionsFromTestSpreadsheet() {
        HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("SubtotalsNested.xls");
        HSSFSheet sheet = workbook.GetSheetAt(0);
        FormulaEvaluator Evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

        Assert.AreEqual("A1", 10.0, sheet.GetRow(0).GetCell(0).GetNumericCellValue());
        Assert.AreEqual("A2", 20.0, sheet.GetRow(1).GetCell(0).GetNumericCellValue());

        //Test simple subtotal over one area
        Cell cellA3 = sheet.GetRow(2).GetCell(0);
        ConfirmExpectedResult(evaluator, "A3", cellA3, 30.0);

        //Test existence of the second area
        assertNotNull("B1 must not be null", sheet.GetRow(0).GetCell(1));
        Assert.AreEqual("B1", 7.0, sheet.GetRow(0).GetCell(1).GetNumericCellValue());

        Cell cellC1 = sheet.GetRow(0).GetCell(2);
        Cell cellC2 = sheet.GetRow(1).GetCell(2);
        Cell cellC3 = sheet.GetRow(2).GetCell(2);

        //Test Functions SUM, COUNT and COUNTA calculation of SUBTOTAL
        //a) areas A and B are used
        //b) first 2 subtotals don't consider the value of nested subtotal in A3
        ConfirmExpectedResult(evaluator, "SUBTOTAL(SUM;A1:A7;B1:B7)", cellC1, 37.0);
        ConfirmExpectedResult(evaluator, "SUBTOTAL(COUNT;A1:A7;B1:B7)", cellC2, 3.0);
        ConfirmExpectedResult(evaluator, "SUBTOTAL(COUNTA;A1:A7;B1:B7)", cellC3, 5.0);
    }
}

