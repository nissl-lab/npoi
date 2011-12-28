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
using junit.framework.AssertionFailedError;
using NPOI.hssf.UserModel.*;
using NPOI.hssf.HSSFTestDataSamples;
using NPOI.SS.UserModel.CellValue;

/**
 * Tests for {@link Irr}
 *
 * @author Marcel May
 */
public class TestIrr  {

    public void TestIrr() {
        // http://en.wikipedia.org/wiki/Internal_rate_of_return#Example
        double[] incomes = {-4000d, 1200d, 1410d, 1875d, 1050d};
        double irr = Irr.irr(incomes);
        double irrRounded = Math.round(irr * 1000d) / 1000d;
        Assert.AreEqual("irr", 0.143d, irrRounded);

        // http://www.techonthenet.com/excel/formulas/irr.php
        incomes = new double[]{-7500d, 3000d, 5000d, 1200d, 4000d};
        irr = Irr.irr(incomes);
        irrRounded = Math.round(irr * 100d) / 100d;
        Assert.AreEqual("irr", 0.28d, irrRounded);

        incomes = new double[]{-10000d, 3400d, 6500d, 1000d};
        irr = Irr.irr(incomes);
        irrRounded = Math.round(irr * 100d) / 100d;
        Assert.AreEqual("irr", 0.05, irrRounded);

        incomes = new double[]{100d, -10d, -110d};
        irr = Irr.irr(incomes);
        irrRounded = Math.round(irr * 100d) / 100d;
        Assert.AreEqual("irr", 0.1, irrRounded);

        incomes = new double[]{-70000d, 12000, 15000};
        irr = Irr.irr(incomes, -0.1);
        irrRounded = Math.round(irr * 100d) / 100d;
        Assert.AreEqual("irr", -0.44, irrRounded);
    }

    public void TestEvaluateInSheet() {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet("Sheet1");
        HSSFRow row = sheet.CreateRow(0);

        row.CreateCell(0).SetCellValue(-4000d);
        row.CreateCell(1).SetCellValue(1200d);
        row.CreateCell(2).SetCellValue(1410d);
        row.CreateCell(3).SetCellValue(1875d);
        row.CreateCell(4).SetCellValue(1050d);

        HSSFCell cell = row.CreateCell(5);
        cell.SetCellFormula("IRR(A1:E1)");

        HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
        fe.ClearAllCachedResultValues();
        fe.EvaluateFormulaCell(cell);
        double res = cell.GetNumericCellValue();
        Assert.AreEqual(0.143d, Math.round(res * 1000d) / 1000d);
    }

    public void TestIrrFromSpreadsheet(){
        HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("IrrNpvTestCaseData.xls");
        HSSFSheet sheet = wb.GetSheet("IRR-NPV");
        HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
        StringBuilder failures = new StringBuilder();
        int failureCount = 0;
        // TODO YK: Formulas in rows 16 and 17 operate with ArrayPtg which isn't yet supported
        // FormulaEvaluator as of r1041407 throws "Unexpected ptg class (NPOI.SS.Formula.PTG.ArrayPtg)"
        for(int rownum = 9; rownum <= 15; rownum++){
            HSSFRow row = sheet.GetRow(rownum);
            HSSFCell cellA = row.GetCell(0);
            try {
                CellValue cv = fe.Evaluate(cellA);
                assertFormulaResult(cv, cellA);
            } catch (Throwable e){
                if(failures.Length > 0) failures.Append('\n');
                failures.Append("Row[" + (cellA.RowIndex + 1)+ "]: " + cellA.GetCellFormula() + " ");
                failures.Append(e.GetMessage());
                failureCount++;
            }

            HSSFCell cellC = row.GetCell(2); //IRR-NPV relationship: NPV(IRR(values), values) = 0
            try {
                CellValue cv = fe.Evaluate(cellC);
                Assert.AreEqual(0, cv.GetNumberValue(), 0.0001);  // should agree within 0.01%
            } catch (Throwable e){
                if(failures.Length > 0) failures.Append('\n');
                failures.Append("Row[" + (cellC.RowIndex + 1)+ "]: " + cellC.GetCellFormula() + " ");
                failures.Append(e.GetMessage());
                failureCount++;
            }
        }

        if(failures.Length > 0) {
            throw new AssertionFailedError(failureCount + " IRR assertions failed:\n" + failures.ToString());
        }

    }

    private static void assertFormulaResult(CellValue cv, HSSFCell cell){
        double actualValue = cv.GetNumberValue();
        double expectedValue = cell.GetNumericCellValue(); // cached formula result calculated by Excel
        Assert.AreEqual("Invalid formula result: " + cv.ToString(), HSSFCell.CELL_TYPE_NUMERIC, cv.GetCellType());
        Assert.AreEqual(expectedValue, actualValue, 1E-4); // should agree within 0.01%
    }
}

