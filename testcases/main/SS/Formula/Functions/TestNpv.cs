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
using NPOI.SS.UserModel.Cell;
using NPOI.SS.UserModel.Row;
using NPOI.SS.UserModel.CellValue;

/**
 * Tests for {@link Npv}
 *
 * @author Marcel May
 * @see <a href="http://office.microsoft.com/en-us/excel-help/npv-HP005209199.aspx">Excel Help</a>
 */
public class TestNpv  {

    public void TestEvaluateInSheetExample2() {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet("Sheet1");
        HSSFRow row = sheet.CreateRow(0);

        sheet.CreateRow(1).createCell(0).SetCellValue(0.08d);
        sheet.CreateRow(2).createCell(0).SetCellValue(-40000d);
        sheet.CreateRow(3).createCell(0).SetCellValue(8000d);
        sheet.CreateRow(4).createCell(0).SetCellValue(9200d);
        sheet.CreateRow(5).createCell(0).SetCellValue(10000d);
        sheet.CreateRow(6).createCell(0).SetCellValue(12000d);
        sheet.CreateRow(7).createCell(0).SetCellValue(14500d);

        HSSFCell cell = row.CreateCell(8);
        HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

        // Enumeration
        cell.SetCellFormula("NPV(A2, A4,A5,A6,A7,A8)+A3");
        fe.ClearAllCachedResultValues();
        fe.EvaluateFormulaCell(cell);
        double res = cell.GetNumericCellValue();
        Assert.AreEqual(1922.06d, Math.round(res * 100d) / 100d);

        // Range
        cell.SetCellFormula("NPV(A2, A4:A8)+A3");

        fe.ClearAllCachedResultValues();
        fe.EvaluateFormulaCell(cell);
        res = cell.GetNumericCellValue();
        Assert.AreEqual(1922.06d, Math.round(res * 100d) / 100d);
    }

    /**
     * Evaluate formulas with NPV and compare the result with
     * the cached formula result pre-calculated by Excel
     */
    public void TestNpvFromSpreadsheet(){
        HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("IrrNpvTestCaseData.xls");
        HSSFSheet sheet = wb.GetSheet("IRR-NPV");
        HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
        StringBuilder failures = new StringBuilder();
        int failureCount = 0;
        // TODO YK: Formulas in rows 16 and 17 operate with ArrayPtg which isn't yet supported
        // FormulaEvaluator as of r1041407 throws "Unexpected ptg class (NPOI.SS.Formula.PTG.ArrayPtg)"
        for(int rownum = 9; rownum <= 15; rownum++){
            HSSFRow row = sheet.GetRow(rownum);
            HSSFCell cellB = row.GetCell(1);
            try {
                CellValue cv = fe.Evaluate(cellB);
                assertFormulaResult(cv, cellB);
            } catch (Throwable e){
                if(failures.Length > 0) failures.Append('\n');
                failures.Append("Row[" + (cellB.RowIndex + 1)+ "]: " + cellB.GetCellFormula() + " ");
                failures.Append(e.GetMessage());
                failureCount++;
            }
        }

        if(failures.Length > 0) {
            throw new AssertionFailedError(failureCount + " IRR Evaluations failed:\n" + failures.ToString());
        }
    }

    private static void assertFormulaResult(CellValue cv, HSSFCell cell){
        double actualValue = cv.GetNumberValue();
        double expectedValue = cell.GetNumericCellValue(); // cached formula result calculated by Excel
        Assert.AreEqual(expectedValue, actualValue, 1E-4); // should agree within 0.01%
    }
}

