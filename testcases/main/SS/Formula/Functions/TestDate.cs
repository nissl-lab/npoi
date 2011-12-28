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

using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.Cell;
using NPOI.SS.UserModel.CellValue;

/**
 * @author Pavel Krupets (pkrupets at palmtreebusiness dot com)
 */
public class TestDate  {

    private HSSFCell cell11;
    private HSSFFormulaEvaluator Evaluator;

    public void SetUp() {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet("new sheet");
        cell11 = sheet.CreateRow(0).createCell(0);
        cell11.SetCellType(HSSFCell.CELL_TYPE_FORMULA);
        Evaluator = new HSSFFormulaEvaluator(wb);
    }

    /**
     * Test disabled pending a fix in the formula Evaluator
     * TODO - create MissingArgEval and modify the formula Evaluator to handle this
     */
    public void DISABLEDtestSomeArgumentsMissing() {
        Confirm("DATE(, 1, 0)", 0.0);
        Confirm("DATE(, 1, 1)", 1.0);
    }

    public void TestValid() {

        Confirm("DATE(1900, 1, 1)", 1);
        Confirm("DATE(1900, 1, 32)", 32);
        Confirm("DATE(1900, 222, 1)", 6727);
        Confirm("DATE(1900, 2, 0)", 31);
        Confirm("DATE(2000, 1, 222)", 36747.00);
        Confirm("DATE(2007, 1, 1)", 39083);
    }

    public void TestBugDate() {
        Confirm("DATE(1900, 2, 29)", 60);
        Confirm("DATE(1900, 2, 30)", 61);
        Confirm("DATE(1900, 1, 222)", 222);
        Confirm("DATE(1900, 1, 2222)", 2222);
        Confirm("DATE(1900, 1, 22222)", 22222);
    }

    public void TestPartYears() {
        Confirm("DATE(4, 1, 1)", 1462.00);
        Confirm("DATE(14, 1, 1)", 5115.00);
        Confirm("DATE(104, 1, 1)", 37987.00);
        Confirm("DATE(1004, 1, 1)", 366705.00);
    }

    private void Confirm(String formulaText, double expectedResult) {
        cell11.SetCellFormula(formulaText);
        Evaluator.clearAllCachedResultValues();
        CellValue cv = Evaluator.evaluate(cell11);
        if (cv.GetCellType() != Cell.CELL_TYPE_NUMERIC) {
            throw new AssertionFailedError("Wrong result type: " + cv.formatAsString());
        }
        double actualValue = cv.GetNumberValue();
        Assert.AreEqual(expectedResult, actualValue, 0);
    }
}

