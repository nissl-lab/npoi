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
 * Test for YEAR / MONTH / DAY / HOUR / MINUTE / SECOND
 */
public class TestCalendarFieldFunction  {

    private HSSFCell cell11;
    private HSSFFormulaEvaluator Evaluator;

    public void SetUp() {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet("new sheet");
        cell11 = sheet.CreateRow(0).createCell(0);
        cell11.SetCellType(HSSFCell.CELL_TYPE_FORMULA);
        Evaluator = new HSSFFormulaEvaluator(wb);
    }

    public void TestValid() {
        Confirm("YEAR(2.26)", 1900);
        Confirm("MONTH(2.26)", 1);
        Confirm("DAY(2.26)", 2);
        Confirm("HOUR(2.26)", 6);
        Confirm("MINUTE(2.26)", 14);
        Confirm("SECOND(2.26)", 24);
        
        Confirm("YEAR(40627.4860417)", 2011);
        Confirm("MONTH(40627.4860417)", 3);
        Confirm("DAY(40627.4860417)", 25);
        Confirm("HOUR(40627.4860417)", 11);
        Confirm("MINUTE(40627.4860417)", 39);
        Confirm("SECOND(40627.4860417)", 54);
    }

    public void TestBugDate() {
        Confirm("YEAR(0.0)", 1900);
        Confirm("MONTH(0.0)", 1);
        Confirm("DAY(0.0)", 0);
        
        Confirm("YEAR(0.26)", 1900);
        Confirm("MONTH(0.26)", 1);
        Confirm("DAY(0.26)", 0);
        Confirm("HOUR(0.26)", 6);
        Confirm("MINUTE(0.26)", 14);
        Confirm("SECOND(0.26)", 24);
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

