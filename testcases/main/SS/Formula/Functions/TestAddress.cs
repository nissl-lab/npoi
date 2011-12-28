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

using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.UserModel.CellValue;

using junit.framework.TestCase;
using NPOI.SS.Util.CellReference;

public class TestAddress  {

    public void TestAddress() {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFCell cell = wb.CreateSheet().createRow(0).createCell(0);
        HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

        String formulaText = "ADDRESS(1,2)";
        ConfirmResult(fe, cell, formulaText, "$B$1");

        formulaText = "ADDRESS(22,44)";
        ConfirmResult(fe, cell, formulaText, "$AR$22");

        formulaText = "ADDRESS(1,1)";
        ConfirmResult(fe, cell, formulaText, "$A$1");

        formulaText = "ADDRESS(1,128)";
        ConfirmResult(fe, cell, formulaText, "$DX$1");

        formulaText = "ADDRESS(1,512)";
        ConfirmResult(fe, cell, formulaText, "$SR$1");

        formulaText = "ADDRESS(1,1000)";
        ConfirmResult(fe, cell, formulaText, "$ALL$1");

        formulaText = "ADDRESS(1,10000)";
        ConfirmResult(fe, cell, formulaText, "$NTP$1");

        formulaText = "ADDRESS(2,3)";
        ConfirmResult(fe, cell, formulaText, "$C$2");

        formulaText = "ADDRESS(2,3,2)";
        ConfirmResult(fe, cell, formulaText, "C$2");

        formulaText = "ADDRESS(2,3,2,,\"EXCEL SHEET\")";
        ConfirmResult(fe, cell, formulaText, "'EXCEL SHEET'!C$2");

        formulaText = "ADDRESS(2,3,3,TRUE,\"[Book1]Sheet1\")";
        ConfirmResult(fe, cell, formulaText, "'[Book1]Sheet1'!$C2");
    }

    private static void ConfirmResult(HSSFFormulaEvaluator fe, HSSFCell cell, String formulaText,
                                      String expectedResult) {
        cell.SetCellFormula(formulaText);
        fe.NotifyUpdateCell(cell);
        CellValue result = fe.Evaluate(cell);
        Assert.AreEqual(result.GetCellType(), HSSFCell.CELL_TYPE_STRING);
        Assert.AreEqual(expectedResult, result.StringValue);
    }
}

