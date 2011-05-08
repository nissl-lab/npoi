/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record.Formula.Functions
{
    using System;
    using System.Collections;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
/**
 * Tests for Excel function ISBLANK()
 * 
 * @author Josh Micich
 */
    [TestClass]
    public class TestIsBlank
    {
        [TestMethod]
        public void Test3DArea()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet sheet1 = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");
            wb.CreateSheet();
            wb.SetSheetName(1, "Sheet2");
            Row row = sheet1.CreateRow(0);
            Cell cell = row.CreateCell((short)0);


            cell.CellFormula = ("isblank(Sheet2!A1:A1)");

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(sheet1, wb);
            //fe.SetCurrentRow(row);
            NPOI.SS.UserModel.CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BOOLEAN, result.CellType);
            Assert.AreEqual(true, result.BooleanValue);

            cell.CellFormula = ("isblank(D7:D7)");

            result = fe.Evaluate(cell);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BOOLEAN, result.CellType);
            Assert.AreEqual(true, result.BooleanValue);

        }
    }
}