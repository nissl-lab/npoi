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

namespace TestCases.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    /**
     * Tests for Excel function ISBLANK()
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestIsBlank
    {
        [Test]
        public void Test3DArea()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");
            wb.CreateSheet();
            wb.SetSheetName(1, "Sheet2");
            IRow row = sheet1.CreateRow(0);
            ICell cell = row.CreateCell(0);


            cell.CellFormula=("isblank(Sheet2!A1:A1)");

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Boolean, result.CellType);
            Assert.AreEqual(true, result.BooleanValue);

            cell.CellFormula=("isblank(D7:D7)");

            result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Boolean, result.CellType);
            Assert.AreEqual(true, result.BooleanValue);
        }
    }

}