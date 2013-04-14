/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using NPOI.SS.Util;
using NUnit.Framework;
using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
namespace TestCases.SS.Util
{

    /**
     * Tests SheetBuilder.
     * @see org.apache.poi.ss.util.SheetBuilder
     */
    [TestFixture]
    public class TestSheetBuilder
    {

        private static Object[][] testData = new Object[][] {
	new object[]{         1,     2,        3},
	new object[]{DateTime.Now,  null,     null},
	new object[]{     "one", "two", "=A1+B2"}
    };
        [Test]
        public void TestNotCreateEmptyCells()
        {
            IWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = new SheetBuilder(wb, testData).Build();

            Assert.AreEqual(sheet.PhysicalNumberOfRows, 3);

            NPOI.SS.UserModel.IRow firstRow = sheet.GetRow(0);
            NPOI.SS.UserModel.ICell firstCell = firstRow.GetCell(0);

            Assert.AreEqual(firstCell.CellType, CellType.Numeric);
            Assert.AreEqual(1.0, firstCell.NumericCellValue, 0.00001);


            NPOI.SS.UserModel.IRow secondRow = sheet.GetRow(1);
            Assert.IsNotNull(secondRow.GetCell(0));
            Assert.IsNull(secondRow.GetCell(2));

            NPOI.SS.UserModel.IRow thirdRow = sheet.GetRow(2);
            Assert.AreEqual(CellType.String, thirdRow.GetCell(0).CellType);
            String cellValue = thirdRow.GetCell(0).StringCellValue;
            Assert.AreEqual(testData[2][0].ToString(), cellValue);

            Assert.AreEqual(CellType.Formula, thirdRow.GetCell(2).CellType);
            Assert.AreEqual("A1+B2", thirdRow.GetCell(2).CellFormula);
        }
        [Test]
        public void TestEmptyCells()
        {
            NPOI.SS.UserModel.IWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = new SheetBuilder(wb, testData).SetCreateEmptyCells(true).Build();

            NPOI.SS.UserModel.ICell emptyCell = sheet.GetRow(1).GetCell(1);
            Assert.IsNotNull(emptyCell);
            Assert.AreEqual(CellType.Blank, emptyCell.CellType);
        }

        [Test]
        public void TestSheetName()
        {
            String sheetName = "TEST SHEET NAME";
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, testData).SetSheetName(sheetName).Build();
            Assert.AreEqual(sheetName, sheet.SheetName);
        }
    }
}