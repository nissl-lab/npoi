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

using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF;
using NPOI.XSSF.Model;

namespace TestCases.XSSF.Model
{

    [TestFixture]
    public class TestCalculationChain
    {
        [Test]
        public void Test46535()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("46535.xlsx");

            CalculationChain chain = wb.GetCalculationChain();
            //the bean holding the reference to the formula to be deleted
            CT_CalcCell c = chain.GetCTCalcChain().GetCArray(0);
            int cnt = chain.GetCTCalcChain().c.Count;
            Assert.AreEqual(10, c.i);
            Assert.AreEqual("E1", c.r);

            ISheet sheet = wb.GetSheet("Test");
            ICell cell = sheet.GetRow(0).GetCell(4);

            Assert.AreEqual(CellType.Formula, cell.CellType);
            cell.SetCellFormula(null);

            //the count of items is less by one
            c = chain.GetCTCalcChain().GetCArray(0);
            int cnt2 = chain.GetCTCalcChain().c.Count;
            Assert.AreEqual(cnt - 1, cnt2);
            //the first item in the calculation chain is the former second one
            Assert.AreEqual(10, c.i);
            Assert.AreEqual("C1", c.r);

            Assert.AreEqual(CellType.String, cell.CellType);
            cell.SetCellValue("ABC");
            Assert.AreEqual(CellType.String, cell.CellType);
        }

        [Test]
        public void RemoveAllFormulas()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("TwoFunctions.xlsx");

            CalculationChain chain = wb.GetCalculationChain();
            //the bean holding the reference to the formula to be deleted
            CT_CalcCell c = chain.GetCTCalcChain().GetCArray(0);
            int cnt = chain.GetCTCalcChain().c.Count;
            Assert.AreEqual(1, c.i);
            Assert.AreEqual("A5", c.r);
            Assert.AreEqual(2, cnt);

            ISheet sheet = wb.GetSheet("Sheet1");
            ICell cell = sheet.GetRow(4).GetCell(0);

            Assert.AreEqual(CellType.Formula, cell.CellType);
            cell.SetCellFormula(null);

            //the count of items is less by one
            c = chain.GetCTCalcChain().GetCArray(0);
            int cnt2 = chain.GetCTCalcChain().c.Count;
            Assert.AreEqual(cnt - 1, cnt2);
            //the first item in the calculation chain is the former second one
            Assert.AreEqual(1, c.i);
            Assert.AreEqual("A4", c.r);
            Assert.AreEqual(1, cnt2);

            //remove final formula from spread sheet
            ICell cell2 = sheet.GetRow(3).GetCell(0);
            Assert.AreEqual(CellType.Formula, cell2.CellType);
            cell2.SetCellFormula(null);

            //the count of items within the chain should be 0
            int cnt3 = chain.GetCTCalcChain().c.Count;
            Assert.AreEqual(0, cnt3);
        }
    }
}