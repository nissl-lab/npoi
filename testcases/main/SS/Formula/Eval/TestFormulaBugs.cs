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

namespace TestCases.SS.Formula.Eval
{

    using System;
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;

    /**
     * Miscellaneous Tests for bugzilla entries.<p/> The Test name Contains the
     * bugzilla bug id.
     * 
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestFormulaBugs
    {

        /**
         * Bug 27349 - VLOOKUP with reference to another sheet.<p/> This Test was
         * Added <em>long</em> After the relevant functionality was fixed.
         */
        [Test]
        public void Test27349()
        {
            // 27349-vLookupAcrossSheets.xls is bugzilla/attachment.cgi?id=10622
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("27349-vlookupAcrossSheets.xls");
            IWorkbook wb = new HSSFWorkbook(is1);
            ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(1);
            ICell cell = row.GetCell(0);

            // this defInitely would have failed due to 27349
            Assert.AreEqual("VLOOKUP(1,'DATA TABLE'!$A$8:'DATA TABLE'!$B$10,2)", cell
                    .CellFormula);

            // We might as well Evaluate the formula
            IFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);

            Assert.AreEqual(CellType.Numeric, cv.CellType);
            Assert.AreEqual(3.0, cv.NumberValue, 0.0);
        }

        /**
         * Bug 27405 - isnumber() formula always Evaluates to false in if statement<p/>
         * 
         * seems to be a duplicate of 24925
         */
        [Test]
        public void Test27405()
        {

            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("input");
            // input row 0
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell = row.CreateCell(1);
            cell.SetCellValue(1); // B1
            // input row 1
            row = sheet.CreateRow(1);
            cell = row.CreateCell(1);
            cell.SetCellValue(999); // B2

            int rno = 4;
            row = sheet.CreateRow(rno);
            cell = row.CreateCell(1); // B5
            cell.CellFormula = ("isnumber(b1)");
            cell = row.CreateCell(3); // D5
            cell.CellFormula = ("IF(ISNUMBER(b1),b1,b2)");
#if !HIDE_UNREACHABLE_CODE
            if (false)
            { // Set true to check excel file manually
                // bug report mentions 'Editing the formula in excel "fixes" the problem.'
                try
                {
                    FileStream fileOut = new FileStream("27405output.xls", FileMode.Create);
                    wb.Write(fileOut, false);
                    fileOut.Close();
                }
                catch (IOException e)
                {
                    throw new SystemException(e.Message);
                }
            }
#endif
            // use POI's Evaluator as an extra sanity check
            IFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv;
            cv = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Numeric, cv.CellType);
            Assert.AreEqual(1.0, cv.NumberValue, 0.0);

            cv = fe.Evaluate(row.GetCell(1));
            Assert.AreEqual(CellType.Boolean, cv.CellType);
            Assert.AreEqual(true, cv.BooleanValue);
        }

        /**
         * Bug 42448 - Can't parse SUMPRODUCT(A!C7:A!C67, B8:B68) / B69 <p/>
         */
        [Test]
        public void Test42448()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet1 = wb.CreateSheet("Sheet1");

            IRow row = sheet1.CreateRow(0);
            ICell cell = row.CreateCell(0);

            // it's important to create the referenced sheet first
            ISheet sheet2 = wb.CreateSheet("A"); // note name 'A'
            // TODO - POI crashes if the formula is Added before this sheet
            // RuntimeException("Zero length string is an invalid sheet name")
            // Excel doesn't crash but the formula doesn't work until it is
            // re-entered

            String inputFormula = "SUMPRODUCT(A!C7:A!C67, B8:B68) / B69"; // as per bug report
            try
            {
                cell.CellFormula = (inputFormula);
            }
            catch (IndexOutOfRangeException)
            {
                throw new AssertionException("Identified bug 42448");
            }

            Assert.AreEqual("SUMPRODUCT(A!C7:A!C67,B8:B68)/B69", cell.CellFormula);

            // might as well Evaluate the sucker...

            AddCell(sheet2, 5, 2, 3.0); // A!C6
            AddCell(sheet2, 6, 2, 4.0); // A!C7
            AddCell(sheet2, 66, 2, 5.0); // A!C67
            AddCell(sheet2, 67, 2, 6.0); // A!C68

            AddCell(sheet1, 6, 1, 7.0); // B7
            AddCell(sheet1, 7, 1, 8.0); // B8
            AddCell(sheet1, 67, 1, 9.0); // B68
            AddCell(sheet1, 68, 1, 10.0); // B69

            double expectedResult = (4.0 * 8.0 + 5.0 * 9.0) / 10.0;

            IFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);

            Assert.AreEqual(CellType.Numeric, cv.CellType);
            Assert.AreEqual(expectedResult, cv.NumberValue, 0.0);
        }

        private static void AddCell(ISheet sheet, int rowIx, int colIx,
                double value)
        {
            sheet.CreateRow(rowIx).CreateCell(colIx).SetCellValue(value);
        }

        [Test]
        public void Test55032()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("input");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(1);

            checkFormulaValue(wb, cell, "PV(0.08/12, 20*12, 500, ,0)", -59777.14585);
            checkFormulaValue(wb, cell, "PV(0.08/12, 20*12, 500, ,)", -59777.14585);
            checkFormulaValue(wb, cell, "PV(0.08/12, 20*12, 500, 500,)", -59878.6315455);

            checkFormulaValue(wb, cell, "FV(0.08/12, 20*12, 500, ,)", -294510.2078107270);
            checkFormulaValue(wb, cell, "PMT(0.08/12, 20*12, 500, ,)", -4.1822003450);
            checkFormulaValue(wb, cell, "NPER(0.08/12, 20*12, 500, ,)", -2.0758873434);

            wb.Close();
        }

        private void checkFormulaValue(IWorkbook wb, ICell cell, String formula, double expectedValue)
        {
            cell.SetCellFormula(formula);

            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            CellValue value = evaluator.Evaluate(cell);

            Assert.AreEqual(expectedValue, value.NumberValue, 0.0001);
        }
    }
}
