/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace TestCases.SS.Formula.Eval
{
    using System;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Miscellaneous tests for bugzilla entries.<p/> The test name contains the
     * bugzilla bug id.
     * 
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestFormulaBugs
    {

        /**
         * Bug 27349 - VLOOKUP with reference to another sheet.<p/> This test was
         * Added <em>long</em> after the relevant functionality was fixed.
         */
        [TestMethod]
        public void Test27349()
        {
            // 27349-vlookupAcrossSheets.xls is1 bugzilla/attachment.cgi?id=10622
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("27349-vlookupAcrossSheets.xls");
            HSSFWorkbook wb;
            try
            {
                // original bug may have thrown exception here, or output warning to
                // stderr
                wb = new HSSFWorkbook(is1);
            }
            catch (IOException)
            {
                throw;
            }

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(1);
            ICell cell = row.GetCell(0);

            // this definitely would have failed due to 27349
            Assert.AreEqual("VLOOKUP(1,'DATA TABLE'!$A$8:'DATA TABLE'!$B$10,2)", cell.CellFormula);

            // We might as well evaluate the formula
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(sheet, wb);
            //fe.SetCurrentRow(row);
            NPOI.SS.UserModel.CellValue cv = fe.Evaluate(cell);

            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(3.0, cv.NumberValue, 0.0);
        }

        /**
         * Bug 27405 - isnumber() formula always evaluates to false in if statement<p/>
         * 
         * seems to be a duplicate of 24925
         */
        [TestMethod]
        public void Test27405()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("input");
            // input row 0
            IRow row = sheet.CreateRow((short)0);
            ICell cell = row.CreateCell((short)0);
            cell = row.CreateCell((short)1);
            cell.SetCellValue(1); // B1
            // input row 1
            row = sheet.CreateRow((short)1);
            cell = row.CreateCell((short)1);
            cell.SetCellValue(999); // B2

            int rno = 4;
            row = sheet.CreateRow(rno);
            cell = row.CreateCell((short)1); // B5
            cell.CellFormula = ("isnumber(b1)");
            cell = row.CreateCell((short)3); // D5
            cell.CellFormula = ("IF(ISNUMBER(b1),b1,b2)");

            //if (false)
            //{ // Set true to check excel file manually
            //    // bug report mentions 'Editing the formula in excel "fixes" the problem.'
            //    try
            //    {
            //        FileStream fileOut = new FileStream("27405output.xls",FileMode.Open);
            //        wb.Write(fileOut);
            //        fileOut.Close();
            //    }
            //    catch (IOException)
            //    {
            //        throw;
            //    }
            //}

            // use POI's evaluator as an extra sanity check
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(sheet, wb);
            //fe.SetCurrentRow(row);
            NPOI.SS.UserModel.CellValue cv;
            cv = fe.Evaluate(cell);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(1.0, cv.NumberValue, 0.0);

            cv = fe.Evaluate(row.GetCell(1));
            Assert.AreEqual(NPOI.SS.UserModel.CellType.BOOLEAN, cv.CellType);
            Assert.AreEqual(true, cv.BooleanValue);
        }
        /**
         * Bug 21334: "File error: data may have been lost" with a file
         * that contains macros and this formula:
         * {=SUM(IF(FREQUENCY(IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),""),IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),""))>0,1))}
         */
        [TestMethod]
        public void Test21334()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sh = wb.CreateSheet();
            ICell cell = sh.CreateRow(0).CreateCell(0);
            String formula = "SUM(IF(FREQUENCY(IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),\"\"),IF(LEN(V4:V220)>0,MATCH(V4:V220,V4:V220,0),\"\"))>0,1))";
            cell.CellFormula = (formula);

            HSSFWorkbook wb_sv = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            ICell cell_sv = wb_sv.GetSheetAt(0).GetRow(0).GetCell(0);
            Assert.AreEqual(formula, cell_sv.CellFormula);
        }
        /**
         * Bug 42448 - Can't parse SUMPRODUCT(A!C7:A!C67, B8:B68) / B69 <p/>
         */
        [TestMethod]
        public void Test42448()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet1 = wb.CreateSheet("Sheet1");

            IRow row = sheet1.CreateRow(0);
            ICell cell = row.CreateCell((short)0);

            // it's important to create the referenced sheet first
            NPOI.SS.UserModel.ISheet sheet2 = wb.CreateSheet("A"); // note name 'A'
            // TODO - POI crashes if the formula is added before this sheet
            // Exception("Zero length string is an invalid sheet name")
            // Excel doesn't crash but the formula doesn't work until it is
            // re-entered

            String inputFormula = "SUMPRODUCT(A!C7:A!C67, B8:B68) / B69"; // as per bug report
            try
            {
                cell.CellFormula = (inputFormula);
            }
            catch (IndexOutOfRangeException)
            {
                throw new AssertFailedException("Identified bug 42448");
            }

            Assert.AreEqual("SUMPRODUCT(A!C7:A!C67,B8:B68)/B69", cell.CellFormula);

            // might as well evaluate the sucker...

            AddCell(sheet2, 5, 2, 3.0); // A!C6
            AddCell(sheet2, 6, 2, 4.0); // A!C7
            AddCell(sheet2, 66, 2, 5.0); // A!C67
            AddCell(sheet2, 67, 2, 6.0); // A!C68

            AddCell(sheet1, 6, 1, 7.0); // B7
            AddCell(sheet1, 7, 1, 8.0); // B8
            AddCell(sheet1, 67, 1, 9.0); // B68
            AddCell(sheet1, 68, 1, 10.0); // B69

            double expectedResult = (4.0 * 8.0 + 5.0 * 9.0) / 10.0;

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(sheet1, wb);
            //fe.SetCurrentRow(row);
            NPOI.SS.UserModel.CellValue cv = fe.Evaluate(cell);

            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(expectedResult, cv.NumberValue, 0.0);
        }

        private static void AddCell(NPOI.SS.UserModel.ISheet sheet, int rowIx, int colIx,
                double value)
        {
            sheet.CreateRow(rowIx).CreateCell((short)colIx).SetCellValue(value);
        }
    }
}