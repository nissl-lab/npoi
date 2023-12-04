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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.HSSF.UserModel;

    using NUnit.Framework;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * Tests row shifting capabilities.
     *
     *
     * @author Shawn Laubach (slaubach at apache dot com)
     * @author Toshiaki Kamoshida (kamoshida.toshiaki at future dot co dot jp)
     */
    [TestFixture]
    public class TestSheetShiftRows
    {

        /**
         * Tests the ShiftRows function.  Does three different shifts.
         * After each shift, Writes the workbook to file and reads back to
         * Check.  This ensures that if some changes code that breaks
         * writing or what not, they realize it.
         *
         * @author Shawn Laubach (slaubach at apache dot org)
         */
        [Test]
        public void TestShiftRows()
        {
            // Read initial file in
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleMultiCell.xls");
            NPOI.SS.UserModel.ISheet s = wb.GetSheetAt(0);

            // Shift the second row down 1 and Write to temp file
            s.ShiftRows(1, 1, 1);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            // Read from temp file and Check the number of cells in each
            // row (in original file each row was unique)
            s = wb.GetSheetAt(0);

            Assert.AreEqual(1, s.GetRow(0).PhysicalNumberOfCells);
            ConfirmEmptyRow(s, 1);
            Assert.AreEqual(2, s.GetRow(2).PhysicalNumberOfCells);
            Assert.AreEqual(4, s.GetRow(3).PhysicalNumberOfCells);
            Assert.AreEqual(5, s.GetRow(4).PhysicalNumberOfCells);

            // Shift rows 1-3 down 3 in the current one.  This Tests when
            // 1 row is blank.  Write to a another temp file
            s.ShiftRows(0, 2, 3);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            // Read and ensure things are where they should be
            s = wb.GetSheetAt(0);
            ConfirmEmptyRow(s, 0);
            ConfirmEmptyRow(s, 1);
            ConfirmEmptyRow(s, 2);
            Assert.AreEqual(1, s.GetRow(3).PhysicalNumberOfCells);
            ConfirmEmptyRow(s, 4);
            Assert.AreEqual(2, s.GetRow(5).PhysicalNumberOfCells);

            // Read the first file again
            wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleMultiCell.xls");
            s = wb.GetSheetAt(0);

            // Shift rows 3 and 4 up and Write to temp file
            s.ShiftRows(2, 3, -2);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            Assert.AreEqual(s.GetRow(0).PhysicalNumberOfCells, 3);
            Assert.AreEqual(s.GetRow(1).PhysicalNumberOfCells, 4);
            ConfirmEmptyRow(s, 2);
            ConfirmEmptyRow(s, 3);
            Assert.AreEqual(s.GetRow(4).PhysicalNumberOfCells, 5);
        }
        private static void ConfirmEmptyRow(NPOI.SS.UserModel.ISheet s, int rowIx)
        {
            IRow row = s.GetRow(rowIx);
            Assert.IsTrue(row == null || row.PhysicalNumberOfCells == 0);
        }
        /**
         * Tests when rows are null.
         *
         * @author Toshiaki Kamoshida (kamoshida.toshiaki at future dot co dot jp)
         */
        [Test]
        public void TestShiftRow()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = b.CreateSheet();
            s.CreateRow(0).CreateCell(0).SetCellValue("TEST1");
            s.CreateRow(3).CreateCell(0).SetCellValue("TEST2");
            s.ShiftRows(0, 4, 1);
        }

        /**
         * Tests when shifting the first row.
         *
         * @author Toshiaki Kamoshida (kamoshida.toshiaki at future dot co dot jp)
         */
        [Test]
        public void TestShiftRow0()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = b.CreateSheet();
            s.CreateRow(0).CreateCell(0).SetCellValue("TEST1");
            s.CreateRow(3).CreateCell(0).SetCellValue("TEST2");
            s.ShiftRows(0, 4, 1);
        }

        /**
         * When shifting rows, the page breaks should go with it
         *
         */
        [Test]
        public void TestShiftRowBreaks()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = b.CreateSheet();
            IRow row = s.CreateRow(4);
            row.CreateCell(0).SetCellValue("Test");
            s.SetRowBreak(4);

            s.ShiftRows(4, 4, 2);
            Assert.IsTrue(s.IsRowBroken(6), "Row number 6 should have a pagebreak");

        }

        [Test]
        public void TestShiftWithComments()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("comments.xls");

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheet("Sheet1");
            Assert.AreEqual(3, sheet.LastRowNum);

            // Verify comments are in the position expected
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(0, 0)));
            Assert.IsNull(sheet.GetCellComment(new CellAddress(1, 0)));
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(2, 0)));
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(3, 0)));

            String comment1 = sheet.GetCellComment(new CellAddress(0, 0)).String.String;
            Assert.AreEqual(comment1, "comment top row1 (index0)\n");
            String comment3 = sheet.GetCellComment(new CellAddress(2, 0)).String.String;
            Assert.AreEqual(comment3, "comment top row3 (index2)\n");
            String comment4 = sheet.GetCellComment(new CellAddress(3, 0)).String.String;
            Assert.AreEqual(comment4, "comment top row4 (index3)\n");

            // Shifting all but first line down to Test comments shifting
            sheet.ShiftRows(1, sheet.LastRowNum, 1, true, true);
            MemoryStream outputStream = new MemoryStream();
            wb.Write(outputStream);

            // Test that comments were shifted as expected
            Assert.AreEqual(4, sheet.LastRowNum);
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(0, 0)));
            Assert.IsNull(sheet.GetCellComment(new CellAddress(1, 0)));
            Assert.IsNull(sheet.GetCellComment(new CellAddress(2, 0)));
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(3, 0)));
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(4, 0)));

            String comment1_shifted = sheet.GetCellComment(new CellAddress(0, 0)).String.String;
            Assert.AreEqual(comment1, comment1_shifted);
            String comment3_shifted = sheet.GetCellComment(new CellAddress(3, 0)).String.String;
            Assert.AreEqual(comment3, comment3_shifted);
            String comment4_shifted = sheet.GetCellComment(new CellAddress(4, 0)).String.String;
            Assert.AreEqual(comment4, comment4_shifted);

            // Write out and read back in again
            // Ensure that the changes were persisted
            wb = new HSSFWorkbook(new MemoryStream(outputStream.ToArray()));
            sheet = wb.GetSheet("Sheet1");
            Assert.AreEqual(4, sheet.LastRowNum);

            // Verify comments are in the position expected after the shift
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(0, 0)));
            Assert.IsNull(sheet.GetCellComment(new CellAddress(1, 0)));
            Assert.IsNull(sheet.GetCellComment(new CellAddress(2, 0)));
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(3, 0)));
            Assert.IsNotNull(sheet.GetCellComment(new CellAddress(4, 0)));

            comment1_shifted = sheet.GetCellComment(new CellAddress(0, 0)).String.String;
            Assert.AreEqual(comment1, comment1_shifted);
            comment3_shifted = sheet.GetCellComment(new CellAddress(3, 0)).String.String;
            Assert.AreEqual(comment3, comment3_shifted);
            comment4_shifted = sheet.GetCellComment(new CellAddress(4, 0)).String.String;
            Assert.AreEqual(comment4, comment4_shifted);
        }

        /**
         * See bug #34023
         */
        [Test]
        public void TestShiftWithFormulas()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ForShifting.xls");

            NPOI.SS.UserModel.ISheet sheet = wb.GetSheet("Sheet1");
            Assert.AreEqual(20, sheet.LastRowNum);

            ConfirmRow(sheet, 0, 1, 171, 1, "ROW(D1)", "100+B1", "COUNT(D1:E1)");
            ConfirmRow(sheet, 1, 2, 172, 1, "ROW(D2)", "100+B2", "COUNT(D2:E2)");
            ConfirmRow(sheet, 2, 3, 173, 1, "ROW(D3)", "100+B3", "COUNT(D3:E3)");

            ConfirmCell(sheet, 6, 1, 271, "200+B1");
            ConfirmCell(sheet, 7, 1, 272, "200+B2");
            ConfirmCell(sheet, 8, 1, 273, "200+B3");

            ConfirmCell(sheet, 14, 0, 0.0, "A12"); // the cell referred to by this formula will be replaced

            // -----------
            // Row index 1 -> 11 (row "2" -> row "12")
            sheet.ShiftRows(1, 1, 10);

            // Now Check what sheet looks like after move

            // no changes on row "1"
            ConfirmRow(sheet, 0, 1, 171, 1, "ROW(D1)", "100+B1", "COUNT(D1:E1)");

            // row "2" is now empty
            Assert.AreEqual(0, sheet.GetRow(1).PhysicalNumberOfCells);

            // Row "2" moved to row "12", and the formula has been updated.
            // note however that the cached formula result (2) has not been updated. (POI differs from Excel here)
            ConfirmRow(sheet, 11, 2, 172, 1, "ROW(D12)", "100+B12", "COUNT(D12:E12)");

            // no changes on row "3"
            ConfirmRow(sheet, 2, 3, 173, 1, "ROW(D3)", "100+B3", "COUNT(D3:E3)");


            ConfirmCell(sheet, 14, 0, 0.0, "#REF!");


            // Formulas on rows that weren't shifted:
            ConfirmCell(sheet, 6, 1, 271, "200+B1");
            ConfirmCell(sheet, 7, 1, 272, "200+B12"); // this one moved
            ConfirmCell(sheet, 8, 1, 273, "200+B3");

            // Check formulas on other sheets
            NPOI.SS.UserModel.ISheet sheet2 = wb.GetSheet("Sheet2");
            ConfirmCell(sheet2, 0, 0, 371, "300+Sheet1!B1");
            ConfirmCell(sheet2, 1, 0, 372, "300+Sheet1!B12");
            ConfirmCell(sheet2, 2, 0, 373, "300+Sheet1!B3");

            ConfirmCell(sheet2, 11, 0, 300, "300+Sheet1!#REF!");


            // Note - named ranges formulas have not been updated
        }

        private static void ConfirmRow(NPOI.SS.UserModel.ISheet sheet, int rowIx, double valA, double valB, double valC,
                    String formulaA, String formulaB, String formulaC)
        {
            ConfirmCell(sheet, rowIx, 4, valA, formulaA);
            ConfirmCell(sheet, rowIx, 5, valB, formulaB);
            ConfirmCell(sheet, rowIx, 6, valC, formulaC);
        }

        private static void ConfirmCell(NPOI.SS.UserModel.ISheet sheet, int rowIx, int colIx,
                double expectedValue, String expectedFormula)
        {
            ICell cell = sheet.GetRow(rowIx).GetCell(colIx);
            Assert.AreEqual(expectedValue, cell.NumericCellValue, 0.0);
            Assert.AreEqual(expectedFormula, cell.CellFormula);
        }
    }
}