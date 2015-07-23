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

namespace TestCases.SS.UserModel
{
    using System;

    using NUnit.Framework;

    using NPOI.SS;
    using NPOI.SS.Util;
    using TestCases.SS;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.UserModel;

    /**
     * Tests row Shifting capabilities.
     *
     * @author Shawn Laubach (slaubach at apache dot com)
     * @author Toshiaki Kamoshida (kamoshida.Toshiaki at future dot co dot jp)
     */
    public class BaseTestSheetShiftRows
    {

        private ITestDataProvider _testDataProvider;
        public BaseTestSheetShiftRows()
        {
            _testDataProvider = TestCases.HSSF.HSSFITestDataProvider.Instance;
        }
        protected BaseTestSheetShiftRows(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }

        /**
         * Tests the ShiftRows function.  Does three different Shifts.
         * After each Shift, Writes the workbook to file and Reads back to
         * Check.  This ensures that if some Changes code that breaks
         * writing or what not, they realize it.
         *
         * @param sampleName the sample file to Test against
         */
        [Test]
        public void TestShiftRows()
        {
            // Read Initial file in
            String sampleName = "SimpleMultiCell." + _testDataProvider.StandardFileNameExtension;
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(sampleName);
            ISheet s = wb.GetSheetAt(0);

            // Shift the second row down 1 and write to temp file
            s.ShiftRows(1, 1, 1);
            {
                Console.WriteLine("Shift the second row down 1");
                var msg = string.Format("1a {0}-{1}-{2}-{3}-{4}-{5}", GetRowValue(s, 0), GetRowValue(s, 1), GetRowValue(s, 2), GetRowValue(s, 3), GetRowValue(s, 4), GetRowValue(s, 5));
                Console.WriteLine(msg);
            }

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            // Read from temp file and check the number of cells in each
            // row (in original file each row was unique)
            s = wb.GetSheetAt(0);
            {
                var msg = string.Format("1b {0}-{1}-{2}-{3}-{4}-{5}", GetRowValue(s, 0), GetRowValue(s, 1), GetRowValue(s, 2), GetRowValue(s, 3), GetRowValue(s, 4), GetRowValue(s, 5));
                Console.WriteLine(msg);
            }

            Assert.AreEqual(s.GetRow(0).PhysicalNumberOfCells, 1);
            ConfirmEmptyRow(s, 1);
            Assert.AreEqual(s.GetRow(2).PhysicalNumberOfCells, 2);
            Assert.AreEqual(s.GetRow(3).PhysicalNumberOfCells, 4);
            Assert.AreEqual(s.GetRow(4).PhysicalNumberOfCells, 5);

            // Shift rows 1-3 down 3 in the current one.  This Tests when
            // 1 row is blank.  Write to a another temp file
            s.ShiftRows(0, 2, 3);
            {
                Console.WriteLine("Shift rows 1-3 down 3 in the current one");
                var msg = string.Format("2a {0}-{1}-{2}-{3}-{4}-{5}", GetRowValue(s, 0), GetRowValue(s, 1), GetRowValue(s, 2), GetRowValue(s, 3), GetRowValue(s, 4), GetRowValue(s, 5));
                Console.WriteLine(msg);
            }
            wb = _testDataProvider.WriteOutAndReadBack(wb);

            // Read and ensure things are where they should be
            s = wb.GetSheetAt(0);
            {
                var msg = string.Format("2b {0}-{1}-{2}-{3}-{4}-{5}", GetRowValue(s, 0), GetRowValue(s, 1), GetRowValue(s, 2), GetRowValue(s, 3), GetRowValue(s, 4), GetRowValue(s, 5));
                Console.WriteLine(msg);
            }
            ConfirmEmptyRow(s, 0);
            ConfirmEmptyRow(s, 1);
            ConfirmEmptyRow(s, 2);
            Assert.AreEqual(1, s.GetRow(3).PhysicalNumberOfCells);
            ConfirmEmptyRow(s, 4);
            Assert.AreEqual(2, s.GetRow(5).PhysicalNumberOfCells);

            // Read the first file again
            wb = _testDataProvider.OpenSampleWorkbook(sampleName);
            s = wb.GetSheetAt(0);

            // Shift rows 3 and 4 up and write to temp file
            s.ShiftRows(2, 3, -2);
            {
                Console.WriteLine("Shift rows 3 and 4 up");
                var msg = string.Format("3a {0}-{1}-{2}-{3}-{4}-{5}", GetRowValue(s, 0), GetRowValue(s, 1), GetRowValue(s, 2), GetRowValue(s, 3), GetRowValue(s, 4), GetRowValue(s, 5));
                Console.WriteLine(msg);
            }
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            s = wb.GetSheetAt(0);
            {
                var msg = string.Format("3b {0}-{1}-{2}-{3}-{4}-{5}", GetRowValue(s, 0), GetRowValue(s, 1), GetRowValue(s, 2), GetRowValue(s, 3), GetRowValue(s, 4), GetRowValue(s, 5));
                Console.WriteLine(msg);
            }
            Assert.AreEqual(3, s.GetRow(0).PhysicalNumberOfCells);
            Assert.AreEqual(4, s.GetRow(1).PhysicalNumberOfCells);
            ConfirmEmptyRow(s, 2);
            ConfirmEmptyRow(s, 3);
            Assert.AreEqual(5, s.GetRow(4).PhysicalNumberOfCells);
        }
        private string GetRowValue(ISheet s, int rowIx)
        {
            IRow row = s.GetRow(rowIx);
            if (row == null) { return "null"; }
            return row.PhysicalNumberOfCells.ToString();
        }
        private string GetRowNum(ISheet s, int rowIx)
        {
            IRow row = s.GetRow(rowIx);
            if (row == null) { return "null"; }
            return row.RowNum.ToString();
        }
        private void ConfirmEmptyRow(ISheet s, int rowIx)
        {
            IRow row = s.GetRow(rowIx);
            Assert.IsTrue(row == null || row.PhysicalNumberOfCells == 0);
        }

        /**
         * Tests when rows are null.
         */
        [Test]
        public void TestShiftRow()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            ISheet s = b.CreateSheet();
            s.CreateRow(0).CreateCell(0).SetCellValue("TEST1");
            s.CreateRow(3).CreateCell(0).SetCellValue("TEST2");
            s.ShiftRows(0, 4, 1);
        }

        /**
         * Tests when Shifting the first row.
         */
        [Test]
        public void TestActiveCell()
        {
            IWorkbook b = _testDataProvider.CreateWorkbook();
            ISheet s = b.CreateSheet();

            s.CreateRow(0).CreateCell(0).SetCellValue("TEST1");
            s.CreateRow(3).CreateCell(0).SetCellValue("TEST2");
            s.ShiftRows(0, 4, 1);
        }

        /**
         * When Shifting rows, the page breaks should go with it
         */
        [Test]
        public virtual void TestShiftRowBreaks()
        { // TODO - enable XSSF Test
            IWorkbook b = _testDataProvider.CreateWorkbook();
            ISheet s = b.CreateSheet();
            IRow row = s.CreateRow(4);
            row.CreateCell(0).SetCellValue("test");
            s.SetRowBreak(4);

            s.ShiftRows(4, 4, 2);
            Assert.IsTrue(s.IsRowBroken(6), "Row number 6 should have a pagebreak");
        }
        [Test]
        public void TestShiftWithComments()
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook("comments." + _testDataProvider.StandardFileNameExtension);

            ISheet sheet = wb.GetSheet("Sheet1");
            Assert.AreEqual(3, sheet.LastRowNum);

            // Verify comments are in the position expected
            Assert.IsNotNull(sheet.GetCellComment(0, 0));
            Assert.IsNull(sheet.GetCellComment(1, 0));
            Assert.IsNotNull(sheet.GetCellComment(2, 0));
            Assert.IsNotNull(sheet.GetCellComment(3, 0));

            String comment1 = sheet.GetCellComment(0, 0).String.String;
            Assert.AreEqual(comment1, "comment top row1 (index0)\n");
            String comment3 = sheet.GetCellComment(2, 0).String.String;
            Assert.AreEqual(comment3, "comment top row3 (index2)\n");
            String comment4 = sheet.GetCellComment(3, 0).String.String;
            Assert.AreEqual(comment4, "comment top row4 (index3)\n");

            //Workbook wbBack = _testDataProvider.writeOutAndReadBack(wb);

            // Shifting all but first line down to Test comments Shifting
            sheet.ShiftRows(1, sheet.LastRowNum, 1, true, true);

            // Test that comments were Shifted as expected
            Assert.AreEqual(4, sheet.LastRowNum);
            Assert.IsNotNull(sheet.GetCellComment(0, 0));
            Assert.IsNull(sheet.GetCellComment(1, 0));
            Assert.IsNull(sheet.GetCellComment(2, 0));
            Assert.IsNotNull(sheet.GetCellComment(3, 0));
            Assert.IsNotNull(sheet.GetCellComment(4, 0));

            String comment1_Shifted = sheet.GetCellComment(0, 0).String.String;
            Assert.AreEqual(comment1, comment1_Shifted);
            String comment3_Shifted = sheet.GetCellComment(3, 0).String.String;
            Assert.AreEqual(comment3, comment3_Shifted);
            String comment4_Shifted = sheet.GetCellComment(4, 0).String.String;
            Assert.AreEqual(comment4, comment4_Shifted);

            // Write out and read back in again
            // Ensure that the Changes were persisted
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            sheet = wb.GetSheet("Sheet1");
            Assert.AreEqual(4, sheet.LastRowNum);

            // Verify comments are in the position expected After the shift
            Assert.IsNotNull(sheet.GetCellComment(0, 0));
            Assert.IsNull(sheet.GetCellComment(1, 0));
            Assert.IsNull(sheet.GetCellComment(2, 0));
            Assert.IsNotNull(sheet.GetCellComment(3, 0));
            Assert.IsNotNull(sheet.GetCellComment(4, 0));

            comment1_Shifted = sheet.GetCellComment(0, 0).String.String;
            Assert.AreEqual(comment1, comment1_Shifted);
            comment3_Shifted = sheet.GetCellComment(3, 0).String.String;
            Assert.AreEqual(comment3, comment3_Shifted);
            comment4_Shifted = sheet.GetCellComment(4, 0).String.String;
            Assert.AreEqual(comment4, comment4_Shifted);

            // Shifting back up again, now two rows
            sheet.ShiftRows(2, sheet.LastRowNum, -2, true, true);

            // TODO: it seems HSSFSheet does not correctly remove comments from rows that are overwritten
            // by Shifting rows...
            if (!(wb is HSSFWorkbook))
            {
                Assert.AreEqual(2, sheet.LastRowNum);

                // Verify comments are in the position expected
                Assert.IsNull(sheet.GetCellComment(0, 0),
                    "Had: " + (sheet.GetCellComment(0, 0) == null ? "null" : sheet.GetCellComment(0, 0).String.String));
                Assert.IsNotNull(sheet.GetCellComment(1, 0));
                Assert.IsNotNull(sheet.GetCellComment(2, 0));
            }

            comment1 = sheet.GetCellComment(1, 0).String.String;
            Assert.AreEqual(comment1, "comment top row3 (index2)\n");
            String comment2 = sheet.GetCellComment(2, 0).String.String;
            Assert.AreEqual(comment2, "comment top row4 (index3)\n");

        }
        [Test]
        public void TestShiftWithNames()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet("Sheet1");
            wb.CreateSheet("Sheet2");
            IRow row = sheet1.CreateRow(0);
            row.CreateCell(0).SetCellValue(1.1);
            row.CreateCell(1).SetCellValue(2.2);

            IName name1 = wb.CreateName();
            name1.NameName = (/*setter*/"name1");
            name1.RefersToFormula = (/*setter*/"Sheet1!$A$1+Sheet1!$B$1");

            IName name2 = wb.CreateName();
            name2.NameName = (/*setter*/"name2");
            name2.RefersToFormula = (/*setter*/"Sheet1!$A$1");

            //refers to A1 but on Sheet2. Should stay unaffected.
            IName name3 = wb.CreateName();
            name3.NameName = (/*setter*/"name3");
            name3.RefersToFormula = (/*setter*/"Sheet2!$A$1");

            //The scope of this one is Sheet2. Should stay unaffected.
            IName name4 = wb.CreateName();
            name4.NameName = (/*setter*/"name4");
            name4.RefersToFormula = (/*setter*/"A1");
            name4.SheetIndex = (/*setter*/1);

            sheet1.ShiftRows(0, 1, 2);  //shift down the top row on Sheet1.
            name1 = wb.GetNameAt(0);
            Assert.AreEqual(name1.RefersToFormula, "Sheet1!$A$3+Sheet1!$B$3");

            name2 = wb.GetNameAt(1);
            Assert.AreEqual(name2.RefersToFormula, "Sheet1!$A$3");

            //name3 and name4 refer to Sheet2 and should not be affected
            name3 = wb.GetNameAt(2);
            Assert.AreEqual(name3.RefersToFormula, "Sheet2!$A$1");

            name4 = wb.GetNameAt(3);
            Assert.AreEqual(name4.RefersToFormula, "A1");
        }
        [Test]
        public void TestShiftWithMergedRegions()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(1.1);
            row.CreateCell(1).SetCellValue(2.2);
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
            Assert.AreEqual("A1:C1", region.FormatAsString());

            sheet.AddMergedRegion(region);

            sheet.ShiftRows(0, 1, 2);
            region = sheet.GetMergedRegion(0);
            Assert.AreEqual("A3:C3", region.FormatAsString());
        }

        /**
         * See bug #34023
         */
        [Test]
        public void TestShiftWithFormulas()
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook("ForShifting." + _testDataProvider.StandardFileNameExtension);

            ISheet sheet = wb.GetSheet("Sheet1");
            Assert.AreEqual(20, sheet.LastRowNum);

            ConfirmRow(sheet, 0, 1, 171, 1, "ROW(D1)", "100+B1", "COUNT(D1:E1)");
            ConfirmRow(sheet, 1, 2, 172, 1, "ROW(D2)", "100+B2", "COUNT(D2:E2)");
            ConfirmRow(sheet, 2, 3, 173, 1, "ROW(D3)", "100+B3", "COUNT(D3:E3)");

            ConfirmCell(sheet, 6, 1, 271, "200+B1");
            ConfirmCell(sheet, 7, 1, 272, "200+B2");
            ConfirmCell(sheet, 8, 1, 273, "200+B3");

            ConfirmCell(sheet, 14, 0, 0.0, "A12"); // the cell referred to by this formula will be Replaced

            // -----------
            // Row index 1 -> 11 (row "2" -> row "12")
            sheet.ShiftRows(1, 1, 10);

            // Now check what sheet looks like After Move

            // no Changes on row "1"
            ConfirmRow(sheet, 0, 1, 171, 1, "ROW(D1)", "100+B1", "COUNT(D1:E1)");

            // row "2" is now empty
            ConfirmEmptyRow(sheet, 1);

            // Row "2" Moved to row "12", and the formula has been updated.
            // note however that the cached formula result (2) has not been updated. (POI differs from Excel here)
            ConfirmRow(sheet, 11, 2, 172, 1, "ROW(D12)", "100+B12", "COUNT(D12:E12)");

            // no Changes on row "3"
            ConfirmRow(sheet, 2, 3, 173, 1, "ROW(D3)", "100+B3", "COUNT(D3:E3)");


            ConfirmCell(sheet, 14, 0, 0.0, "#REF!");


            // Formulas on rows that weren't Shifted:
            ConfirmCell(sheet, 6, 1, 271, "200+B1");
            ConfirmCell(sheet, 7, 1, 272, "200+B12"); // this one Moved
            ConfirmCell(sheet, 8, 1, 273, "200+B3");

            // check formulas on other sheets
            ISheet sheet2 = wb.GetSheet("Sheet2");
            ConfirmCell(sheet2, 0, 0, 371, "300+Sheet1!B1");
            ConfirmCell(sheet2, 1, 0, 372, "300+Sheet1!B12");
            ConfirmCell(sheet2, 2, 0, 373, "300+Sheet1!B3");

            ConfirmCell(sheet2, 11, 0, 300, "300+Sheet1!#REF!");


            // Note - named ranges formulas have not been updated
        }

        private static void ConfirmRow(ISheet sheet, int rowIx, double valA, double valB, double valC,
                    String formulaA, String formulaB, String formulaC)
        {
            ConfirmCell(sheet, rowIx, 4, valA, formulaA);
            ConfirmCell(sheet, rowIx, 5, valB, formulaB);
            ConfirmCell(sheet, rowIx, 6, valC, formulaC);
        }

        private static void ConfirmCell(ISheet sheet, int rowIx, int colIx,
                double expectedValue, String expectedFormula)
        {
            ICell cell = sheet.GetRow(rowIx).GetCell(colIx);
            Assert.AreEqual(expectedValue, cell.NumericCellValue, 0.0);
            Assert.AreEqual(expectedFormula, cell.CellFormula);
        }
        [Test]
        public void TestShiftSharedFormulasBug54206()
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook("54206." + _testDataProvider.StandardFileNameExtension);

            ISheet sheet = wb.GetSheetAt(0);
            Assert.AreEqual(sheet.GetRow(3).GetCell(6).CellFormula, "SUMIF($B$19:$B$82,$B4,G$19:G$82)");
            Assert.AreEqual(sheet.GetRow(3).GetCell(7).CellFormula, "SUMIF($B$19:$B$82,$B4,H$19:H$82)");
            Assert.AreEqual(sheet.GetRow(3).GetCell(8).CellFormula, "SUMIF($B$19:$B$82,$B4,I$19:I$82)");

            Assert.AreEqual(sheet.GetRow(14).GetCell(6).CellFormula, "SUMIF($B$19:$B$82,$B15,G$19:G$82)");
            Assert.AreEqual(sheet.GetRow(14).GetCell(7).CellFormula, "SUMIF($B$19:$B$82,$B15,H$19:H$82)");
            Assert.AreEqual(sheet.GetRow(14).GetCell(8).CellFormula, "SUMIF($B$19:$B$82,$B15,I$19:I$82)");

            // now the whole block G4L:15
            for (int i = 3; i <= 14; i++)
            {
                for (int j = 6; j <= 8; j++)
                {
                    String col = CellReference.ConvertNumToColString(j);
                    String expectedFormula = "SUMIF($B$19:$B$82,$B" + (i + 1) + "," + col + "$19:" + col + "$82)";
                    Assert.AreEqual(expectedFormula, sheet.GetRow(i).GetCell(j).CellFormula);
                }
            }

            Assert.AreEqual(sheet.GetRow(23).GetCell(9).CellFormula, "SUM(G24:I24)");
            Assert.AreEqual(sheet.GetRow(24).GetCell(9).CellFormula, "SUM(G25:I25)");
            Assert.AreEqual(sheet.GetRow(25).GetCell(9).CellFormula, "SUM(G26:I26)");

            sheet.ShiftRows(24, sheet.LastRowNum, 4, true, false);

            Assert.AreEqual(sheet.GetRow(3).GetCell(6).CellFormula, "SUMIF($B$19:$B$86,$B4,G$19:G$86)");
            Assert.AreEqual(sheet.GetRow(3).GetCell(7).CellFormula, "SUMIF($B$19:$B$86,$B4,H$19:H$86)");
            Assert.AreEqual(sheet.GetRow(3).GetCell(8).CellFormula, "SUMIF($B$19:$B$86,$B4,I$19:I$86)");

            Assert.AreEqual(sheet.GetRow(14).GetCell(6).CellFormula, "SUMIF($B$19:$B$86,$B15,G$19:G$86)");
            Assert.AreEqual(sheet.GetRow(14).GetCell(7).CellFormula, "SUMIF($B$19:$B$86,$B15,H$19:H$86)");
            Assert.AreEqual(sheet.GetRow(14).GetCell(8).CellFormula, "SUMIF($B$19:$B$86,$B15,I$19:I$86)");

            // now the whole block G4L:15
            for (int i = 3; i <= 14; i++)
            {
                for (int j = 6; j <= 8; j++)
                {
                    String col = CellReference.ConvertNumToColString(j);
                    String expectedFormula = "SUMIF($B$19:$B$86,$B" + (i + 1) + "," + col + "$19:" + col + "$86)";
                    Assert.AreEqual(expectedFormula, sheet.GetRow(i).GetCell(j).CellFormula);
                }
            }

            Assert.AreEqual(sheet.GetRow(23).GetCell(9).CellFormula, "SUM(G24:I24)");

            // shifted rows
            Assert.IsTrue(sheet.GetRow(24) == null || sheet.GetRow(24).GetCell(9) == null);
            Assert.IsTrue(sheet.GetRow(25) == null || sheet.GetRow(25).GetCell(9) == null);
            Assert.IsTrue(sheet.GetRow(26) == null || sheet.GetRow(26).GetCell(9) == null);
            Assert.IsTrue(sheet.GetRow(27) == null || sheet.GetRow(27).GetCell(9) == null);

            Assert.AreEqual(sheet.GetRow(28).GetCell(9).CellFormula, "SUM(G29:I29)");
            Assert.AreEqual(sheet.GetRow(29).GetCell(9).CellFormula, "SUM(G30:I30)");

        }

        [Test]
        public void TestBug55280()
        {
            IWorkbook w = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet s = w.CreateSheet();
                for (int row = 0; row < 5000; ++row)
                    s.AddMergedRegion(new CellRangeAddress(row, row, 0, 3));

                s.ShiftRows(0, 4999, 1);        // takes a long time...
            }
            finally
            {
                w.Close();
            }
        }

    }
}
