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
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;

    /**
     * Common superclass for Testing {@link NPOI.xssf.UserModel.XSSFCell}  and
     * {@link NPOI.HSSF.UserModel.HSSFCell}
     */
    public abstract class BaseTestSheet
    {
        private static int ROW_COUNT = 40000;
        private ITestDataProvider _testDataProvider;
        protected BaseTestSheet(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }

        protected virtual void TrackColumnsForAutoSizingIfSXSSF(ISheet sheet)
        {
            // do nothing for Sheet base class. This will be overridden for SXSSFSheets.
        }
        [Test]
        public void TestCreateRow()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            Assert.AreEqual(0, sheet.PhysicalNumberOfRows);

            //Test that we Get null for undefined rownumber
            Assert.IsNull(sheet.GetRow(1));

            // Test row creation with consecutive indexes
            IRow row1 = sheet.CreateRow(0);
            IRow row2 = sheet.CreateRow(1);
            Assert.AreEqual(0, row1.RowNum);
            Assert.AreEqual(1, row2.RowNum);
            IEnumerator it = sheet.GetRowEnumerator();
            Assert.IsTrue(it.MoveNext());
            Assert.AreSame(row1, it.Current);
            Assert.IsTrue(it.MoveNext());
            Assert.AreSame(row2, it.Current);
            Assert.AreEqual(1, sheet.LastRowNum);

            // Test row creation with non consecutive index
            IRow row101 = sheet.CreateRow(100);
            Assert.IsNotNull(row101);
            Assert.AreEqual(100, sheet.LastRowNum);
            Assert.AreEqual(3, sheet.PhysicalNumberOfRows);

            // Test overwriting an existing row
            IRow row2_ovrewritten = sheet.CreateRow(1);
            ICell cell = row2_ovrewritten.CreateCell(0);
            cell.SetCellValue(100);
            IEnumerator it2 = sheet.GetRowEnumerator();
            Assert.IsTrue(it2.MoveNext());
            Assert.AreSame(row1, it2.Current);
            Assert.IsTrue(it2.MoveNext());
            IRow row2_ovrewritten_ref = (IRow)it2.Current;
            Assert.AreSame(row2_ovrewritten, row2_ovrewritten_ref);
            Assert.AreEqual(100.0, row2_ovrewritten_ref.GetCell(0).NumericCellValue, 0.0);

            workbook.Close();
        }


        [Test]
        public void CreateRowBeforeFirstRow()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sh = workbook.CreateSheet();
            sh.CreateRow(0);
            try
            {
                //Negative rows not allowed
                Assert.Throws<ArgumentException>(() =>
                {
                    sh.CreateRow(-1);
                });

            }
            finally
            {
                workbook.Close();
            }
        }

        [Test]
        public void CreateRowAfterLastRow()
        {
            SpreadsheetVersion version = _testDataProvider.GetSpreadsheetVersion();
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sh = workbook.CreateSheet();
            sh.CreateRow(version.LastRowIndex);
            try
            {
                Assert.Throws<ArgumentException>(()=> {
                    sh.CreateRow(version.LastRowIndex + 1);
                });
            }
            finally
            {
                workbook.Close();
            }
        }


        [Test]
        public void TestRemoveRow()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = workbook.CreateSheet();
            Assert.AreEqual(0, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(0, sheet1.FirstRowNum);
            Assert.AreEqual(0, sheet1.LastRowNum);

            IRow row0 = sheet1.CreateRow(0);
            Assert.AreEqual(1, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(0, sheet1.FirstRowNum);
            Assert.AreEqual(0, sheet1.LastRowNum);
            sheet1.RemoveRow(row0);
            Assert.AreEqual(0, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(0, sheet1.FirstRowNum);
            Assert.AreEqual(0, sheet1.LastRowNum);

            sheet1.CreateRow(1);
            IRow row2 = sheet1.CreateRow(2);
            Assert.AreEqual(2, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(1, sheet1.FirstRowNum);
            Assert.AreEqual(2, sheet1.LastRowNum);

            Assert.IsNotNull(sheet1.GetRow(1));
            Assert.IsNotNull(sheet1.GetRow(2));
            sheet1.RemoveRow(row2);
            Assert.IsNotNull(sheet1.GetRow(1));
            Assert.IsNull(sheet1.GetRow(2));
            Assert.AreEqual(1, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(1, sheet1.FirstRowNum);
            Assert.AreEqual(1, sheet1.LastRowNum);

            IRow row3 = sheet1.CreateRow(3);
            ISheet sheet2 = workbook.CreateSheet();
            ArgumentException e = Assert.Throws<ArgumentException>(() =>
            {
                sheet2.RemoveRow(row3);
            });
            Assert.AreEqual("Specified row does not belong to this sheet", e.Message);

            workbook.Close();
        }
        [Test]
        public virtual void CloneSheet()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ICreationHelper factory = workbook.GetCreationHelper();
            ISheet sheet = workbook.CreateSheet("Test Clone");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            ICell cell2 = row.CreateCell(1);
            cell.SetCellValue(factory.CreateRichTextString("Clone_test"));
            cell2.CellFormula = "SIN(1)";

            ISheet clonedSheet = workbook.CloneSheet(0);
            IRow clonedRow = clonedSheet.GetRow(0);

            //Check for a good clone
            Assert.AreEqual(clonedRow.GetCell(0).RichStringCellValue.String, "Clone_test");

            //Check that the cells are not somehow linked
            cell.SetCellValue(factory.CreateRichTextString("Difference Check"));
            cell2.CellFormula = (/*setter*/"cos(2)");
            if ("Difference Check".Equals(clonedRow.GetCell(0).RichStringCellValue.String))
            {
                Assert.Fail("string cell not properly Cloned");
            }
            if ("COS(2)".Equals(clonedRow.GetCell(1).CellFormula))
            {
                Assert.Fail("formula cell not properly Cloned");
            }
            Assert.AreEqual(clonedRow.GetCell(0).RichStringCellValue.String, "Clone_test");
            Assert.AreEqual(clonedRow.GetCell(1).CellFormula, "SIN(1)");

            workbook.Close();
        }

        /** Tests that the sheet name for multiple Clones of the same sheet is unique
         * BUG 37416
         */
        [Test]
        public virtual void CloneSheetMultipleTimes()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ICreationHelper factory = workbook.GetCreationHelper();
            ISheet sheet = workbook.CreateSheet("Test Clone");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(factory.CreateRichTextString("Clone_test"));
            //Clone the sheet multiple times
            workbook.CloneSheet(0);
            workbook.CloneSheet(0);

            Assert.IsNotNull(workbook.GetSheet("Test Clone"));
            Assert.IsNotNull(workbook.GetSheet("Test Clone (2)"));
            Assert.AreEqual("Test Clone (3)", workbook.GetSheetName(2));
            Assert.IsNotNull(workbook.GetSheet("Test Clone (3)"));

            workbook.RemoveSheetAt(0);
            workbook.RemoveSheetAt(0);
            workbook.RemoveSheetAt(0);
            workbook.CreateSheet("abc ( 123)");
            workbook.CloneSheet(0);
            Assert.AreEqual("abc (124)", workbook.GetSheetName(1));

            workbook.Close();
        }

        /**
         * Setting landscape and portrait stuff on new sheets
         */
        [Test]
        public void TestPrintSetupLandscapeNew()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheetL = wb1.CreateSheet("LandscapeS");
            ISheet sheetP = wb1.CreateSheet("LandscapeP");

            // Check two aspects of the print Setup
            Assert.IsFalse(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetP.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(1, sheetP.PrintSetup.Copies);

            // Change one on each
            sheetL.PrintSetup.Landscape = (/*setter*/true);
            sheetP.PrintSetup.Copies = (/*setter*/(short)3);

            // Check taken
            Assert.IsTrue(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetP.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(3, sheetP.PrintSetup.Copies);

            // Save and re-load, and check still there
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheetL = wb2.GetSheet("LandscapeS");
            sheetP = wb2.GetSheet("LandscapeP");

            Assert.IsTrue(sheetL.PrintSetup.Landscape);
            Assert.IsFalse(sheetP.PrintSetup.Landscape);
            Assert.AreEqual(1, sheetL.PrintSetup.Copies);
            Assert.AreEqual(3, sheetP.PrintSetup.Copies);

            wb2.Close();
        }


        /**
         * Disallow creating wholly or partially overlapping merged regions
         * as this results in a corrupted workbook
         */
        [Test]
        public void AddOverlappingMergedRegions()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            CellRangeAddress baseRegion = new CellRangeAddress(0, 1, 0, 1); //A1:B2
            sheet.AddMergedRegion(baseRegion);

            try
            {
                CellRangeAddress duplicateRegion = new CellRangeAddress(0, 1, 0, 1); //A1:B2
                sheet.AddMergedRegion(duplicateRegion);
                Assert.Fail("Should not be able to add a merged region (" + duplicateRegion.FormatAsString() + ") " +
                 "if sheet already contains the same merged region (" + baseRegion.FormatAsString() + ")");
            }
            catch (InvalidOperationException)
            {
            } 

            try
            {
                CellRangeAddress partiallyOverlappingRegion = new CellRangeAddress(1, 2, 1, 2); //B2:C3
                sheet.AddMergedRegion(partiallyOverlappingRegion);
                Assert.Fail("Should not be able to add a merged region (" + partiallyOverlappingRegion.FormatAsString() + ") " +
                 "if it partially overlaps with an existing merged region (" + baseRegion.FormatAsString() + ")");
            }
            catch (InvalidOperationException)
            {
            } 

            try
            {
                CellRangeAddress subsetRegion = new CellRangeAddress(0, 1, 0, 0); //A1:A2
                sheet.AddMergedRegion(subsetRegion);
                Assert.Fail("Should not be able to add a merged region (" + subsetRegion.FormatAsString() + ") " +
                 "if it is a formal subset of an existing merged region (" + baseRegion.FormatAsString() + ")");
            }
            catch (InvalidOperationException)
            {
            } 

            try
            {
                CellRangeAddress supersetRegion = new CellRangeAddress(0, 2, 0, 2); //A1:C3
                sheet.AddMergedRegion(supersetRegion);
                Assert.Fail("Should not be able to add a merged region (" + supersetRegion.FormatAsString() + ") " +
                 "if it is a formal superset of an existing merged region (" + baseRegion.FormatAsString() + ")");
            }
            catch (InvalidOperationException)
            {
            }

            CellRangeAddress disjointRegion = new CellRangeAddress(10, 11, 10, 11);
            sheet.AddMergedRegion(disjointRegion);

            wb.Close();
        }

        /*
        * Bug 56345: Reject single-cell merged regions
        */
        [Test]
        public void AddMergedRegionWithSingleCellShouldFail()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            CellRangeAddress region = CellRangeAddress.ValueOf("A1:A1");
            try
            {
                sheet.AddMergedRegion(region);
                Assert.Fail("Should not be able to add a single-cell merged region (" + region.FormatAsString() + ")");
            }
            catch (ArgumentException)
            {
                // expected
            }
            wb.Close();
        }


        /**
         * Test Adding merged regions. If the region's bounds are outside of the allowed range
         * then an ArgumentException should be thrown
         *
         */
        [Test]
        public void TestAddMerged()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            Assert.AreEqual(0, sheet.NumMergedRegions);
            SpreadsheetVersion ssVersion = _testDataProvider.GetSpreadsheetVersion();

            CellRangeAddress region = new CellRangeAddress(0, 1, 0, 1);
            sheet.AddMergedRegion(region);
            Assert.AreEqual(1, sheet.NumMergedRegions);

            try
            {
                region = new CellRangeAddress(-1, -1, -1, -1);
                sheet.AddMergedRegion(region);
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException)
            {
                // TODO           Assert.AreEqual("Minimum row number is 0.", e.Message);
            }
            try
            {
                region = new CellRangeAddress(0, 0, 0, ssVersion.LastColumnIndex + 1);
                sheet.AddMergedRegion(region);
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Maximum column number is " + ssVersion.LastColumnIndex, e.Message);
            }
            try
            {
                region = new CellRangeAddress(0, ssVersion.LastRowIndex + 1, 0, 1);
                sheet.AddMergedRegion(region);
                Assert.Fail("Expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Maximum row number is " + ssVersion.LastRowIndex, e.Message);
            }
            Assert.AreEqual(1, sheet.NumMergedRegions);

            wb.Close();
        }

        /**
         * When removing one merged region, it would break
         *
         */
        [Test]
        public void TestRemoveMerged()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            CellRangeAddress region = new CellRangeAddress(0, 1, 0, 1);
            sheet.AddMergedRegion(region);
            region = new CellRangeAddress(2, 3, 0, 1);
            sheet.AddMergedRegion(region);

            sheet.RemoveMergedRegion(0);

            region = sheet.GetMergedRegion(0);
            Assert.AreEqual(2, region.FirstRow, "Left over region should be starting at row 2");

            sheet.RemoveMergedRegion(0);

            Assert.AreEqual(0, sheet.NumMergedRegions, "there should be no merged regions left!");

            //an, Add, Remove, Get(0) would null pointer
            sheet.AddMergedRegion(region);
            Assert.AreEqual(1, sheet.NumMergedRegions, "there should now be one merged region!");
            sheet.RemoveMergedRegion(0);
            Assert.AreEqual(0, sheet.NumMergedRegions, "there should now be zero merged regions!");
            //add it again!
            region.LastRow = (/*setter*/4);

            sheet.AddMergedRegion(region);
            Assert.AreEqual(1, sheet.NumMergedRegions, "there should now be one merged region!");

            //should exist now!
            Assert.IsTrue(1 <= sheet.NumMergedRegions, "there isn't more than one merged region in there");
            region = sheet.GetMergedRegion(0);
            Assert.AreEqual(4, region.LastRow, "the merged row to doesnt match the one we Put in ");

            wb.Close();
        }

        /**
         * Remove multiple merged regions
         */
        [Test]
        public void RemoveMergedRegions()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            Dictionary<int, CellRangeAddress> mergedRegions = new Dictionary<int, CellRangeAddress>();
            for (int r = 0; r < 10; r++)
            {
                CellRangeAddress region = new CellRangeAddress(r, r, 0, 1);
                mergedRegions.Add(r, region);
                sheet.AddMergedRegion(region);
            }

            assertCollectionEquals(mergedRegions.Values.ToList(), sheet.MergedRegions);

            IList<int> removed = new List<int> { 0, 2, 3, 6, 8 };
            foreach (int rx in removed)
                mergedRegions.Remove(rx);
            sheet.RemoveMergedRegions(removed);
            assertCollectionEquals(mergedRegions.Values.ToList(), sheet.MergedRegions);

            wb.Close();
        }

        private static void assertCollectionEquals<T>(IList<T> expected, IList<T> actual)
        {
            ISet<T> e = new HashSet<T>(expected);
            ISet<T> a = new HashSet<T>(actual);
            //Assert.AreEqual(e, a);
            CollectionAssert.AreEquivalent(expected, actual);
        }


        [Test]
        public virtual void ShiftMerged()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICreationHelper factory = wb.GetCreationHelper();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(factory.CreateRichTextString("first row, first cell"));

            row = sheet.CreateRow(1);
            cell = row.CreateCell(1);
            cell.SetCellValue(factory.CreateRichTextString("second row, second cell"));

            CellRangeAddress region = CellRangeAddress.ValueOf("A2:B2");
            sheet.AddMergedRegion(region);

            sheet.ShiftRows(1, 1, 1);

            region = sheet.GetMergedRegion(0);
            CellRangeAddress expectedRegion = CellRangeAddress.ValueOf("A3:B3");
            Assert.AreEqual(expectedRegion, region, "Merged region not Moved over to row 2");

            wb.Close();
        }

        /**
 * bug 58885: checking for overlapping merged regions when
 * adding a merged region is safe, but runs in O(n).
 * the check for merged regions when adding a merged region
 * can be skipped (unsafe) and run in O(1).
 */
        [Test]
        public void AddMergedRegionUnsafe()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            CellRangeAddress region1 = CellRangeAddress.ValueOf("A1:B2");
            CellRangeAddress region2 = CellRangeAddress.ValueOf("B2:C3");
            CellRangeAddress region3 = CellRangeAddress.ValueOf("C3:D4");
            CellRangeAddress region4 = CellRangeAddress.ValueOf("J10:K11");
            Assume.That(region1.Intersects(region2));
            Assume.That(region2.Intersects(region3));

            sh.AddMergedRegionUnsafe(region1);
            Assert.IsTrue(sh.MergedRegions.Contains(region1));
            // adding a duplicate or overlapping merged region should not
            // raise an exception with the unsafe version of addMergedRegion.

            sh.AddMergedRegionUnsafe(region2);
            // the safe version of addMergedRegion should throw when trying to add a merged region that overlaps an existing region
            Assert.IsTrue(sh.MergedRegions.Contains(region2));
            try
            {
                sh.AddMergedRegion(region3);
                Assert.Fail("Expected InvalidOperationException. region3 overlaps already added merged region2.");
            }
            catch (InvalidOperationException)
            {
                // expected
                Assert.IsFalse(sh.MergedRegions.Contains(region3));
            }
            // addMergedRegion should not re-validate previously-added merged regions
            sh.AddMergedRegion(region4);
            // validation methods should detect a problem with previously added merged regions (runs in O(n^2) time)
            try
            {
                sh.ValidateMergedRegions();
                Assert.Fail("Expected validation to Assert.Fail. Sheet contains merged regions A1:B2 and B2:C3, which overlap at B2.");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
            wb.Close();
        }

        /**
         * Tests the display of gridlines, formulas, and rowcolheadings.
         * @author Shawn Laubach (slaubach at apache dot org)
         */
        [Test]
        public void TestDisplayOptions()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();

            Assert.AreEqual(sheet.DisplayGridlines, true);
            Assert.AreEqual(sheet.DisplayRowColHeadings, true);
            Assert.AreEqual(sheet.DisplayFormulas, false);
            Assert.AreEqual(sheet.DisplayZeros, true);

            sheet.DisplayGridlines = (/*setter*/false);
            sheet.DisplayRowColHeadings = (/*setter*/false);
            sheet.DisplayFormulas = (/*setter*/true);
            sheet.DisplayZeros = (/*setter*/false);

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = wb2.GetSheetAt(0);

            Assert.AreEqual(sheet.DisplayGridlines, false);
            Assert.AreEqual(sheet.DisplayRowColHeadings, false);
            Assert.AreEqual(sheet.DisplayFormulas, true);
            Assert.AreEqual(sheet.DisplayZeros, false);

            wb2.Close();
        }
        [Test]
        public void TestColumnWidth()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();

            //default column width measured in characters
            sheet.DefaultColumnWidth = (/*setter*/10);
            Assert.AreEqual(10, sheet.DefaultColumnWidth);
            //columns A-C have default width
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth(0));
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth(1));
            Assert.AreEqual(256 * 10, sheet.GetColumnWidth(2));

            //set custom width for D-F
            for (char i = 'D'; i <= 'F'; i++)
            {
                //Sheet#setColumnWidth accepts the width in units of 1/256th of a character width
                int w = 256 * 12;
                sheet.SetColumnWidth(/*setter*/i, w);
                Assert.AreEqual(w, sheet.GetColumnWidth(i));
            }
            //reset the default column width, columns A-C Change, D-F still have custom width
            sheet.DefaultColumnWidth = (/*setter*/20);
            Assert.AreEqual(20, sheet.DefaultColumnWidth);
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(0));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(1));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(2));
            for (char i = 'D'; i <= 'F'; i++)
            {
                int w = 256 * 12;
                Assert.AreEqual(w, sheet.GetColumnWidth(i));
            }

            // check for 16-bit signed/unsigned error:
            sheet.SetColumnWidth(/*setter*/10, 40000);
            Assert.AreEqual(40000, sheet.GetColumnWidth(10));

            //The maximum column width for an individual cell is 255 characters
            try
            {
                sheet.SetColumnWidth(/*setter*/9, 256 * 256);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("The maximum column width for an individual cell is 255 characters.", e.Message);
            }

            //serialize and read again
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);
            Assert.AreEqual(20, sheet.DefaultColumnWidth);
            //columns A-C have default width
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(0));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(1));
            Assert.AreEqual(256 * 20, sheet.GetColumnWidth(2));
            //columns D-F have custom width
            for (char i = 'D'; i <= 'F'; i++)
            {
                short w = (256 * 12);
                Assert.AreEqual(w, sheet.GetColumnWidth(i));
            }
            Assert.AreEqual(40000, sheet.GetColumnWidth(10));

            wb2.Close();
        }
        [Test]
        public void TestDefaultRowHeight()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            sheet.DefaultRowHeightInPoints = (/*setter*/15);
            Assert.AreEqual((short)300, sheet.DefaultRowHeight);
            Assert.AreEqual(15.0F, sheet.DefaultRowHeightInPoints, 0.01F);

            IRow row = sheet.CreateRow(1);
            // new row inherits  default height from the sheet
            Assert.AreEqual(sheet.DefaultRowHeight, row.Height);

            // Set a new default row height in twips and Test Getting the value in points
            sheet.DefaultRowHeight = (/*setter*/(short)360);
            Assert.AreEqual(18.0f, sheet.DefaultRowHeightInPoints, 0.01F);
            Assert.AreEqual((short)360, sheet.DefaultRowHeight);

            // Test that defaultRowHeight is a tRuncated short: E.G. 360inPoints -> 18; 361inPoints -> 18
            sheet.DefaultRowHeight = (/*setter*/(short)361);
            Assert.AreEqual((float)361 / 20, sheet.DefaultRowHeightInPoints, 0.01F);
            Assert.AreEqual((short)361, sheet.DefaultRowHeight);

            // Set a new default row height in points and Test Getting the value in twips
            sheet.DefaultRowHeightInPoints = (/*setter*/17.5f);
            Assert.AreEqual(17.5f, sheet.DefaultRowHeightInPoints, 0.01F);
            Assert.AreEqual((short)(17.5f * 20), sheet.DefaultRowHeight);

            workbook.Close();
        }

        /** cell with formula becomes null on cloning a sheet*/
        [Test]
        public virtual void Bug35084()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet("Sheet1");
            IRow r = s.CreateRow(0);
            r.CreateCell(0).SetCellValue(1);
            r.CreateCell(1).CellFormula = (/*setter*/"A1*2");
            ISheet s1 = wb.CloneSheet(0);
            r = s1.GetRow(0);
            Assert.AreEqual(r.GetCell(0).NumericCellValue, 1, 0, "double"); // sanity check
            Assert.IsNotNull(r.GetCell(1));
            Assert.AreEqual(r.GetCell(1).CellFormula, "A1*2", "formula");

            wb.Close();
        }

        /** Test that new default column styles Get applied */
        [Test]
        public virtual void DefaultColumnStyle()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICellStyle style = wb.CreateCellStyle();
            ISheet sheet = wb.CreateSheet();
            sheet.SetDefaultColumnStyle(/*setter*/0, style);
            Assert.IsNotNull(sheet.GetColumnStyle(0));
            Assert.AreEqual(style.Index, sheet.GetColumnStyle(0).Index);

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            ICellStyle style2 = cell.CellStyle;
            Assert.IsNotNull(style2);
            Assert.AreEqual(style.Index, style2.Index, "style should match");

            wb.Close();
        }
        [Test]
        public void TestOutlineProperties()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();

            ISheet sheet = wb1.CreateSheet();

            //TODO defaults are different in HSSF and XSSF
            //Assert.IsTrue(sheet.RowSumsBelow);
            //Assert.IsTrue(sheet.RowSumsRight);

            sheet.RowSumsBelow = (/*setter*/false);
            sheet.RowSumsRight = (/*setter*/false);

            Assert.IsFalse(sheet.RowSumsBelow);
            Assert.IsFalse(sheet.RowSumsRight);

            sheet.RowSumsBelow = (/*setter*/true);
            sheet.RowSumsRight = (/*setter*/true);

            Assert.IsTrue(sheet.RowSumsBelow);
            Assert.IsTrue(sheet.RowSumsRight);

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            sheet = wb2.GetSheetAt(0);
            Assert.IsTrue(sheet.RowSumsBelow);
            Assert.IsTrue(sheet.RowSumsRight);

            wb2.Close();
        }

        /**
         * Test basic display and print properties
         */
        [Test]
        public void TestSheetProperties()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            Assert.IsFalse(sheet.HorizontallyCenter);
            sheet.HorizontallyCenter = (/*setter*/true);
            Assert.IsTrue(sheet.HorizontallyCenter);
            sheet.HorizontallyCenter = (/*setter*/false);
            Assert.IsFalse(sheet.HorizontallyCenter);

            Assert.IsFalse(sheet.VerticallyCenter);
            sheet.VerticallyCenter = (/*setter*/true);
            Assert.IsTrue(sheet.VerticallyCenter);
            sheet.VerticallyCenter = (/*setter*/false);
            Assert.IsFalse(sheet.VerticallyCenter);

            Assert.IsFalse(sheet.IsPrintGridlines);
            sheet.IsPrintGridlines = (/*setter*/true);
            Assert.IsTrue(sheet.IsPrintGridlines);

            Assert.IsFalse(sheet.IsPrintRowAndColumnHeadings);
            sheet.IsPrintRowAndColumnHeadings = (true);
            Assert.IsTrue(sheet.IsPrintRowAndColumnHeadings);

            Assert.IsFalse(sheet.DisplayFormulas);
            sheet.DisplayFormulas = (/*setter*/true);
            Assert.IsTrue(sheet.DisplayFormulas);

            Assert.IsTrue(sheet.DisplayGridlines);
            sheet.DisplayGridlines = (/*setter*/false);
            Assert.IsFalse(sheet.DisplayGridlines);

            //TODO: default "guts" is different in HSSF and XSSF
            //Assert.IsTrue(sheet.DisplayGuts);
            sheet.DisplayGuts = (/*setter*/false);
            Assert.IsFalse(sheet.DisplayGuts);

            Assert.IsTrue(sheet.DisplayRowColHeadings);
            sheet.DisplayRowColHeadings = (/*setter*/false);
            Assert.IsFalse(sheet.DisplayRowColHeadings);

            //TODO: default "autobreaks" is different in HSSF and XSSF
            //Assert.IsTrue(sheet.Autobreaks);
            sheet.Autobreaks = (/*setter*/false);
            Assert.IsFalse(sheet.Autobreaks);

            Assert.IsFalse(sheet.ScenarioProtect);

            //TODO: default "fit-to-page" is different in HSSF and XSSF
            //Assert.IsFalse(sheet.FitToPage);
            sheet.FitToPage = (/*setter*/true);
            Assert.IsTrue(sheet.FitToPage);
            sheet.FitToPage = (/*setter*/false);
            Assert.IsFalse(sheet.FitToPage);

            wb.Close();
        }
        public void BaseTestGetSetMargin(double[] defaultMargins)
        {
            double marginLeft = defaultMargins[0];
            double marginRight = defaultMargins[1];
            double marginTop = defaultMargins[2];
            double marginBottom = defaultMargins[3];
            double marginHeader = defaultMargins[4];
            double marginFooter = defaultMargins[5];

            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet 1");
            Assert.AreEqual(marginLeft, sheet.GetMargin(MarginType.LeftMargin), 0.0);
            sheet.SetMargin(MarginType.LeftMargin, 10.0);
            //left margin is custom, all others are default
            Assert.AreEqual(10.0, sheet.GetMargin(MarginType.LeftMargin), 0.0);
            Assert.AreEqual(marginRight, sheet.GetMargin(MarginType.RightMargin), 0.0);
            Assert.AreEqual(marginTop, sheet.GetMargin(MarginType.TopMargin), 0.0);
            Assert.AreEqual(marginBottom, sheet.GetMargin(MarginType.BottomMargin), 0.0);
            sheet.SetMargin(/*setter*/MarginType.RightMargin, 11.0);
            Assert.AreEqual(11.0, sheet.GetMargin(MarginType.RightMargin), 0.0);
            sheet.SetMargin(/*setter*/MarginType.TopMargin, 12.0);
            Assert.AreEqual(12.0, sheet.GetMargin(MarginType.TopMargin), 0.0);
            sheet.SetMargin(/*setter*/MarginType.BottomMargin, 13.0);
            Assert.AreEqual(13.0, sheet.GetMargin(MarginType.BottomMargin), 0.0);

            sheet.SetMargin(MarginType.FooterMargin, 5.6);
            Assert.AreEqual(5.6, sheet.GetMargin(MarginType.FooterMargin), 0.0);
            sheet.SetMargin(MarginType.HeaderMargin, 11.5);
            Assert.AreEqual(11.5, sheet.GetMargin(MarginType.HeaderMargin), 0.0);

            // incorrect margin constant
            try
            {
                sheet.SetMargin((MarginType)65, 15);
                Assert.Fail("Expected exception");
            }
            catch (InvalidOperationException e)
            {
                Assert.AreEqual("Unknown margin constant:  65", e.Message);
            }

            workbook.Close();
        }
        [Test]
        public void TestRowBreaks()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            //Sheet#getRowBreaks() returns an empty array if no row breaks are defined
            Assert.IsNotNull(sheet.RowBreaks);
            Assert.AreEqual(0, sheet.RowBreaks.Length);

            sheet.SetRowBreak(1);
            Assert.AreEqual(1, sheet.RowBreaks.Length);
            sheet.SetRowBreak(15);
            Assert.AreEqual(2, sheet.RowBreaks.Length);
            Assert.AreEqual(1, sheet.RowBreaks[0]);
            Assert.AreEqual(15, sheet.RowBreaks[1]);
            sheet.SetRowBreak(1);
            Assert.AreEqual(2, sheet.RowBreaks.Length);
            Assert.IsTrue(sheet.IsRowBroken(1));
            Assert.IsTrue(sheet.IsRowBroken(15));

            //now remove the Created breaks
            sheet.RemoveRowBreak(1);
            Assert.AreEqual(1, sheet.RowBreaks.Length);
            sheet.RemoveRowBreak(15);
            Assert.AreEqual(0, sheet.RowBreaks.Length);

            Assert.IsFalse(sheet.IsRowBroken(1));
            Assert.IsFalse(sheet.IsRowBroken(15));

            workbook.Close();
        }
        [Test]
        public void TestColumnBreaks()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            Assert.IsNotNull(sheet.ColumnBreaks);
            Assert.AreEqual(0, sheet.ColumnBreaks.Length);

            Assert.IsFalse(sheet.IsColumnBroken(0));

            sheet.SetColumnBreak(11);
            Assert.IsNotNull(sheet.ColumnBreaks);
            Assert.AreEqual(11, sheet.ColumnBreaks[0]);
            sheet.SetColumnBreak(12);
            Assert.AreEqual(2, sheet.ColumnBreaks.Length);
            Assert.IsTrue(sheet.IsColumnBroken(11));
            Assert.IsTrue(sheet.IsColumnBroken(12));

            sheet.RemoveColumnBreak((short)11);
            Assert.AreEqual(1, sheet.ColumnBreaks.Length);
            sheet.RemoveColumnBreak((short)15); //remove non-existing
            Assert.AreEqual(1, sheet.ColumnBreaks.Length);
            sheet.RemoveColumnBreak((short)12);
            Assert.AreEqual(0, sheet.ColumnBreaks.Length);

            Assert.IsFalse(sheet.IsColumnBroken(11));
            Assert.IsFalse(sheet.IsColumnBroken(12));

            workbook.Close();
        }
        [Test]
        public void TestGetFirstLastRowNum()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet 1");
            sheet.CreateRow(9);
            sheet.CreateRow(0);
            sheet.CreateRow(1);
            Assert.AreEqual(0, sheet.FirstRowNum);
            Assert.AreEqual(9, sheet.LastRowNum);

            workbook.Close();
        }
        [Test]
        public void TestGetFooter()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet 1");
            Assert.IsNotNull(sheet.Footer);
            sheet.Footer.Center = (/*setter*/"test center footer");
            Assert.AreEqual("test center footer", sheet.Footer.Center);

            workbook.Close();
        }
        [Test]
        public void TestGetSetColumnHidden()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet 1");
            sheet.SetColumnHidden(2, true);
            Assert.IsTrue(sheet.IsColumnHidden(2));

            workbook.Close();
        }
        [Test]
        public void TestProtectSheet()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            Assert.IsFalse(sheet.Protect);
            sheet.ProtectSheet("Test");
            Assert.IsTrue(sheet.Protect);
            sheet.ProtectSheet(null);
            Assert.IsFalse(sheet.Protect);

            wb.Close();

        }
        [Test]
        public void TestCreateFreezePane()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            // create a workbook
            ISheet sheet = wb.CreateSheet();
            Assert.IsNull(sheet.PaneInformation);
            sheet.CreateFreezePane(0, 0);
            // still null
            Assert.IsNull(sheet.PaneInformation);

            sheet.CreateFreezePane(2, 3);

            PaneInformation info = sheet.PaneInformation;


            Assert.AreEqual(PaneInformation.PANE_LOWER_RIGHT, info.ActivePane);
            Assert.AreEqual(3, info.HorizontalSplitPosition);
            Assert.AreEqual(3, info.HorizontalSplitTopRow);
            Assert.AreEqual(2, info.VerticalSplitLeftColumn);
            Assert.AreEqual(2, info.VerticalSplitPosition);

            sheet.CreateFreezePane(0, 0);
            // If both colSplit and rowSplit are zero then the existing freeze pane is Removed
            Assert.IsNull(sheet.PaneInformation);

            sheet.CreateFreezePane(0, 3);

            info = sheet.PaneInformation;

            Assert.AreEqual(PaneInformation.PANE_LOWER_LEFT, info.ActivePane);
            Assert.AreEqual(3, info.HorizontalSplitPosition);
            Assert.AreEqual(3, info.HorizontalSplitTopRow);
            Assert.AreEqual(0, info.VerticalSplitLeftColumn);
            Assert.AreEqual(0, info.VerticalSplitPosition);

            sheet.CreateFreezePane(3, 0);

            info = sheet.PaneInformation;

            Assert.AreEqual(PaneInformation.PANE_UPPER_RIGHT, info.ActivePane);
            Assert.AreEqual(0, info.HorizontalSplitPosition);
            Assert.AreEqual(0, info.HorizontalSplitTopRow);
            Assert.AreEqual(3, info.VerticalSplitLeftColumn);
            Assert.AreEqual(3, info.VerticalSplitPosition);

            sheet.CreateFreezePane(0, 0);
            // If both colSplit and rowSplit are zero then the existing freeze pane is Removed
            Assert.IsNull(sheet.PaneInformation);

            wb.Close();
        }
        [Test]
        public void TestGetRepeatingRowsAndColumns()
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(
                "RepeatingRowsCols."
                + _testDataProvider.StandardFileNameExtension);

            CheckRepeatingRowsAndColumns(wb.GetSheetAt(0), null, null);
            CheckRepeatingRowsAndColumns(wb.GetSheetAt(1), "1:1", null);
            CheckRepeatingRowsAndColumns(wb.GetSheetAt(2), null, "A:A");
            CheckRepeatingRowsAndColumns(wb.GetSheetAt(3), "2:3", "A:B");

            wb.Close();
        }

        [Test]
        public void TestSetRepeatingRowsAndColumnsBug47294()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            sheet1.RepeatingRows = (CellRangeAddress.ValueOf("1:4"));
            Assert.AreEqual("1:4", sheet1.RepeatingRows.FormatAsString());

            //must handle sheets with quotas, see Bugzilla #47294
            ISheet sheet2 = wb.CreateSheet("My' Sheet");
            sheet2.RepeatingRows = (CellRangeAddress.ValueOf("1:4"));
            Assert.AreEqual("1:4", sheet2.RepeatingRows.FormatAsString());

            wb.Close();
        }
        [Test]
        public void TestSetRepeatingRowsAndColumns()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb1.CreateSheet("Sheet1");
            ISheet sheet2 = wb1.CreateSheet("Sheet2");
            ISheet sheet3 = wb1.CreateSheet("Sheet3");

            CheckRepeatingRowsAndColumns(sheet1, null, null);

            sheet1.RepeatingRows = (CellRangeAddress.ValueOf("4:5"));
            sheet2.RepeatingColumns = (CellRangeAddress.ValueOf("A:C"));
            sheet3.RepeatingRows = (CellRangeAddress.ValueOf("1:4"));
            sheet3.RepeatingColumns = (CellRangeAddress.ValueOf("A:A"));

            CheckRepeatingRowsAndColumns(sheet1, "4:5", null);
            CheckRepeatingRowsAndColumns(sheet2, null, "A:C");
            CheckRepeatingRowsAndColumns(sheet3, "1:4", "A:A");

            // write out, read back, and test refrain...
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet1 = wb2.GetSheetAt(0);
            sheet2 = wb2.GetSheetAt(1);
            sheet3 = wb2.GetSheetAt(2);

            CheckRepeatingRowsAndColumns(sheet1, "4:5", null);
            CheckRepeatingRowsAndColumns(sheet2, null, "A:C");
            CheckRepeatingRowsAndColumns(sheet3, "1:4", "A:A");

            // check removing repeating rows and columns       
            sheet3.RepeatingRows = (null);
            CheckRepeatingRowsAndColumns(sheet3, null, "A:A");

            sheet3.RepeatingColumns = (null);
            CheckRepeatingRowsAndColumns(sheet3, null, null);

            wb2.Close();
        }

        private void CheckRepeatingRowsAndColumns(
            ISheet s, String expectedRows, String expectedCols)
        {
            if (expectedRows == null)
            {
                Assert.IsNull(s.RepeatingRows);
            }
            else
            {
                Assert.AreEqual(expectedRows, s.RepeatingRows.FormatAsString());
            }
            if (expectedCols == null)
            {
                Assert.IsNull(s.RepeatingColumns);
            }
            else
            {
                Assert.AreEqual(expectedCols, s.RepeatingColumns.FormatAsString());
            }
        }
        [Test]
        public void TestBaseZoom()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            // here we can only verify that setting some zoom values works, range-checking is different between the implementations
            sheet.SetZoom(75);

            wb.Close();
        }

        [Test]
        public void TestBaseShowInPane()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            sheet.ShowInPane(2, 3);

            wb.Close();
        }

        [Test]
        public void TestBug55723()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            CellRangeAddress range = CellRangeAddress.ValueOf("A:B");
            IAutoFilter filter = sheet.SetAutoFilter(range);
            Assert.IsNotNull(filter);
            // there seems to be currently no generic way to check the Setting...

            range = CellRangeAddress.ValueOf("B:C");
            filter = sheet.SetAutoFilter(range);
            Assert.IsNotNull(filter);
            // there seems to be currently no generic way to check the Setting...

            wb.Close();
        }

        [Test]
        public void TestBug55723_Rows()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            CellRangeAddress range = CellRangeAddress.ValueOf("A4:B55000");
            IAutoFilter filter = sheet.SetAutoFilter(range);
            Assert.IsNotNull(filter);

            wb.Close();
        }


        [Test]
        public void TestBug55723d_RowsOver65k()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            CellRangeAddress range = CellRangeAddress.ValueOf("A4:B75000");
            IAutoFilter filter = sheet.SetAutoFilter(range);
            Assert.IsNotNull(filter);

            wb.Close();
        }

        /**
     * XSSFSheet autoSizeColumn() on empty RichTextString fails
     * @
     */
        [Test]
        public void Bug48325()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("Test");
            TrackColumnsForAutoSizingIfSXSSF(sheet);
            ICreationHelper factory = wb.GetCreationHelper();

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            IFont font = wb.CreateFont();
            IRichTextString rts = factory.CreateRichTextString("");
            rts.ApplyFont(font);
            cell.SetCellValue(rts);

            sheet.AutoSizeColumn(0);

            Assert.IsNotNull(_testDataProvider.WriteOutAndReadBack(wb));

            wb.Close();
        }

        [Test]
        public virtual void GetCellComment()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IDrawing dg = sheet.CreateDrawingPatriarch();
            IComment comment = dg.CreateCellComment(workbook.GetCreationHelper().CreateClientAnchor());
            ICell cell = sheet.CreateRow(9).CreateCell(2);
            comment.Author = (/*setter*/"test C10 author");
            cell.CellComment = (/*setter*/comment);

            CellAddress ref1 = new CellAddress(9, 2);
            Assert.IsNotNull(sheet.GetCellComment(ref1));
            Assert.AreEqual("test C10 author", sheet.GetCellComment(ref1).Author);

            Assert.IsNotNull(_testDataProvider.WriteOutAndReadBack(workbook));

            workbook.Close();
        }

        [Test]
        public void GetCellComments()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("TEST");

            // a sheet with no cell comments should return an empty map (not null or raise NPE).
            //assertEquals(Collections.emptyMap(), sheet.getCellComments());
            Assert.AreEqual(0, sheet.GetCellComments().Count);

            IDrawing dg = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = workbook.GetCreationHelper().CreateClientAnchor();

            int nRows = 5;
            int nCols = 6;

            for (int r = 0; r < nRows; r++)
            {
                sheet.CreateRow(r);
                // Create columns in reverse order
                for (int c = nCols - 1; c >= 0; c--)
                {
                    // When the comment box is visible, have it show in a 1x3 space
                    anchor.Col1 = c;
                    anchor.Col2 = c;
                    anchor.Row1 = r;
                    anchor.Row2 = r;

                    // Create the comment and set the text-author
                    IComment comment = dg.CreateCellComment(anchor);

                    ICell cell = sheet.GetRow(r).CreateCell(c);
                    comment.Author = ("Author " + r);
                    IRichTextString text = workbook.GetCreationHelper().CreateRichTextString("Test comment at row=" + r + ", column=" + c);
                    comment.String = (text);

                    // Assign the comment to the cell
                    cell.CellComment = (comment);
                }
            }

            IWorkbook wb = _testDataProvider.WriteOutAndReadBack(workbook);
            ISheet sh = wb.GetSheet("TEST");
            Dictionary<CellAddress, IComment> cellComments = sh.GetCellComments();
            Assert.AreEqual(nRows * nCols, cellComments.Count);

            foreach (KeyValuePair<CellAddress, IComment> e in cellComments)
            {
                CellAddress ref1 = e.Key;
                IComment aComment = e.Value;
                Assert.AreEqual("Author " + ref1.Row, aComment.Author);
                String text = "Test comment at row=" + ref1.Row + ", column=" + ref1.Column;
                Assert.AreEqual(text, aComment.String.String);
            }

            workbook.Close();
            wb.Close();
        }

        [Test]
        public void GetHyperlink()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            IHyperlink hyperlink = workbook.GetCreationHelper().CreateHyperlink(HyperlinkType.Url);
            hyperlink.Address = "https://poi.apache.org/";

            ISheet sheet = workbook.CreateSheet();
            ICell cell = sheet.CreateRow(5).CreateCell(1);

            Assert.AreEqual(0, sheet.GetHyperlinkList().Count, "list size before add");
            cell.Hyperlink = hyperlink;
            Assert.AreEqual(1, sheet.GetHyperlinkList().Count, "list size after add");

            Assert.AreEqual(hyperlink, sheet.GetHyperlinkList()[0], "list");
            Assert.AreEqual(hyperlink, sheet.GetHyperlink(5, 1), "row, col");
            CellAddress B6 = new CellAddress(5, 1);
            Assert.AreEqual(hyperlink, sheet.GetHyperlink(B6), "addr");
            Assert.AreEqual(null, sheet.GetHyperlink(CellAddress.A1), "no hyperlink at A1");

            workbook.Close();
        }

        [Test]
        public void RemoveAllHyperlinks() 
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            IHyperlink hyperlink = workbook.GetCreationHelper().CreateHyperlink(HyperlinkType.Url);
            hyperlink.Address = "https://poi.apache.org/";
            ISheet sheet = workbook.CreateSheet();
            ICell cell = sheet.CreateRow(5).CreateCell(1);
            cell.Hyperlink = hyperlink;

            Assert.AreEqual(1, workbook.GetSheetAt(0).GetHyperlinkList().Count);
            // Save a workbook with a hyperlink
            IWorkbook workbook2 = _testDataProvider.WriteOutAndReadBack(workbook);
            Assert.AreEqual(1, workbook2.GetSheetAt(0).GetHyperlinkList().Count);
        
            // Remove all hyperlinks from a saved workbook
            workbook2.GetSheetAt(0).GetRow(5).GetCell(1).RemoveHyperlink();
            Assert.AreEqual(0, workbook2.GetSheetAt(0).GetHyperlinkList().Count);
        
            // Verify that hyperlink was removed from workbook after writing out
            IWorkbook workbook3 = _testDataProvider.WriteOutAndReadBack(workbook2);
            Assert.AreEqual(0, workbook3.GetSheetAt(0).GetHyperlinkList().Count);
        }

        [Test]
        public void NewMergedRegionAt()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            CellRangeAddress region = CellRangeAddress.ValueOf("B2:D4");
            sheet.AddMergedRegion(region);
            Assert.AreEqual("B2:D4", sheet.GetMergedRegion(0).FormatAsString());
            Assert.AreEqual(1, sheet.NumMergedRegions);

            Assert.IsNotNull(_testDataProvider.WriteOutAndReadBack(workbook));

            workbook.Close();
        }

        [Test]
        public void ShowInPaneManyRowsBug55248()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet("Sheet 1");

            sheet.ShowInPane(0, 0);
            int i;
            for (i = ROW_COUNT / 2; i < ROW_COUNT; i++)
            {
                sheet.CreateRow(i);
                sheet.ShowInPane(i, 0);
                // this one fails: sheet.ShowInPane((short)i, 0);
            }

            i = 0;
            sheet.ShowInPane(i, i);

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();
            CheckRowCount(wb2);

            wb2.Close();
        }

        private void CheckRowCount(IWorkbook wb)
        {
            Assert.IsNotNull(wb);
            ISheet sh = wb.GetSheet("Sheet 1");
            Assert.IsNotNull(sh);
            Assert.AreEqual(ROW_COUNT - 1, sh.LastRowNum);
        }


        [Test]
        public void TestRightToLeft()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            Assert.IsFalse(sheet.IsRightToLeft);
            sheet.IsRightToLeft = (/*setter*/true);
            Assert.IsTrue(sheet.IsRightToLeft);
            sheet.IsRightToLeft = (/*setter*/false);
            Assert.IsFalse(sheet.IsRightToLeft);

            wb.Close();
        }

        [Test]
        public void TestNoMergedRegionsIsEmptyList()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            Assert.IsTrue(sheet.MergedRegions.Count ==0 );
            wb.Close();
        }
        [Test]
        public void MergedRangeValidateError()
        {
            var workbook = _testDataProvider.CreateWorkbook();
            var sheet = workbook.CreateSheet();
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 1));
            sheet.ValidateMergedRegions();
        }
        /**
         * Tests that the setAsActiveCell and getActiveCell function pairs work together
         */
        [Test]
        public void SetActiveCell()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();
            CellAddress B42 = new CellAddress("B42");

            // active cell behavior is undefined if not set.
            // HSSFSheet defaults to A1 active cell, while XSSFSheet defaults to null.
            if (sheet.ActiveCell != null && !sheet.ActiveCell.Equals(CellAddress.A1))
            {
                Assert.Fail("If not set, active cell should default to null or A1");
            }
            sheet.ActiveCell = (B42);
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            Assert.AreEqual(B42, sheet.ActiveCell);
            wb1.Close();
            wb2.Close();
        }

        [Test]
        public void TestAutoSizeDate()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet("Sheet1");
            IRow r = s.CreateRow(0);
            r.CreateCell(0).SetCellValue(1);
            r.CreateCell(1).SetCellValue(123456);

            // for the streaming-variant we need to enable autosize-tracking to make it work
            TrackColumnsForAutoSizingIfSXSSF(s);

            // Will be sized fairly small
            s.AutoSizeColumn((short) 0);
            s.AutoSizeColumn((short) 1);

            // Size ranges due to different fonts on different machines
            POITestCase.AssertBetween("Single number column width", (int) s.GetColumnWidth(0), 350, 570);
            POITestCase.AssertBetween("6 digit number column width", (int) s.GetColumnWidth(1), 1500, 2100);

            // Set a date format
            ICellStyle cs = wb.CreateCellStyle();
            IDataFormat f = wb.CreateDataFormat();
            cs.DataFormat = (/*setter*/f.GetFormat("yyyy-mm-dd MMMM hh:mm:ss"));
            r.GetCell(0).CellStyle = (/*setter*/cs);
            r.GetCell(1).CellStyle = (/*setter*/cs);

            Assert.IsTrue(DateUtil.IsCellDateFormatted(r.GetCell(0)));
            Assert.IsTrue(DateUtil.IsCellDateFormatted(r.GetCell(1)));

            // Should Get much bigger now
            s.AutoSizeColumn((short) 0);
            s.AutoSizeColumn((short) 1);

            POITestCase.AssertBetween("Date column width", (int) s.GetColumnWidth(0), 4750, 7300);
            POITestCase.AssertBetween("Date column width", (int) s.GetColumnWidth(1), 4750, 7300);

            wb.Close();
        }
        [Test]
        public void TestGetCells_SingleCell()
        {
            var wb1 = _testDataProvider.CreateWorkbook();
            var sheet = wb1.CreateSheet();
            var cellRanges= sheet.GetCells("A1");
            Assert.AreEqual(1, cellRanges.CountRanges());
        }

        [Test]
        public void TestGetCells_MultipleCellRange()
        {
            var wb1 = _testDataProvider.CreateWorkbook();
            var sheet = wb1.CreateSheet();
            var cellRanges = sheet.GetCells("A1:B2, D5:F7");
            Assert.AreEqual(2, cellRanges.CountRanges());
            Assert.AreEqual(4+9, cellRanges.NumberOfCells());
        }
        [Test]
        public void TestGetCells_SingleCellRange()
        {
            var wb1 = _testDataProvider.CreateWorkbook();
            var sheet = wb1.CreateSheet();
            var cellRanges = sheet.GetCells("Sheet1!B1:D3");
            Assert.AreEqual(1, cellRanges.CountRanges());
            Assert.AreEqual(1, cellRanges.GetCellRangeAddress(0).FirstColumn);
            Assert.AreEqual(3, cellRanges.GetCellRangeAddress(0).LastColumn);
            Assert.AreEqual(0, cellRanges.GetCellRangeAddress(0).FirstRow);
            Assert.AreEqual(2, cellRanges.GetCellRangeAddress(0).LastRow);
            Assert.AreEqual(9, cellRanges.GetCellRangeAddress(0).NumberOfCells);
        }        
    }

}