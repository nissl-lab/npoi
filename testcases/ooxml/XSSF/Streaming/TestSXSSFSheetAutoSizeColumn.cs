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
namespace TestCases.XSSF.Streaming
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.Streaming;
    using NUnit.Framework;

    //Add a new test fixture, set useMergedCells with true value.
    [TestFixture]
    public class TestSXSSFSheetAutoSizeColumn_useMergedCells_true : TestSXSSFSheetAutoSizeColumn
    {
        public TestSXSSFSheetAutoSizeColumn_useMergedCells_true()
        {
            this.useMergedCells = true;
        }
    }

    /**
     * Tests the auto-sizing behaviour of {@link SXSSFSheet} when not all
     * rows fit into the memory window size etc.
     * 
     * @see Bug #57450 which reported the original mis-behaviour
     */
    [TestFixture]
    public class TestSXSSFSheetAutoSizeColumn
    {

        private static String SHORT_CELL_VALUE = "Ben";
        private static String LONG_CELL_VALUE = "B Be Ben Beni Benif Benify Benif Beni Ben Be B";

        // Approximative threshold to decide whether test is PASS or FAIL:
        //  shortCellValue ends up with approx column width 1_000 (on my machine),
        //  longCellValue ends up with approx. column width 10_000 (on my machine)
        //  so shortCellValue can be expected to be < 5000 for all fonts
        //  and longCellValue can be expected to be > 5000 for all fonts
        private static int COLUMN_WIDTH_THRESHOLD_BETWEEN_SHORT_AND_LONG = 4000;
        private static int MAX_COLUMN_WIDTH = 255 * 256;

        private static SortedSet<int> columns;
        static TestSXSSFSheetAutoSizeColumn()
        {
            SortedSet<int> _columns = new SortedSet<int>();
            _columns.Add(0);
            _columns.Add(1);
            _columns.Add(3);
            columns = (_columns);
        }


        private SXSSFSheet sheet;
        private SXSSFWorkbook workbook;
        
        public bool useMergedCells;

        //public static ICollection<Object[]> data()
        //{
        //    return Array.AsList(new Object[][] {
        //         {false},
        //         {true},
        //   });
        //}

        [TearDown]
        public void TearDownSheetAndWorkbook()
        {
            if (sheet != null)
            {
                sheet.Dispose();
            }
            if (workbook != null)
            {
                workbook.Close();
            }
        }

        [Test]
        public void Test_EmptySheet_NoException()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;
            sheet.TrackAllColumnsForAutoSizing();

            for (int i = 0; i < 10; i++)
            {
                sheet.AutoSizeColumn(i, useMergedCells);
            }
        }

        [Test]
        public void Test_WindowSizeDefault_AllRowsFitIntoWindowSize()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;
            sheet.TrackAllColumnsForAutoSizing();

            ICell cellRow0 = CreateRowWithCellValues(sheet, 0, LONG_CELL_VALUE);

            assumeRequiredFontsAreInstalled(workbook, cellRow0);

            CreateRowWithCellValues(sheet, 1, SHORT_CELL_VALUE);

            sheet.AutoSizeColumn(0, useMergedCells);

            assertColumnWidthStrictlyWithinRange(sheet.GetColumnWidth(0), COLUMN_WIDTH_THRESHOLD_BETWEEN_SHORT_AND_LONG, MAX_COLUMN_WIDTH);
        }

        [Test]
        public void Test_WindowSizeEqualsOne_ConsiderFlushedRows()
        {
            workbook = new SXSSFWorkbook(null, 1); // Window size 1 so only last row will be in memory
            sheet = workbook.CreateSheet() as SXSSFSheet;
            sheet.TrackAllColumnsForAutoSizing();

            ICell cellRow0 = CreateRowWithCellValues(sheet, 0, LONG_CELL_VALUE);

            assumeRequiredFontsAreInstalled(workbook, cellRow0);

            CreateRowWithCellValues(sheet, 1, SHORT_CELL_VALUE);

            sheet.AutoSizeColumn(0, useMergedCells);

            assertColumnWidthStrictlyWithinRange(sheet.GetColumnWidth(0), COLUMN_WIDTH_THRESHOLD_BETWEEN_SHORT_AND_LONG, MAX_COLUMN_WIDTH);
        }

        [Test]
        public void Test_WindowSizeEqualsOne_lastRowIsNotWidest()
        {
            workbook = new SXSSFWorkbook(null, 1); // Window size 1 so only last row will be in memory
            sheet = workbook.CreateSheet() as SXSSFSheet;
            sheet.TrackAllColumnsForAutoSizing();

            ICell cellRow0 = CreateRowWithCellValues(sheet, 0, LONG_CELL_VALUE);

            assumeRequiredFontsAreInstalled(workbook, cellRow0);

            CreateRowWithCellValues(sheet, 1, SHORT_CELL_VALUE);

            sheet.AutoSizeColumn(0, useMergedCells);

            assertColumnWidthStrictlyWithinRange(sheet.GetColumnWidth(0), COLUMN_WIDTH_THRESHOLD_BETWEEN_SHORT_AND_LONG, MAX_COLUMN_WIDTH);
        }

        [Test]
        public void Test_WindowSizeEqualsOne_lastRowIsWidest()
        {
            workbook = new SXSSFWorkbook(null, 1); // Window size 1 so only last row will be in memory
            sheet = workbook.CreateSheet() as SXSSFSheet;
            sheet.TrackAllColumnsForAutoSizing();

            ICell cellRow0 = CreateRowWithCellValues(sheet, 0, SHORT_CELL_VALUE);

            assumeRequiredFontsAreInstalled(workbook, cellRow0);

            CreateRowWithCellValues(sheet, 1, LONG_CELL_VALUE);

            sheet.AutoSizeColumn(0, useMergedCells);

            assertColumnWidthStrictlyWithinRange(sheet.GetColumnWidth(0), COLUMN_WIDTH_THRESHOLD_BETWEEN_SHORT_AND_LONG, MAX_COLUMN_WIDTH);
        }

        // fails only for useMergedCell=true
        [Test]
        public void Test_WindowSizeEqualsOne_flushedRowHasMergedCell()
        {
            workbook = new SXSSFWorkbook(null, 1); // Window size 1 so only last row will be in memory
            sheet = workbook.CreateSheet() as SXSSFSheet;
            sheet.TrackAllColumnsForAutoSizing();

            ICell a1 = CreateRowWithCellValues(sheet, 0, LONG_CELL_VALUE);

            assumeRequiredFontsAreInstalled(workbook, a1);
            sheet.AddMergedRegion(CellRangeAddress.ValueOf("A1:B1"));

            CreateRowWithCellValues(sheet, 1, SHORT_CELL_VALUE, SHORT_CELL_VALUE);

            /**
             *    A      B
             * 1 LONGMERGED
             * 2 SHORT SHORT
             */

            sheet.AutoSizeColumn(0, useMergedCells);
            sheet.AutoSizeColumn(1, useMergedCells);

            if (useMergedCells)
            {
                // Excel and LibreOffice behavior: ignore merged cells for auto-sizing.
                // POI behavior: evenly distribute the column width among the merged columns.
                //               each column must be auto-sized in order for the column widths
                //               to add up to the best fit width.
                int colspan = 2;
                int expectedWidth = (10000 + 1000) / colspan; //average of 1_000 and 10_000
                int minExpectedWidth = expectedWidth / 2;
                int maxExpectedWidth = expectedWidth * 3 / 2;
                assertColumnWidthStrictlyWithinRange(sheet.GetColumnWidth(0), minExpectedWidth, maxExpectedWidth); //short
            }
            else
            {
                assertColumnWidthStrictlyWithinRange(sheet.GetColumnWidth(0), COLUMN_WIDTH_THRESHOLD_BETWEEN_SHORT_AND_LONG, MAX_COLUMN_WIDTH); //long
            }
            assertColumnWidthStrictlyWithinRange(sheet.GetColumnWidth(1), 0, COLUMN_WIDTH_THRESHOLD_BETWEEN_SHORT_AND_LONG); //short
        }

        [Test]
        public void AutoSizeColumn_trackColumnForAutoSizing()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;
            sheet.TrackColumnForAutoSizing(0);

            SortedSet<int> expected = new SortedSet<int>();
            expected.Add(0);
            Assert.AreEqual(expected, sheet.TrackedColumnsForAutoSizing);

            sheet.AutoSizeColumn(0, useMergedCells);
            try
            {
                sheet.AutoSizeColumn(1, useMergedCells);
                Assert.Fail("Should not be able to auto-size an untracked column");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
        }


        [Test]
        public void AutoSizeColumn_trackColumnsForAutoSizing()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;

            sheet.TrackColumnsForAutoSizing(columns);
            SortedSet<int> sorted = new SortedSet<int>(columns);
            Assert.AreEqual(sorted, sheet.TrackedColumnsForAutoSizing);

            sheet.AutoSizeColumn(sorted.First(), useMergedCells);
            try
            {
                //assumeFalse(columns.Contains(5));
                Assume.That(!columns.Contains(5));
                sheet.AutoSizeColumn(5, useMergedCells);
                Assert.Fail("Should not be able to auto-size an untracked column");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
        }


        [Test]
        public void AutoSizeColumn_untrackColumnForAutoSizing()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;

            sheet.TrackColumnsForAutoSizing(columns);
            sheet.UntrackColumnForAutoSizing(columns.First());

            Assume.That(sheet.TrackedColumnsForAutoSizing.Contains(columns.Last()));
            sheet.AutoSizeColumn(columns.Last(), useMergedCells);
            try
            {
                Assume.That(!sheet.TrackedColumnsForAutoSizing.Contains(columns.First()));
                sheet.AutoSizeColumn(columns.First(), useMergedCells);
                Assert.Fail("Should not be able to auto-size an untracked column");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
        }

        [Test]
        public void AutoSizeColumn_untrackColumnsForAutoSizing()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;

            sheet.TrackColumnForAutoSizing(15);
            sheet.TrackColumnsForAutoSizing(columns);
            sheet.UntrackColumnsForAutoSizing(columns);

            Assume.That(sheet.TrackedColumnsForAutoSizing.Contains(15));
            sheet.AutoSizeColumn(15, useMergedCells);
            try
            {
                Assume.That(!sheet.TrackedColumnsForAutoSizing.Contains(columns.First()));
                sheet.AutoSizeColumn(columns.First(), useMergedCells);
                Assert.Fail("Should not be able to auto-size an untracked column");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
        }

        [Test]
        public void AutoSizeColumn_isColumnTrackedForAutoSizing()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;

            sheet.TrackColumnsForAutoSizing(columns);
            foreach (int column in columns)
            {
                Assert.IsTrue(sheet.IsColumnTrackedForAutoSizing(column));

                Assume.That(!columns.Contains(column + 10));
                Assert.IsFalse(sheet.IsColumnTrackedForAutoSizing(column + 10));
            }
        }

        [Test]
        public void AutoSizeColumn_trackAllColumns()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;

            sheet.TrackAllColumnsForAutoSizing();
            sheet.AutoSizeColumn(0, useMergedCells);

            sheet.UntrackAllColumnsForAutoSizing();
            try
            {
                sheet.AutoSizeColumn(0, useMergedCells);
                Assert.Fail("Should not be able to auto-size an implicitly untracked column");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
        }

        [Test]
        public void AutoSizeColumn_trackAllColumns_explicitUntrackColumn()
        {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;

            sheet.TrackColumnsForAutoSizing(columns);
            sheet.TrackAllColumnsForAutoSizing();

            sheet.UntrackColumnForAutoSizing(0);
            try
            {
                sheet.AutoSizeColumn(0, useMergedCells);
                Assert.Fail("Should not be able to auto-size an explicitly untracked column");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
        }


        private static void assumeRequiredFontsAreInstalled(IWorkbook workbook, ICell cell)
        {
            // autoSize will fail if required fonts are not installed, skip this test then
            IFont font = workbook.GetFontAt(cell.CellStyle.FontIndex);
            Assume.That(SheetUtil.CanComputeColumnWidth(font),
                "Cannot verify autoSizeColumn() because the necessary Fonts are not installed on this machine: " + font);
        }

        private static ICell CreateRowWithCellValues(ISheet sheet, int rowNumber, params string[] cellValues)
        {
            IRow row = sheet.CreateRow(rowNumber);
            int cellIndex = 0;
            ICell firstCell = null;
            foreach (String cellValue in cellValues)
            {
                ICell cell = row.CreateCell(cellIndex++);
                if (firstCell == null)
                {
                    firstCell = cell;
                }
                cell.SetCellValue(cellValue);
            }
            return firstCell;
        }

        private static void assertColumnWidthStrictlyWithinRange(double actualColumnWidth, int lowerBoundExclusive, int upperBoundExclusive)
        {
            Assert.IsTrue(actualColumnWidth > lowerBoundExclusive,
                "Expected a column width greater than " + lowerBoundExclusive + " but found " + actualColumnWidth);
            Assert.IsTrue(actualColumnWidth < upperBoundExclusive,
                "Expected column width less than " + upperBoundExclusive + " but found " + actualColumnWidth);

        }
    }
}

