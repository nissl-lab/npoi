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
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.Streaming;
    using NUnit.Framework;


    /**
     * Tests the auto-sizing behaviour of {@link SXSSFSheet} when not all
     * rows fit into the memory window size etc.
     * 
     * @see Bug #57450 which reported the original mis-behaviour
     */
    [TestFixture]
    public class TestAutoSizeColumnTracker
    {

        private SXSSFSheet sheet;
        private SXSSFWorkbook workbook;
        private AutoSizeColumnTracker tracker;
        private static SortedSet<int> columns;
        static TestAutoSizeColumnTracker() {
            SortedSet<int> _columns = new SortedSet<int>();
            _columns.Add(0);
            _columns.Add(1);
            _columns.Add(3);
            columns = (_columns);
        }
        private static String SHORT_MESSAGE = "short";
        private static String LONG_MESSAGE = "This is a test of a long message! This is a test of a long message!";

        [SetUp]
        public void SetUpSheetAndWorkbook() {
            workbook = new SXSSFWorkbook();
            sheet = workbook.CreateSheet() as SXSSFSheet;
            tracker = new AutoSizeColumnTracker(sheet);
        }

        [TearDown]
        public void TearDownSheetAndWorkbook() {
            if (sheet != null) {
                sheet.Dispose();
            }
            if (workbook != null) {
                workbook.Close();
            }
        }

        [Test]
        public void trackAndUntrackColumn() {
            Assume.That(tracker.TrackedColumns.Count == 0);
            tracker.TrackColumn(0);
            ISet<int> expected = new HashSet<int>();
            expected.Add(0);
            Assert.AreEqual(expected, tracker.TrackedColumns);
            tracker.UntrackColumn(0);
            Assert.IsTrue(tracker.TrackedColumns.Count == 0);
        }

        [Test]
        public void trackAndUntrackColumns() {
            Assume.That(tracker.TrackedColumns.Count == 0);
            tracker.TrackColumns(columns);
            Assert.AreEqual(columns, tracker.TrackedColumns);
            tracker.UntrackColumn(3);
            tracker.UntrackColumn(0);
            tracker.UntrackColumn(1);
            Assert.IsTrue(tracker.TrackedColumns.Count == 0);
            tracker.TrackColumn(0);
            tracker.TrackColumns(columns);
            tracker.UntrackColumn(4);
            Assert.AreEqual(columns, tracker.TrackedColumns);
            tracker.UntrackColumns(columns);
            Assert.IsTrue(tracker.TrackedColumns.Count == 0);
        }

        [Test]
        public void trackAndUntrackAllColumns() {
            Assume.That(tracker.TrackedColumns.Count == 0);
            tracker.TrackAllColumns();
            Assert.IsTrue(tracker.TrackedColumns.Count == 0);

            IRow row = sheet.CreateRow(0);
            foreach (int column in columns) {
                row.CreateCell(column);
            }
            // implicitly track the columns
            tracker.UpdateColumnWidths(row);
            Assert.AreEqual(columns, tracker.TrackedColumns);

            tracker.UntrackAllColumns();
            Assert.IsTrue(tracker.TrackedColumns.Count == 0);
        }

        [Test]
        public void isColumnTracked() {
            Assert.IsFalse(tracker.IsColumnTracked(0));
            tracker.TrackColumn(0);
            Assert.IsTrue(tracker.IsColumnTracked(0));
            tracker.UntrackColumn(0);
            Assert.IsFalse(tracker.IsColumnTracked(0));
        }

        [Test]
        public void GetTrackedColumns() {
            Assume.That(tracker.TrackedColumns.Count == 0);

            foreach (int column in columns) {
                tracker.TrackColumn(column);
            }

            Assert.AreEqual(3, tracker.TrackedColumns.Count);
            Assert.AreEqual(columns, tracker.TrackedColumns);
        }

        [Test]
        public void isAllColumnsTracked() {
            Assert.IsFalse(tracker.IsAllColumnsTracked());
            tracker.TrackAllColumns();
            Assert.IsTrue(tracker.IsAllColumnsTracked());
            tracker.UntrackAllColumns();
            Assert.IsFalse(tracker.IsAllColumnsTracked());
        }

        [Test]
        [Platform("Win")]
        public void updateColumnWidths_and_getBestFitColumnWidth() {
            tracker.TrackAllColumns();
            IRow row1 = sheet.CreateRow(0);
            IRow row2 = sheet.CreateRow(1);
            // A1, B1, D1
            foreach (int column in columns) {
                row1.CreateCell(column).SetCellValue(LONG_MESSAGE);
                row2.CreateCell(column + 1).SetCellValue(SHORT_MESSAGE);
            }
            tracker.UpdateColumnWidths(row1);
            tracker.UpdateColumnWidths(row2);
            sheet.AddMergedRegion(CellRangeAddress.ValueOf("D1:E1"));

            assumeRequiredFontsAreInstalled(workbook, row1.GetCell(columns.GetEnumerator().Current));

            // Excel 2013 and LibreOffice 4.2.8.2 both treat columns with merged regions as blank
            /**    A     B    C      D   E
             * 1 LONG  LONG        LONGMERGE
             * 2       SHORT SHORT     SHORT
             */

            // measured in Excel 2013. Sizes may vary.
            int longMsgWidth = (int)(57.43 * 256);
            int shortMsgWidth = (int)(4.86 * 256);

            CheckColumnWidth(longMsgWidth, 0, true);
            CheckColumnWidth(longMsgWidth, 0, false);
            CheckColumnWidth(longMsgWidth, 1, true);
            CheckColumnWidth(longMsgWidth, 1, false);
            CheckColumnWidth(shortMsgWidth, 2, true);
            CheckColumnWidth(shortMsgWidth, 2, false);
            CheckColumnWidth(-1, 3, true);
            CheckColumnWidth(longMsgWidth, 3, false);
            CheckColumnWidth(shortMsgWidth, 4, true); //but is it really? shouldn't autosizing column E use "" from E1 and SHORT from E2?
            CheckColumnWidth(shortMsgWidth, 4, false);
        }

        private void CheckColumnWidth(int expectedWidth, int column, bool useMergedCells) {
            int bestFitWidth = tracker.GetBestFitColumnWidth(column, useMergedCells);
            if (bestFitWidth < 0 && expectedWidth < 0) return;
            double abs_error = Math.Abs(bestFitWidth - expectedWidth);
            double rel_error = abs_error / expectedWidth;
            if (rel_error > 0.25) {
                Assert.Fail("check column width: " +
                        rel_error + ", " + abs_error + ", " +
                        expectedWidth + ", " + bestFitWidth);
            }

        }

        private static void assumeRequiredFontsAreInstalled(IWorkbook workbook, ICell cell) {
            // autoSize will fail if required fonts are not installed, skip this test then
            IFont font = workbook.GetFontAt(cell.CellStyle.FontIndex);
            //System.out.Println(font.FontHeightInPoints);
            //System.out.Println(font.FontName);
            Assume.That(SheetUtil.CanComputeColumnWidth(font),
                "Cannot verify autoSizeColumn() because the necessary Fonts are not installed on this machine: " + font);
        }
    }
}
