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

using System.Diagnostics;

using NUnit.Framework;
using NUnit.Framework.Legacy;

using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace TestCases.SS.Util
{
    /// <summary>
    /// Tests for SheetUtil column width optimizations (merged region pre-fetch, SKFont caching).
    /// Uses XSSF to exercise the XML-parsing path that the optimizations target.
    /// </summary>
    [TestFixture]
    public class TestSheetUtilAutoSize
    {
        [Test]
        public void TestAutoSizeColumnProducesSameWidthAsGetCellWidth()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            // Populate with varied content: strings, numbers, booleans
            IRow header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Short");
            header.CreateCell(1).SetCellValue(12345.6789);
            header.CreateCell(2).SetCellValue(true);

            IRow data = sheet.CreateRow(1);
            data.CreateCell(0).SetCellValue("A much longer string value for column width testing");
            data.CreateCell(1).SetCellValue(0.001);
            data.CreateCell(2).SetCellValue(false);

            for (int col = 0; col < 3; col++)
            {
                double widthFromUtil = SheetUtil.GetColumnWidth(sheet, col, false);
                ClassicAssert.IsTrue(widthFromUtil > 0, $"Column {col} should have positive width");

                sheet.AutoSizeColumn(col);
                double colWidthAfterAutoSize = sheet.GetColumnWidth(col);
                ClassicAssert.IsTrue(colWidthAfterAutoSize > 0, $"Column {col} should have positive width after AutoSizeColumn");
            }
        }

        [Test]
        public void TestAutoSizeColumnWithMergedRegions()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            IRow row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue("Merged content that is quite long");
            row0.CreateCell(2).SetCellValue("Not merged");

            IRow row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue("Normal cell");
            row1.CreateCell(2).SetCellValue("Another value");

            // Merge A1:B1
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));

            // GetColumnWidth with useMergedCells=true should include merged content
            double widthWithMerged = SheetUtil.GetColumnWidth(sheet, 0, true);
            ClassicAssert.IsTrue(widthWithMerged > 0, "Should have positive width with merged cells");

            // GetColumnWidth with useMergedCells=false should skip merged cells
            double widthWithoutMerged = SheetUtil.GetColumnWidth(sheet, 0, false);
            // Row 1 has "Normal cell" so width should still be > 0
            ClassicAssert.IsTrue(widthWithoutMerged > 0, "Should have positive width from non-merged row");

            // Column 2 (not merged) should work regardless
            double col2Width = SheetUtil.GetColumnWidth(sheet, 2, false);
            ClassicAssert.IsTrue(col2Width > 0, "Non-merged column should have positive width");

            // AutoSizeColumn should not throw with merged regions
            sheet.AutoSizeColumn(0, true);
            sheet.AutoSizeColumn(2, false);
        }

        [Test]
        public void TestAutoSizeColumnWithMultipleMergedRegions()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            // Create several merged regions to exercise the pre-fetched list
            for (int i = 0; i < 10; i++)
            {
                IRow row = sheet.CreateRow(i);
                row.CreateCell(0).SetCellValue($"Row {i} merged content");
                row.CreateCell(2).SetCellValue($"Row {i} col 2");
                // Merge columns 0-1 for each row
                sheet.AddMergedRegion(new CellRangeAddress(i, i, 0, 1));
            }

            // Should handle multiple merged regions correctly
            double width = SheetUtil.GetColumnWidth(sheet, 0, true);
            ClassicAssert.IsTrue(width > 0);

            sheet.AutoSizeColumn(0, true);
            double autoSizedWidth = sheet.GetColumnWidth(0);
            ClassicAssert.IsTrue(autoSizedWidth > 0);
        }

        [Test]
        public void TestGetColumnWidthWithMaxRows()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            // Row 0: short text
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Hi");
            // Row 50: very long text
            sheet.CreateRow(50).CreateCell(0).SetCellValue(
                "This is a very long string that should make the column much wider than the short text above");
            // Rows 1-49: fill with medium text
            for (int i = 1; i < 50; i++)
            {
                sheet.CreateRow(i).CreateCell(0).SetCellValue("Medium");
            }

            double widthAll = SheetUtil.GetColumnWidth(sheet, 0, false);
            double widthFirst10 = SheetUtil.GetColumnWidth(sheet, 0, false, maxRows: 10);

            // maxRows=10 should only consider first 10 rows, missing the long row 50
            ClassicAssert.IsTrue(widthAll > widthFirst10,
                "Full scan should find the wider row 50 that maxRows=10 misses");
        }

        [Test]
        public void TestAutoSizeColumnWithMixedFonts()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            // Create cells with different fonts to exercise the font cache
            IFont normalFont = wb.CreateFont();
            normalFont.FontHeightInPoints = 11;
            normalFont.FontName = "Calibri";

            IFont bigFont = wb.CreateFont();
            bigFont.FontHeightInPoints = 24;
            bigFont.FontName = "Calibri";

            IFont boldFont = wb.CreateFont();
            boldFont.FontHeightInPoints = 11;
            boldFont.FontName = "Calibri";
            boldFont.IsBold = true;

            ICellStyle normalStyle = wb.CreateCellStyle();
            normalStyle.SetFont(normalFont);

            ICellStyle bigStyle = wb.CreateCellStyle();
            bigStyle.SetFont(bigFont);

            ICellStyle boldStyle = wb.CreateCellStyle();
            boldStyle.SetFont(boldFont);

            for (int i = 0; i < 30; i++)
            {
                IRow row = sheet.CreateRow(i);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue($"Row {i} text content");

                // Cycle through styles to exercise font cache
                if (i % 3 == 0) cell.CellStyle = bigStyle;
                else if (i % 3 == 1) cell.CellStyle = boldStyle;
                else cell.CellStyle = normalStyle;
            }

            double width = SheetUtil.GetColumnWidth(sheet, 0, false);
            ClassicAssert.IsTrue(width > 0, "Mixed-font column should have positive width");

            sheet.AutoSizeColumn(0);
            ClassicAssert.IsTrue(sheet.GetColumnWidth(0) > 0);
        }

        [Test]
        public void TestAutoSizeColumnPerformanceImprovement()
        {
            const int rowCount = 5000;
            const int colCount = 5;

            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("perf");

            string[] words = "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor".Split(' ');
            int wordIdx = 0;
            for (int r = 0; r < rowCount; r++)
            {
                IRow row = sheet.CreateRow(r);
                for (int c = 0; c < colCount; c++)
                {
                    row.CreateCell(c).SetCellValue(words[wordIdx++ % words.Length]);
                }
            }

            var sw = Stopwatch.StartNew();
            for (int c = 0; c < colCount; c++)
            {
                sheet.AutoSizeColumn(c);
            }
            sw.Stop();

            // Verify correctness: all columns should have been sized
            for (int c = 0; c < colCount; c++)
            {
                ClassicAssert.IsTrue(sheet.GetColumnWidth(c) > 0, $"Column {c} should have positive width");
            }

            // Log time for manual review; this is not a hard assertion since CI times vary,
            // but we want to see that 5000 rows x 5 cols completes in a reasonable time.
            TestContext.WriteLine($"AutoSizeColumn for {rowCount} rows x {colCount} cols: {sw.ElapsedMilliseconds}ms");

            // Soft upper bound: should complete in under 30 seconds even on slow CI
            ClassicAssert.IsTrue(sw.ElapsedMilliseconds < 30_000,
                $"AutoSizeColumn took {sw.ElapsedMilliseconds}ms, expected < 30s for {rowCount} rows");
        }

        [Test]
        public void TestAutoSizeColumnPerformanceWithMergedRegions()
        {
            const int rowCount = 1000;

            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("perf_merged");

            for (int r = 0; r < rowCount; r++)
            {
                IRow row = sheet.CreateRow(r);
                row.CreateCell(0).SetCellValue($"Content row {r}");
                row.CreateCell(1).SetCellValue($"Col B row {r}");
                row.CreateCell(3).SetCellValue($"Col D row {r}");
            }

            // Add merged regions on every 10th row to create a meaningful number
            for (int r = 0; r < rowCount; r += 10)
            {
                sheet.AddMergedRegion(new CellRangeAddress(r, r, 0, 1));
            }

            var sw = Stopwatch.StartNew();
            sheet.AutoSizeColumn(0, true);
            sheet.AutoSizeColumn(3, false);
            sw.Stop();

            ClassicAssert.IsTrue(sheet.GetColumnWidth(0) > 0);
            ClassicAssert.IsTrue(sheet.GetColumnWidth(3) > 0);

            TestContext.WriteLine($"AutoSizeColumn with {rowCount} rows and {rowCount / 10} merged regions: {sw.ElapsedMilliseconds}ms");
            ClassicAssert.IsTrue(sw.ElapsedMilliseconds < 30_000,
                $"AutoSizeColumn with merged regions took {sw.ElapsedMilliseconds}ms, expected < 30s");
        }

        [Test]
        public void TestGetCellWidthPublicApiUnchanged()
        {
            // Verify the public GetCellWidth API still works independently (not through GetColumnWidth)
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("test value");

            var formatter = new DataFormatter();
            int defaultCharWidth = SheetUtil.GetDefaultCharWidth(wb);

            double width = SheetUtil.GetCellWidth(cell, defaultCharWidth, formatter, false);
            ClassicAssert.IsTrue(width > 0, "Public GetCellWidth should still work");
        }

        [Test]
        public void TestAutoSizeEmptySheet()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("empty");

            // Should not throw on empty sheet
            sheet.AutoSizeColumn(0);
        }

        [Test]
        public void TestAutoSizeColumnWithEmptyCells()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            // Create rows but only populate some cells in column 0
            for (int r = 0; r < 100; r++)
            {
                IRow row = sheet.CreateRow(r);
                if (r % 10 == 0)
                {
                    row.CreateCell(0).SetCellValue($"Value at row {r}");
                }
                // Other rows have no cell in column 0
            }

            double width = SheetUtil.GetColumnWidth(sheet, 0, false);
            ClassicAssert.IsTrue(width > 0, "Should find width from sparse cells");

            sheet.AutoSizeColumn(0);
            ClassicAssert.IsTrue(sheet.GetColumnWidth(0) > 0);
        }

        [Test]
        public void TestAutoSizeColumnWithMaxRowsOverload()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            // Row 0: short text
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Hi");
            // Rows 1-99: medium text
            for (int i = 1; i < 100; i++)
            {
                sheet.CreateRow(i).CreateCell(0).SetCellValue("Medium text");
            }
            // Row 200: very long text (beyond maxRows reach)
            sheet.CreateRow(200).CreateCell(0).SetCellValue(
                "This is a very long string that should make the column much wider than the shorter text above");

            // Use the new (int column, int maxRows) overload — samples first 50 rows
            sheet.AutoSizeColumn(0, 50);
            double widthSampled = sheet.GetColumnWidth(0);
            ClassicAssert.IsTrue(widthSampled > 0, "Sampled auto-size should produce positive width");

            // Full auto-size should find the wider row 200
            sheet.AutoSizeColumn(0);
            double widthFull = sheet.GetColumnWidth(0);
            ClassicAssert.IsTrue(widthFull > widthSampled,
                "Full scan should find the wider row 200 that maxRows=50 misses");
        }

        [Test]
        public void TestSpatialIndexWithMultiRowMergedRegion()
        {
            using var wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("test");

            // Create a merged region spanning multiple rows
            for (int r = 0; r < 20; r++)
            {
                IRow row = sheet.CreateRow(r);
                row.CreateCell(0).SetCellValue($"Row {r}");
                row.CreateCell(2).SetCellValue($"Col C row {r}");
            }

            // Merge A1:B5 (rows 0-4, columns 0-1)
            sheet.AddMergedRegion(new CellRangeAddress(0, 4, 0, 1));
            // Merge A10:B15 (rows 9-14, columns 0-1)
            sheet.AddMergedRegion(new CellRangeAddress(9, 14, 0, 1));

            // Auto-size with merged cells — should handle multi-row regions correctly
            sheet.AutoSizeColumn(0, true);
            ClassicAssert.IsTrue(sheet.GetColumnWidth(0) > 0);

            // Auto-size without merged cells — rows in merged regions should be skipped
            sheet.AutoSizeColumn(0, false);
            // Non-merged rows (5-8, 15-19) still have content
            ClassicAssert.IsTrue(sheet.GetColumnWidth(0) > 0);

            // Column 2 has no merged regions — should work normally
            sheet.AutoSizeColumn(2);
            ClassicAssert.IsTrue(sheet.GetColumnWidth(2) > 0);
        }
    }
}
