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

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFSheetCopyTo
    {
        [Test]
        public void CopySheetToWorkbookShouldCopyFormulasOver()
        {
            XSSFWorkbook srcWorkbook = new XSSFWorkbook();
            XSSFSheet srcSheet = srcWorkbook.CreateSheet("Sheet1") as XSSFSheet;

            // Set some values
            IRow row1 = srcSheet.CreateRow((short)0);
            ICell cell = row1.CreateCell((short)0);
            cell.SetCellValue(1);
            ICell cell2 = row1.CreateCell((short)1);
            cell2.SetCellFormula("A1+1");
            XSSFWorkbook destWorkbook = new XSSFWorkbook();
            srcSheet.CopyTo(destWorkbook, srcSheet.SheetName, true, true);

            var destSheet = destWorkbook.GetSheet("Sheet1");
            ClassicAssert.NotNull(destSheet);

            ClassicAssert.AreEqual(1, destSheet.GetRow(0)?.GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual("A1+1", destSheet.GetRow(0)?.GetCell(1).CellFormula);

            destSheet.GetRow(0)?.GetCell(0).SetCellValue(10);
            var evaluator = destWorkbook.GetCreationHelper()
                                        .CreateFormulaEvaluator();

            var destCell = destSheet.GetRow(0)?.GetCell(1);
            evaluator.EvaluateFormulaCell(destCell);
            var destCellValue = evaluator.Evaluate(destCell);

            ClassicAssert.AreEqual(11, destCellValue.NumberValue);
        }

        [Test]
        public void CopySheetToWorkbookShouldCopyMergedRegionsOver()
        {
            XSSFWorkbook srcWorkbook = new XSSFWorkbook();
            XSSFSheet srcSheet = srcWorkbook.CreateSheet("Sheet1") as XSSFSheet;

            // Set some merged regions 
            srcSheet.AddMergedRegion(CellRangeAddress.ValueOf("A1:B4"));
            srcSheet.AddMergedRegion(CellRangeAddress.ValueOf("C1:F40"));


            XSSFWorkbook destWorkbook = new XSSFWorkbook();
            srcSheet.CopyTo(destWorkbook, srcSheet.SheetName, true, true);

            var destSheet = destWorkbook.GetSheet("Sheet1");
            ClassicAssert.NotNull(destSheet);
            ClassicAssert.AreEqual(2, destSheet.MergedRegions.Count);

            ClassicAssert.IsTrue(
                new string[]
                {
                    "A1:B4",
                    "C1:F40"
                }
                .SequenceEqual(
                    destSheet.MergedRegions
                             .Select(r => r.FormatAsString())));
        }

        [Test]
        public void CopySheetToWorkbookShouldCopyCharts()
        {
            XSSFWorkbook sourceWorkbook = XSSFTestDataSamples.OpenSampleWorkbook("WithThreeCharts.xlsx");

            XSSFSheet srcDataSheet = (XSSFSheet)sourceWorkbook.GetSheetAt(0);
            XSSFWorkbook destWorkbook = new XSSFWorkbook();
            srcDataSheet.CopyTo(destWorkbook, srcDataSheet.SheetName, true, true);

            XSSFSheet srcChartSheet = (XSSFSheet)sourceWorkbook.GetSheetAt(1);
            srcChartSheet.CopyTo(destWorkbook, srcChartSheet.SheetName, true, true);

            var destSheet = destWorkbook.GetSheetAt(1);
            
            ClassicAssert.NotNull(destSheet);

            ClassicAssert.AreEqual(2, (srcChartSheet.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);
            ClassicAssert.AreEqual(2, (destSheet.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);
            ClassicAssert.AreEqual(2, (destSheet.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[0].GetAxis().Count);
        }

        [Test]
        public void CopySheetToSameWorkbookShouldNotDuplicateIndexedFonts()
        {
            using var workbook = new XSSFWorkbook();
            var font = workbook.CreateFont();
            font.Color = IndexedColors.White.Index;
            var style = workbook.CreateCellStyle();
            style.SetFont(font);
            var sheet = workbook.CreateSheet("Source");
            var cell = sheet.CreateRow(0).CreateCell(0);
            cell.CellStyle = style;
            cell.SetCellValue("source");
            var fontCount = workbook.NumberOfFonts;

            sheet.CopyTo(workbook, "Copied", true, true);

            var copiedColor = ((XSSFFont)workbook.GetSheet("Copied").GetRow(0).GetCell(0).CellStyle.GetFont(workbook)).GetXSSFColor();
            ClassicAssert.IsTrue(copiedColor.IsIndexed);
            ClassicAssert.AreEqual(IndexedColors.White.Index, copiedColor.Indexed);
            ClassicAssert.AreEqual(fontCount, workbook.NumberOfFonts);
        }

        [Test]
        public void CopySheetToWorkbookShouldPreserveReportedWhiteFontColor()
        {
            using var srcWorkbook = XSSFTestDataSamples.OpenSampleWorkbook("1586-test2.xlsx");
            using var destWorkbook = new XSSFWorkbook();
            var srcSheet = (XSSFSheet)srcWorkbook.GetSheetAt(0);

            srcSheet.CopyTo(destWorkbook, "Copied", true, true);

            using var reloadedWorkbook = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(destWorkbook);
            var copiedSheet = reloadedWorkbook.GetSheet("Copied");
            foreach (var address in new[] { "A3", "B3" })
            {
                var cell = copiedSheet.GetRow(2).GetCell(new CellReference(address).Col);
                var font = (XSSFFont)cell.CellStyle.GetFont(reloadedWorkbook);
                ClassicAssert.AreEqual("FFFFFFFF", font.GetXSSFColor().ARGBHex, address);
            }
        }

        [Test]
        public void CopySheetToWorkbookShouldResolveIndexedFontColorWithSourcePalette()
        {
            using var srcWorkbook = new XSSFWorkbook();
            var srcFont = srcWorkbook.CreateFont();
            srcFont.FontName = "SourceWhite";
            srcFont.Color = IndexedColors.White.Index;
            var srcStyle = srcWorkbook.CreateCellStyle();
            srcStyle.SetFont(srcFont);
            var srcCell = srcWorkbook.CreateSheet("Source").CreateRow(0).CreateCell(0);
            srcCell.CellStyle = srcStyle;
            srcCell.SetCellValue("source");
            var automaticFont = srcWorkbook.CreateFont();
            automaticFont.Color = IndexedColors.Automatic.Index;
            var automaticStyle = srcWorkbook.CreateCellStyle();
            automaticStyle.SetFont(automaticFont);
            var automaticCell = srcWorkbook.GetSheet("Source").GetRow(0).CreateCell(1);
            automaticCell.CellStyle = automaticStyle;
            automaticCell.SetCellValue("automatic");
            SetIndexedColor(srcWorkbook, IndexedColors.White.Index, new byte[] { 255, 255, 255 });

            using var destWorkbook = new XSSFWorkbook();
            var destFont = destWorkbook.CreateFont();
            destFont.Color = IndexedColors.White.Index;
            var destStyle = destWorkbook.CreateCellStyle();
            destStyle.SetFont(destFont);
            var destCell = destWorkbook.CreateSheet("Existing").CreateRow(0).CreateCell(0);
            destCell.CellStyle = destStyle;
            destCell.SetCellValue("existing");
            SetIndexedColor(destWorkbook, IndexedColors.White.Index, new byte[] { 255, 0, 0 });

            using var reloadedSource = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(srcWorkbook);
            using var reloadedDest = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(destWorkbook);
            var sourceSheet = (XSSFSheet)reloadedSource.GetSheet("Source");

            sourceSheet.CopyTo(reloadedDest, "Copied", true, true);

            using var result = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(reloadedDest);
            var copiedFont = (XSSFFont)result.GetSheet("Copied").GetRow(0).GetCell(0).CellStyle.GetFont(result);
            var copiedColor = copiedFont.GetXSSFColor();
            ClassicAssert.IsFalse(copiedColor.IsIndexed);
            ClassicAssert.IsTrue(copiedColor.IsRGB);
            ClassicAssert.AreEqual("FFFFFFFF", copiedColor.ARGBHex);
            foreach (var font in result.GetStylesSource().Fonts.Where(f => f.FontName == "SourceWhite"))
            {
                var color = font.GetXSSFColor();
                ClassicAssert.IsFalse(color.IsIndexed);
                ClassicAssert.AreEqual("FFFFFFFF", color.ARGBHex);
            }

            var copiedAutomaticFont = (XSSFFont)result.GetSheet("Copied").GetRow(0).GetCell(1).CellStyle.GetFont(result);
            var automaticColor = copiedAutomaticFont.GetXSSFColor();
            ClassicAssert.IsTrue(automaticColor.IsIndexed);
            ClassicAssert.AreEqual(IndexedColors.Automatic.Index, automaticColor.Indexed);

            var existingColor = ((XSSFFont)result.GetSheet("Existing").GetRow(0).GetCell(0).CellStyle.GetFont(result)).GetXSSFColor();
            ClassicAssert.IsTrue(existingColor.IsIndexed);
            ClassicAssert.AreEqual("FFFF0000", existingColor.ARGBHex);
            CollectionAssert.AreEqual(new byte[] { 255, 0, 0 }, result.GetStylesSource().IndexedColors.GetRGB(IndexedColors.White.Index));
        }

        private static void SetIndexedColor(XSSFWorkbook workbook, short index, byte[] rgb)
        {
            var colors = new CT_Colors();
            colors.indexedColors = new List<CT_RgbColor>();
            for (var i = 0; i < 64; i++)
            {
                colors.indexedColors.Add(new CT_RgbColor
                {
                    rgb = i == index ? rgb : NPOI.OOXML.XSSF.UserModel.DefaultIndexedColorMap.GetDefaultRGB(i)
                });
            }
            workbook.GetStylesSource().GetCTStylesheet().colors = colors;
        }
    }
}
