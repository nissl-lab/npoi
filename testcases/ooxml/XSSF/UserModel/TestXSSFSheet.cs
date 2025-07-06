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

using NPOI;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.POIFS.Crypt;
using NPOI.SS;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.Model;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel.Helpers;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using TestCases.HSSF;
using TestCases.SS.UserModel;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFSheet : BaseTestXSheet
    {

        public TestXSSFSheet()
            : base(XSSFITestDataProvider.instance)
        {

        }

        //[Test]
        //TODO column styles are not yet supported by XSSF
        public override void DefaultColumnStyle()
        {
            base.DefaultColumnStyle();
        }
        [Test]
        public void TestTestGetSetMargin()
        {
            BaseTestGetSetMargin(new double[] { 0.7, 0.7, 0.75, 0.75, 0.3, 0.3 });
        }

        [Test]
        public void ShiftRows_ShiftRowsWithVariousMergedRegions_RowsShiftedWithMergedRegion()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            for (int i = 0; i < 12; i++)
            {
                sheet.CreateRow(i).CreateCell(0).SetCellValue(i);
            }

            sheet.GetRow(1).CreateCell(1).SetCellValue("regionOutsideShiftedRowsOnTop");
            int regionOutsideToTheLeftId = sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 2));

            sheet.GetRow(2).CreateCell(3).SetCellValue("regionInsideShiftedRows");
            int regionInsideId = sheet.AddMergedRegion(new CellRangeAddress(2, 3, 3, 4));

            sheet.GetRow(4).CreateCell(5).SetCellValue("regionRighBelowTheShiftedRows");
            int regionRighNextToId = sheet.AddMergedRegion(new CellRangeAddress(4, 5, 5, 6));

            sheet.GetRow(6).CreateCell(7).SetCellValue("regionInTheWayOfTheShift");
            int regionInTheWayOfTheShiftId = sheet.AddMergedRegion(new CellRangeAddress(6, 7, 7, 8));

            sheet.GetRow(10).CreateCell(9).SetCellValue("regionOutsideShiftedRowsBelow");
            int regionOutsideToTheRightId = sheet.AddMergedRegion(new CellRangeAddress(10, 11, 9, 10));

            sheet.GetRow(1).CreateCell(11).SetCellValue("regionThatEndsWithinShiftedRows");
            int regionThatEndsWithinShiftedRowsId = sheet.AddMergedRegion(new CellRangeAddress(1, 2, 11, 12));

            sheet.GetRow(1).CreateCell(13).SetCellValue("regionThatEndsOnLastShiftedRow");
            int regionThatEndsOnLastShiftedRowId = sheet.AddMergedRegion(new CellRangeAddress(1, 3, 13, 14));

            sheet.GetRow(1).CreateCell(15).SetCellValue("regionThatEndsOutsideShiftedRows");
            int regionThatEndsOutsideShiftedRowsId = sheet.AddMergedRegion(new CellRangeAddress(1, 4, 15, 16));

            sheet.GetRow(1).CreateCell(17).SetCellValue("reallyLongRegion");
            int reallyLongRegionId = sheet.AddMergedRegion(new CellRangeAddress(1, 11, 17, 18));

            ClassicAssert.AreEqual("regionOutsideShiftedRowsOnTop", sheet.GetRow(1).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B2:C2")));

            ClassicAssert.AreEqual("regionInsideShiftedRows", sheet.GetRow(2).GetCell(3).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("D3:E4")));

            ClassicAssert.AreEqual("regionRighBelowTheShiftedRows", sheet.GetRow(4).GetCell(5).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("F5:G6")));

            ClassicAssert.AreEqual("regionInTheWayOfTheShift", sheet.GetRow(6).GetCell(7).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("H7:I8")));

            ClassicAssert.AreEqual("regionOutsideShiftedRowsBelow", sheet.GetRow(10).GetCell(9).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("J11:K12")));

            ClassicAssert.AreEqual("regionThatEndsWithinShiftedRows", sheet.GetRow(1).GetCell(11).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("L2:M3")));

            ClassicAssert.AreEqual("regionThatEndsOnLastShiftedRow", sheet.GetRow(1).GetCell(13).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("N2:O4")));

            ClassicAssert.AreEqual("regionThatEndsOutsideShiftedRows", sheet.GetRow(1).GetCell(15).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("P2:Q5")));

            ClassicAssert.AreEqual("reallyLongRegion", sheet.GetRow(1).GetCell(17).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("R2:S12")));

            sheet.ShiftRows(2, 3, 4);

            ClassicAssert.AreEqual("regionOutsideShiftedRowsOnTop", sheet.GetRow(1).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B2:C2")));

            ClassicAssert.AreEqual("regionInsideShiftedRows", sheet.GetRow(6).GetCell(3).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("D7:E8")));

            ClassicAssert.AreEqual("regionRighBelowTheShiftedRows", sheet.GetRow(4).GetCell(5).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("F5:G6")));

            ClassicAssert.IsNull(sheet.GetRow(6).GetCell(7));
            //TODO: Cell Range H7:I8 was replaced by H4:I4,not merged, this assert shoud be false
            bool merged = sheet.MergedRegions.Any(r => r.FormatAsString().Equals("H7:I8"));
            Assume.That(merged, Is.False);
            ClassicAssert.True(merged);

            ClassicAssert.AreEqual("regionOutsideShiftedRowsBelow", sheet.GetRow(10).GetCell(9).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("J11:K12")));

            ClassicAssert.AreEqual("regionThatEndsWithinShiftedRows", sheet.GetRow(1).GetCell(11).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("L2:M3")));

            ClassicAssert.AreEqual("regionThatEndsOnLastShiftedRow", sheet.GetRow(1).GetCell(13).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("N2:O4")));

            ClassicAssert.AreEqual("regionThatEndsOutsideShiftedRows", sheet.GetRow(1).GetCell(15).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("P2:Q5")));

            ClassicAssert.AreEqual("reallyLongRegion", sheet.GetRow(1).GetCell(17).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("R2:S12")));

            FileInfo file = TempFile.CreateTempFile("ShiftRows-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            ClassicAssert.AreEqual("regionOutsideShiftedRowsOnTop", sheetLoaded.GetRow(1).GetCell(1).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("B2:C2")));

            ClassicAssert.AreEqual("regionInsideShiftedRows", sheetLoaded.GetRow(6).GetCell(3).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("D7:E8")));

            ClassicAssert.AreEqual("regionRighBelowTheShiftedRows", sheetLoaded.GetRow(4).GetCell(5).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("F5:G6")));

            ClassicAssert.IsNull(sheetLoaded.GetRow(6).GetCell(7));
            merged = sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("H7:I8"));
            Assume.That(merged, Is.False);
            ClassicAssert.True(merged);

            ClassicAssert.AreEqual("regionOutsideShiftedRowsBelow", sheetLoaded.GetRow(10).GetCell(9).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("J11:K12")));

            ClassicAssert.AreEqual("regionThatEndsWithinShiftedRows", sheetLoaded.GetRow(1).GetCell(11).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("L2:M3")));

            ClassicAssert.AreEqual("regionThatEndsOnLastShiftedRow", sheetLoaded.GetRow(1).GetCell(13).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("N2:O4")));

            ClassicAssert.AreEqual("regionThatEndsOutsideShiftedRows", sheetLoaded.GetRow(1).GetCell(15).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("P2:Q5")));

            ClassicAssert.AreEqual("reallyLongRegion", sheetLoaded.GetRow(1).GetCell(17).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("R2:S12")));
        }

        [Test]
        public void ShiftColumns_ShiftColumnsWithVariousMergedRegions_ColumnsShiftedWithMergedRegion()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            _ = sheet.CreateRow(0);

            for (int i = 0; i < 12; i++)
            {
                sheet.GetRow(0).CreateCell(i).SetCellValue(i);
            }

            sheet.CreateRow(1).CreateCell(1).SetCellValue("regionOutsideShiftedColumnsToTheLeft");
            int regionOutsideToTheLeftId = sheet.AddMergedRegion(new CellRangeAddress(1, 2, 1, 1));

            sheet.CreateRow(3).CreateCell(2).SetCellValue("regionInsideShiftedColumns");
            int regionInsideId = sheet.AddMergedRegion(new CellRangeAddress(3, 4, 2, 3));

            sheet.CreateRow(5).CreateCell(4).SetCellValue("regionRighNextToTheShiftedColumns");
            int regionRighNextToId = sheet.AddMergedRegion(new CellRangeAddress(5, 6, 4, 5));

            sheet.CreateRow(7).CreateCell(6).SetCellValue("regionInTheWayOfTheShift");
            int regionInTheWayOfTheShiftId = sheet.AddMergedRegion(new CellRangeAddress(7, 8, 6, 7));

            sheet.CreateRow(9).CreateCell(10).SetCellValue("regionOutsideShiftedColumnsToTheRight");
            int regionOutsideToTheRightId = sheet.AddMergedRegion(new CellRangeAddress(9, 10, 10, 11));

            sheet.CreateRow(11).CreateCell(1).SetCellValue("regionThatEndsWithinShiftedColumns");
            int regionThatEndsWithinShiftedColumnsId = sheet.AddMergedRegion(new CellRangeAddress(11, 12, 1, 2));

            sheet.CreateRow(13).CreateCell(1).SetCellValue("regionThatEndsOnLastShiftedColumn");
            int regionThatEndsOnLastShiftedColumnId = sheet.AddMergedRegion(new CellRangeAddress(13, 14, 1, 3));

            sheet.CreateRow(15).CreateCell(1).SetCellValue("regionThatEndsOutsideShiftedColumns");
            int regionThatEndsOutsideShiftedColumnsId = sheet.AddMergedRegion(new CellRangeAddress(15, 16, 1, 4));

            sheet.CreateRow(17).CreateCell(1).SetCellValue("reallyLongRegion");
            int reallyLongRegionId = sheet.AddMergedRegion(new CellRangeAddress(17, 18, 1, 11));

            ClassicAssert.AreEqual("regionOutsideShiftedColumnsToTheLeft", sheet.GetRow(1).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B2:B3")));

            ClassicAssert.AreEqual("regionInsideShiftedColumns", sheet.GetRow(3).GetCell(2).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("C4:D5")));

            ClassicAssert.AreEqual("regionRighNextToTheShiftedColumns", sheet.GetRow(5).GetCell(4).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("E6:F7")));

            ClassicAssert.AreEqual("regionInTheWayOfTheShift", sheet.GetRow(7).GetCell(6).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("G8:H9")));

            ClassicAssert.AreEqual("regionOutsideShiftedColumnsToTheRight", sheet.GetRow(9).GetCell(10).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("K10:L11")));

            ClassicAssert.AreEqual("regionThatEndsWithinShiftedColumns", sheet.GetRow(11).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B12:C13")));

            ClassicAssert.AreEqual("regionThatEndsOnLastShiftedColumn", sheet.GetRow(13).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B14:D15")));

            ClassicAssert.AreEqual("regionThatEndsOutsideShiftedColumns", sheet.GetRow(15).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B16:E17")));

            ClassicAssert.AreEqual("reallyLongRegion", sheet.GetRow(17).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B18:L19")));

            sheet.ShiftColumns(2, 3, 4);

            ClassicAssert.AreEqual("regionOutsideShiftedColumnsToTheLeft", sheet.GetRow(1).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B2:B3")));

            ClassicAssert.AreEqual("regionInsideShiftedColumns", sheet.GetRow(3).GetCell(6).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("G4:H5")));

            ClassicAssert.AreEqual("regionRighNextToTheShiftedColumns", sheet.GetRow(5).GetCell(4).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("E6:F7")));

            ClassicAssert.IsNull(sheet.GetRow(7).GetCell(6));
            ClassicAssert.False(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("G8:H9")));

            ClassicAssert.AreEqual("regionOutsideShiftedColumnsToTheRight", sheet.GetRow(9).GetCell(10).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("K10:L11")));

            ClassicAssert.AreEqual("regionThatEndsWithinShiftedColumns", sheet.GetRow(11).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B12:C13")));

            ClassicAssert.AreEqual("regionThatEndsOnLastShiftedColumn", sheet.GetRow(13).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B14:D15")));

            ClassicAssert.AreEqual("regionThatEndsOutsideShiftedColumns", sheet.GetRow(15).GetCell(1).StringCellValue);
            ClassicAssert.True(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B16:E17")));

            ClassicAssert.AreEqual("reallyLongRegion", sheet.GetRow(17).GetCell(1).StringCellValue);
            ClassicAssert.False(sheet.MergedRegions.Any(r => r.FormatAsString().Equals("B18:L19")));

            FileInfo file = TempFile.CreateTempFile("ShiftCols1-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            ClassicAssert.AreEqual("regionOutsideShiftedColumnsToTheLeft", sheetLoaded.GetRow(1).GetCell(1).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("B2:B3")));

            ClassicAssert.AreEqual("regionInsideShiftedColumns", sheetLoaded.GetRow(3).GetCell(6).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("G4:H5")));

            ClassicAssert.AreEqual("regionRighNextToTheShiftedColumns", sheetLoaded.GetRow(5).GetCell(4).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("E6:F7")));

            ClassicAssert.IsNull(sheetLoaded.GetRow(7).GetCell(6));
            ClassicAssert.False(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("G8:H9")));

            ClassicAssert.AreEqual("regionOutsideShiftedColumnsToTheRight", sheetLoaded.GetRow(9).GetCell(10).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("K10:L11")));

            ClassicAssert.AreEqual("regionThatEndsWithinShiftedColumns", sheetLoaded.GetRow(11).GetCell(1).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("B12:C13")));

            ClassicAssert.AreEqual("regionThatEndsOnLastShiftedColumn", sheetLoaded.GetRow(13).GetCell(1).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("B14:D15")));

            ClassicAssert.AreEqual("regionThatEndsOutsideShiftedColumns", sheetLoaded.GetRow(15).GetCell(1).StringCellValue);
            ClassicAssert.True(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("B16:E17")));

            ClassicAssert.AreEqual("reallyLongRegion", sheetLoaded.GetRow(17).GetCell(1).StringCellValue);
            ClassicAssert.False(sheetLoaded.MergedRegions.Any(r => r.FormatAsString().Equals("B18:L19")));
        }

        [Test]
        public void TestExistingHeaderFooter()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("45540_classic_Header.xlsx");
            XSSFOddHeader hdr;
            XSSFOddFooter ftr;

            // Sheet 1 has a header with center and right text
            XSSFSheet s1 = (XSSFSheet)wb1.GetSheetAt(0);
            ClassicAssert.IsNotNull(s1.Header);
            ClassicAssert.IsNotNull(s1.Footer);
            hdr = (XSSFOddHeader)s1.Header;
            ftr = (XSSFOddFooter)s1.Footer;

            ClassicAssert.AreEqual("&Ctestdoc&Rtest phrase", hdr.Text);
            ClassicAssert.AreEqual(null, ftr.Text);

            ClassicAssert.AreEqual("", hdr.Left);
            ClassicAssert.AreEqual("testdoc", hdr.Center);
            ClassicAssert.AreEqual("test phrase", hdr.Right);

            ClassicAssert.AreEqual("", ftr.Left);
            ClassicAssert.AreEqual("", ftr.Center);
            ClassicAssert.AreEqual("", ftr.Right);

            // Sheet 2 has a footer, but it's empty
            XSSFSheet s2 = (XSSFSheet)wb1.GetSheetAt(1);
            ClassicAssert.IsNotNull(s2.Header);
            ClassicAssert.IsNotNull(s2.Footer);
            hdr = (XSSFOddHeader)s2.Header;
            ftr = (XSSFOddFooter)s2.Footer;

            ClassicAssert.AreEqual(null, hdr.Text);
            ClassicAssert.AreEqual("&L&F", ftr.Text);

            ClassicAssert.AreEqual("", hdr.Left);
            ClassicAssert.AreEqual("", hdr.Center);
            ClassicAssert.AreEqual("", hdr.Right);

            ClassicAssert.AreEqual("&F", ftr.Left);
            ClassicAssert.AreEqual("", ftr.Center);
            ClassicAssert.AreEqual("", ftr.Right);

            // Save and reload
            IWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            hdr = (XSSFOddHeader)wb2.GetSheetAt(0).Header;
            ftr = (XSSFOddFooter)wb2.GetSheetAt(0).Footer;

            ClassicAssert.AreEqual("", hdr.Left);
            ClassicAssert.AreEqual("testdoc", hdr.Center);
            ClassicAssert.AreEqual("test phrase", hdr.Right);

            ClassicAssert.AreEqual("", ftr.Left);
            ClassicAssert.AreEqual("", ftr.Center);
            ClassicAssert.AreEqual("", ftr.Right);

            wb2.Close();
        }

        [Test]
        public void TestGetAllHeadersFooters()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet 1");
            ClassicAssert.IsNotNull(sheet.OddFooter);
            ClassicAssert.IsNotNull(sheet.EvenFooter);
            ClassicAssert.IsNotNull(sheet.FirstFooter);
            ClassicAssert.IsNotNull(sheet.OddHeader);
            ClassicAssert.IsNotNull(sheet.EvenHeader);
            ClassicAssert.IsNotNull(sheet.FirstHeader);

            ClassicAssert.AreEqual("", sheet.OddFooter.Left);
            sheet.OddFooter.Left = "odd footer left";
            ClassicAssert.AreEqual("odd footer left", sheet.OddFooter.Left);

            ClassicAssert.AreEqual("", sheet.EvenFooter.Left);
            sheet.EvenFooter.Left = "even footer left";
            ClassicAssert.AreEqual("even footer left", sheet.EvenFooter.Left);

            ClassicAssert.AreEqual("", sheet.FirstFooter.Left);
            sheet.FirstFooter.Left = "first footer left";
            ClassicAssert.AreEqual("first footer left", sheet.FirstFooter.Left);

            ClassicAssert.AreEqual("", sheet.OddHeader.Left);
            sheet.OddHeader.Left = "odd header left";
            ClassicAssert.AreEqual("odd header left", sheet.OddHeader.Left);

            ClassicAssert.AreEqual("", sheet.OddHeader.Right);
            sheet.OddHeader.Right = "odd header right";
            ClassicAssert.AreEqual("odd header right", sheet.OddHeader.Right);

            ClassicAssert.AreEqual("", sheet.OddHeader.Center);
            sheet.OddHeader.Center = "odd header center";
            ClassicAssert.AreEqual("odd header center", sheet.OddHeader.Center);

            // Defaults are odd
            ClassicAssert.AreEqual("odd footer left", sheet.Footer.Left);
            ClassicAssert.AreEqual("odd header center", sheet.Header.Center);

            workbook.Close();
        }
        [Test]
        public void TestAutoSizeColumn()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet 1");
            sheet.CreateRow(0).CreateCell(13).SetCellValue("test");

            sheet.AutoSizeColumn(13);

            ClassicAssert.IsTrue(sheet.GetColumn(13).IsBestFit);

            workbook.Close();
        }

        [Test]
        [Platform("Win")]
        public void TestAutoSizeRow()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet 1");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(13);
            cell.SetCellValue("test");
            IFont font = cell.CellStyle.GetFont(workbook);
            font.FontHeightInPoints = 20;
            cell.CellStyle.SetFont(font);
            row.Height = 100;

            sheet.AutoSizeRow(row.RowNum);

            ClassicAssert.AreNotEqual(100, row.Height);
            ClassicAssert.AreEqual(540, row.Height);

            workbook.Close();
        }

        [Test]
        public void TestSetCellComment()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();

            IDrawing dg = sheet.CreateDrawingPatriarch();
            IComment comment = dg.CreateCellComment(new XSSFClientAnchor());

            ICell cell = sheet.CreateRow(0).CreateCell(0);
            CommentsTable comments = sheet.GetCommentsTable(false);
            CT_Comments ctComments = comments.GetCTComments();

            cell.CellComment = comment;
            ClassicAssert.AreEqual("A1", ctComments.commentList.GetCommentArray(0).@ref);
            comment.Author = "test A1 author";
            ClassicAssert.AreEqual("test A1 author", comments.GetAuthor((int)ctComments.commentList.GetCommentArray(0).authorId));

            workbook.Close();
        }
        [Test]
        public void TestGetActiveCell()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CellAddress R5 = new CellAddress("R5");
            sheet.ActiveCell = R5;

            ClassicAssert.AreEqual(R5, sheet.ActiveCell);

            workbook.Close();

        }
        [Test]
        public void TestCreateFreezePane_XSSF()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CT_Worksheet ctWorksheet = sheet.GetCTWorksheet();

            sheet.CreateFreezePane(2, 4);
            ClassicAssert.AreEqual(2.0, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.xSplit);
            ClassicAssert.AreEqual(ST_Pane.bottomRight, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.activePane);
            sheet.CreateFreezePane(3, 6, 10, 10);
            ClassicAssert.AreEqual(3.0, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.xSplit);
            //ClassicAssert.AreEqual(10, sheet.TopRow);
            //ClassicAssert.AreEqual(10, sheet.LeftCol);
            sheet.CreateSplitPane(4, 8, 12, 12, PanePosition.LowerRight);
            ClassicAssert.AreEqual(8.0, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.ySplit);
            ClassicAssert.AreEqual(ST_Pane.bottomRight, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.activePane);

            workbook.Close();
        }

        [Test]
        public void TestRemoveMergedRegion_lowlevel()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CT_Worksheet ctWorksheet = sheet.GetCTWorksheet();
            CellRangeAddress region_1 = CellRangeAddress.ValueOf("A1:B2");
            CellRangeAddress region_2 = CellRangeAddress.ValueOf("C3:D4");
            CellRangeAddress region_3 = CellRangeAddress.ValueOf("E5:F6");
            sheet.AddMergedRegion(region_1);
            sheet.AddMergedRegion(region_2);
            sheet.AddMergedRegion(region_3);
            ClassicAssert.AreEqual("C3:D4", ctWorksheet.mergeCells.GetMergeCellArray(1).@ref);
            ClassicAssert.AreEqual(3, sheet.NumMergedRegions);
            sheet.RemoveMergedRegion(1);
            ClassicAssert.AreEqual("E5:F6", ctWorksheet.mergeCells.GetMergeCellArray(1).@ref);
            ClassicAssert.AreEqual(2, sheet.NumMergedRegions);
            sheet.RemoveMergedRegion(1);
            sheet.RemoveMergedRegion(0);
            ClassicAssert.AreEqual(0, sheet.NumMergedRegions);
            ClassicAssert.IsNull(sheet.GetCTWorksheet().mergeCells, " CTMergeCells should be deleted After removing the last merged " +
                    "region on the sheet.");

            workbook.Close();
        }

        [Test]
        public void TestSetDefaultColumnStyle()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CT_Worksheet ctWorksheet = sheet.GetCTWorksheet();
            StylesTable stylesTable = workbook.GetStylesSource();
            XSSFFont font = new XSSFFont
            {
                FontName = "Cambria"
            };
            stylesTable.PutFont(font);
            CT_Xf cellStyleXf = new CT_Xf
            {
                fontId = 1,
                fillId = 0,
                borderId = 0,
                numFmtId = 0
            };
            stylesTable.PutCellStyleXf(cellStyleXf);
            CT_Xf cellXf = new CT_Xf
            {
                xfId = 1
            };
            stylesTable.PutCellXf(cellXf);
            XSSFCellStyle cellStyle = new XSSFCellStyle(1, 1, stylesTable, null);
            ClassicAssert.AreEqual(1, cellStyle.FontIndex);

            sheet.SetDefaultColumnStyle(3, cellStyle);
            ClassicAssert.AreEqual(1u, ctWorksheet.GetColsArray(0).GetColArray(0).style);

            workbook.Close();
        }

        [Test]
        public void TestGroupUngroupColumn()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();

            //one level
            sheet.GroupColumn(2, 7);
            sheet.GroupColumn(10, 11);
            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            ClassicAssert.AreEqual(8, cols.sizeOfColArray());
            List<CT_Col> colArray = cols.GetColList();
            ClassicAssert.IsNotNull(colArray);
            ClassicAssert.AreEqual((uint)(2 + 1), colArray[0].min); // 1 based
            ClassicAssert.AreEqual((uint)(2 + 1), colArray[0].max); // 1 based
            ClassicAssert.AreEqual(1, colArray[0].outlineLevel);

            ClassicAssert.AreEqual((uint)(3 + 1), colArray[1].min); // 1 based
            ClassicAssert.AreEqual((uint)(3 + 1), colArray[1].max); // 1 based
            ClassicAssert.AreEqual(1, colArray[1].outlineLevel);

            ClassicAssert.AreEqual((uint)(4 + 1), colArray[2].min); // 1 based
            ClassicAssert.AreEqual((uint)(4 + 1), colArray[2].max); // 1 based
            ClassicAssert.AreEqual(1, colArray[2].outlineLevel);

            ClassicAssert.AreEqual((uint)(5 + 1), colArray[3].min); // 1 based
            ClassicAssert.AreEqual((uint)(5 + 1), colArray[3].max); // 1 based
            ClassicAssert.AreEqual(1, colArray[3].outlineLevel);

            ClassicAssert.AreEqual((uint)(6 + 1), colArray[4].min); // 1 based
            ClassicAssert.AreEqual((uint)(6 + 1), colArray[4].max); // 1 based
            ClassicAssert.AreEqual(1, colArray[4].outlineLevel);

            ClassicAssert.AreEqual((uint)(7 + 1), colArray[5].min); // 1 based
            ClassicAssert.AreEqual((uint)(7 + 1), colArray[5].max); // 1 based
            ClassicAssert.AreEqual(1, colArray[5].outlineLevel);

            ClassicAssert.AreEqual(0, sheet.GetColumnOutlineLevel(0));

            //two level
            sheet.GroupColumn(1, 2);
            cols = sheet.GetCTWorksheet().GetColsArray(0);
            ClassicAssert.AreEqual(9, cols.sizeOfColArray());
            colArray = cols.GetColList();
            ClassicAssert.AreEqual(2, colArray[1].outlineLevel);

            //three level
            sheet.GroupColumn(6, 8);
            sheet.GroupColumn(2, 3);
            cols = sheet.GetCTWorksheet().GetColsArray(0);
            ClassicAssert.AreEqual(10, cols.sizeOfColArray());
            colArray = cols.GetColList();
            ClassicAssert.AreEqual(3, colArray[1].outlineLevel);
            ClassicAssert.AreEqual(3, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelCol);

            sheet.UngroupColumn(8, 10);
            colArray = cols.GetColList();
            //ClassicAssert.AreEqual(3, colArray[1].outlineLevel);

            sheet.UngroupColumn(4, 6);
            sheet.UngroupColumn(2, 2);
            colArray = cols.GetColList();
            ClassicAssert.AreEqual(6, colArray.Count);
            ClassicAssert.AreEqual(2, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelCol);

            workbook.Close();
        }

        [Test]
        public void UngroupColumn_SimpleOneLevelUngroupNoOverlaps_ColumnsUngroupedAndDestroyed()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderLeft = BorderStyle.Double;

            sheet.GroupColumn(2, 7);
            sheet.GroupColumn(10, 11);
            sheet.GetColumn(2).CreateCell(0);
            sheet.GetColumn(10).ColumnStyle = style;

            ClassicAssert.AreEqual(8, sheet.GetCTWorksheet().cols[0].col.Count);

            for (int i = 2; i <= 7; i++)
            {
                ClassicAssert.AreEqual(1, sheet.GetColumn(i).OutlineLevel);
            }

            for (int i = 10; i <= 11; i++)
            {
                ClassicAssert.AreEqual(1, sheet.GetColumn(i).OutlineLevel);
            }

            sheet.UngroupColumn(2, 7);
            sheet.UngroupColumn(10, 11);

            ClassicAssert.AreEqual(2, sheet.GetCTWorksheet().cols[0].col.Count);
            ClassicAssert.AreEqual(0, sheet.GetColumn(2).OutlineLevel);
            ClassicAssert.AreEqual(0, sheet.GetColumn(10).OutlineLevel);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            workbook.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheetAt(0);

            ClassicAssert.AreEqual(2, sheetLoaded.GetCTWorksheet().cols[0].col.Count);
            ClassicAssert.AreEqual(0, sheetLoaded.GetColumn(2).OutlineLevel);
            ClassicAssert.AreEqual(0, sheetLoaded.GetColumn(10).OutlineLevel);

            wbLoaded.Close();
        }

        [Test]
        public void UngroupColumn_TwoOverlappingGroups_ColumnsUngroupedAndDestroyed()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderLeft = BorderStyle.Double;

            sheet.GroupColumn(2, 7);
            sheet.GroupColumn(6, 11);
            sheet.GetColumn(2).CreateCell(0);
            sheet.GetColumn(6).ColumnStyle = style;

            ClassicAssert.AreEqual(10, sheet.GetCTWorksheet().cols[0].col.Count);

            for (int i = 2; i <= 11; i++)
            {
                if (i == 6 || i == 7)
                {
                    ClassicAssert.AreEqual(2, sheet.GetColumn(i).OutlineLevel);
                }
                else
                {
                    ClassicAssert.AreEqual(1, sheet.GetColumn(i).OutlineLevel);
                }
            }

            sheet.UngroupColumn(2, 7);
            sheet.UngroupColumn(6, 11);

            ClassicAssert.AreEqual(2, sheet.GetCTWorksheet().cols[0].col.Count);
            ClassicAssert.AreEqual(0, sheet.GetColumn(2).OutlineLevel);
            ClassicAssert.AreEqual(0, sheet.GetColumn(6).OutlineLevel);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            workbook.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheetAt(0);

            ClassicAssert.AreEqual(2, sheetLoaded.GetCTWorksheet().cols[0].col.Count);
            ClassicAssert.AreEqual(0, sheetLoaded.GetColumn(2).OutlineLevel);
            ClassicAssert.AreEqual(0, sheetLoaded.GetColumn(6).OutlineLevel);

            wbLoaded.Close();
        }

        [Test]
        public void TestGroupUngroupRow()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();

            //one level
            sheet.GroupRow(9, 10);
            ClassicAssert.AreEqual(2, sheet.PhysicalNumberOfRows);
            CT_Row ctrow = ((XSSFRow)sheet.GetRow(9)).GetCTRow();

            ClassicAssert.IsNotNull(ctrow);
            ClassicAssert.AreEqual(10u, ctrow.r);
            ClassicAssert.AreEqual(1, ctrow.outlineLevel);
            ClassicAssert.AreEqual(1, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelRow);

            //two level
            sheet.GroupRow(10, 13);
            ClassicAssert.AreEqual(5, sheet.PhysicalNumberOfRows);
            ctrow = ((XSSFRow)sheet.GetRow(10)).GetCTRow();
            ClassicAssert.IsNotNull(ctrow);
            ClassicAssert.AreEqual(11u, ctrow.r);
            ClassicAssert.AreEqual(2, ctrow.outlineLevel);
            ClassicAssert.AreEqual(2, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelRow);

            sheet.UngroupRow(8, 10);
            ClassicAssert.AreEqual(4, sheet.PhysicalNumberOfRows);
            ClassicAssert.AreEqual(1, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelRow);

            sheet.UngroupRow(10, 10);
            ClassicAssert.AreEqual(3, sheet.PhysicalNumberOfRows);

            ClassicAssert.AreEqual(1, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelRow);

            workbook.Close();
        }
        [Test]
        public void TestSetZoom()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)workbook.CreateSheet("new sheet");
            sheet1.SetZoom(75);   // 75 percent magnification
            long zoom = sheet1.GetCTWorksheet().sheetViews.GetSheetViewArray(0).zoomScale;
            ClassicAssert.AreEqual(zoom, 75);

            sheet1.SetZoom(200);
            zoom = sheet1.GetCTWorksheet().sheetViews.GetSheetViewArray(0).zoomScale;
            ClassicAssert.AreEqual(zoom, 200);

            try
            {
                sheet1.SetZoom(500);
                Assert.Fail("Expecting exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.AreEqual("Valid scale values range from 10 to 400", e.Message);
            }
            finally
            {
                workbook.Close();
            }
        }

        /**
         * TODO - while this is internally consistent, I'm not
         *  completely clear in all cases what it's supposed to
         *  be doing... Someone who understands the goals a little
         *  better should really review this!
         */
        [Test]
        public void TestSetColumnGroupCollapsed()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)wb1.CreateSheet();
            CT_Cols cols = sheet1.GetCTWorksheet().GetColsArray(0);

            ClassicAssert.AreEqual(0, cols.sizeOfColArray());

            sheet1.GroupColumn(4, 7);
            sheet1.GroupColumn(9, 12);

            ClassicAssert.AreEqual(8, cols.sizeOfColArray());

            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(7).max); // 1 based

            sheet1.GroupColumn(10, 11);

            ClassicAssert.AreEqual(8, cols.sizeOfColArray());

            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(7).max); // 1 based

            // collapse columns - 1
            sheet1.SetColumnGroupCollapsed(5, true);

            ClassicAssert.AreEqual(9, cols.sizeOfColArray());

            ClassicAssert.AreEqual(true, cols.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(true, cols.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(8).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).max); // 1 based

            // expand columns - 1
            sheet1.SetColumnGroupCollapsed(5, false);

            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(8).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).max); // 1 based

            //collapse - 2
            sheet1.SetColumnGroupCollapsed(9, true);

            ClassicAssert.AreEqual(10, cols.sizeOfColArray());

            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(8).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(8).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(9).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(9).IsSetCollapsed());
            ClassicAssert.AreEqual(true, cols.GetColArray(9).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(9).outlineLevel);
            ClassicAssert.AreEqual(13 + 1, cols.GetColArray(9).min); // 1 based
            ClassicAssert.AreEqual(13 + 1, cols.GetColArray(9).max); // 1 based

            //expand - 2
            sheet1.SetColumnGroupCollapsed(9, false);

            ClassicAssert.AreEqual(10, cols.sizeOfColArray());

            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(8).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(9).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(9).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(9).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(9).outlineLevel);
            ClassicAssert.AreEqual(13 + 1, cols.GetColArray(9).min); // 1 based
            ClassicAssert.AreEqual(13 + 1, cols.GetColArray(9).max); // 1 based

            //DOCUMENTARE MEGLIO IL DISCORSO DEL LIVELLO
            //collapse - 3
            sheet1.SetColumnGroupCollapsed(10, true);

            ClassicAssert.AreEqual(10, cols.sizeOfColArray());

            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(true, cols.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(8).IsSetCollapsed());
            ClassicAssert.AreEqual(true, cols.GetColArray(8).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(8).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(9).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(9).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(9).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(9).outlineLevel);
            ClassicAssert.AreEqual(13 + 1, cols.GetColArray(9).min); // 1 based
            ClassicAssert.AreEqual(13 + 1, cols.GetColArray(9).max); // 1 based

            //expand - 3
            sheet1.SetColumnGroupCollapsed(10, false);

            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(8 + 1, cols.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(2, cols.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols.GetColArray(7).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(8).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(8).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(8).collapsed);
            ClassicAssert.AreEqual(1, cols.GetColArray(8).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols.GetColArray(8).max); // 1 based

            ClassicAssert.AreEqual(false, cols.GetColArray(9).IsSetHidden());
            ClassicAssert.AreEqual(true, cols.GetColArray(9).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols.GetColArray(9).collapsed);
            ClassicAssert.AreEqual(0, cols.GetColArray(9).outlineLevel);
            ClassicAssert.AreEqual(13 + 1, cols.GetColArray(9).min); // 1 based
            ClassicAssert.AreEqual(13 + 1, cols.GetColArray(9).max); // 1 based

            // write out and give back
            // Save and re-load
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            var sheet2 = (XSSFSheet)wb2.GetSheetAt(0);
            var cols2 = sheet2.GetCTWorksheet().GetColsArray(0);
            ClassicAssert.AreEqual(10, cols2.sizeOfColArray());

            ClassicAssert.AreEqual(false, cols2.GetColArray(0).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(0).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(0).collapsed);
            ClassicAssert.AreEqual(1, cols2.GetColArray(0).outlineLevel);
            ClassicAssert.AreEqual(4 + 1, cols2.GetColArray(0).min); // 1 based
            ClassicAssert.AreEqual(4 + 1, cols2.GetColArray(0).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(1).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(1).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(1).collapsed);
            ClassicAssert.AreEqual(1, cols2.GetColArray(1).outlineLevel);
            ClassicAssert.AreEqual(5 + 1, cols2.GetColArray(1).min); // 1 based
            ClassicAssert.AreEqual(5 + 1, cols2.GetColArray(1).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(2).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(2).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(2).collapsed);
            ClassicAssert.AreEqual(1, cols2.GetColArray(2).outlineLevel);
            ClassicAssert.AreEqual(6 + 1, cols2.GetColArray(2).min); // 1 based
            ClassicAssert.AreEqual(6 + 1, cols2.GetColArray(2).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(3).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(3).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(3).collapsed);
            ClassicAssert.AreEqual(1, cols2.GetColArray(3).outlineLevel);
            ClassicAssert.AreEqual(7 + 1, cols2.GetColArray(3).min); // 1 based
            ClassicAssert.AreEqual(7 + 1, cols2.GetColArray(3).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(4).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(4).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(4).collapsed);
            ClassicAssert.AreEqual(0, cols2.GetColArray(4).outlineLevel);
            ClassicAssert.AreEqual(8 + 1, cols2.GetColArray(4).min); // 1 based
            ClassicAssert.AreEqual(8 + 1, cols2.GetColArray(4).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(5).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(5).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(5).collapsed);
            ClassicAssert.AreEqual(1, cols2.GetColArray(5).outlineLevel);
            ClassicAssert.AreEqual(9 + 1, cols2.GetColArray(5).min); // 1 based
            ClassicAssert.AreEqual(9 + 1, cols2.GetColArray(5).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(6).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(6).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(6).collapsed);
            ClassicAssert.AreEqual(2, cols2.GetColArray(6).outlineLevel);
            ClassicAssert.AreEqual(10 + 1, cols2.GetColArray(6).min); // 1 based
            ClassicAssert.AreEqual(10 + 1, cols2.GetColArray(6).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(7).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(7).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(7).collapsed);
            ClassicAssert.AreEqual(2, cols2.GetColArray(7).outlineLevel);
            ClassicAssert.AreEqual(11 + 1, cols2.GetColArray(7).min); // 1 based
            ClassicAssert.AreEqual(11 + 1, cols2.GetColArray(7).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(8).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(8).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(8).collapsed);
            ClassicAssert.AreEqual(1, cols2.GetColArray(8).outlineLevel);
            ClassicAssert.AreEqual(12 + 1, cols2.GetColArray(8).min); // 1 based
            ClassicAssert.AreEqual(12 + 1, cols2.GetColArray(8).max); // 1 based

            ClassicAssert.AreEqual(false, cols2.GetColArray(9).IsSetHidden());
            ClassicAssert.AreEqual(false, cols2.GetColArray(9).IsSetCollapsed());
            ClassicAssert.AreEqual(false, cols2.GetColArray(9).collapsed);
            ClassicAssert.AreEqual(0, cols2.GetColArray(9).outlineLevel);
            ClassicAssert.AreEqual(13 + 1, cols2.GetColArray(9).min); // 1 based
            ClassicAssert.AreEqual(13 + 1, cols2.GetColArray(9).max); // 1 based

            wb2.Close();
        }

        /**
     * Verify that column groups were created correctly after Sheet.groupColumn
     *
     * @param col the column group xml bean
     * @param fromColumnIndex 0-indexed
     * @param toColumnIndex 0-indexed
     */
        private static void checkColumnGroup(
                CT_Col col,
                int fromColumnIndex, int toColumnIndex,
                bool isSetHidden, bool isSetCollapsed
                )
        {
            ClassicAssert.AreEqual(fromColumnIndex, col.min - 1, "from column index"); // 1 based
            ClassicAssert.AreEqual(toColumnIndex, col.max - 1, "to column index"); // 1 based
            ClassicAssert.AreEqual(isSetHidden, col.IsSetHidden(), "isSetHidden");
            ClassicAssert.AreEqual(isSetCollapsed, col.IsSetCollapsed(), "isSetCollapsed"); //not necessarily set
        }

        /**
         * Verify that column groups were created correctly after Sheet.groupColumn
         *
         * @param col the column group xml bean
         * @param fromColumnIndex 0-indexed
         * @param toColumnIndex 0-indexed
         */
        private static void checkColumnGroup(
                CT_Col col,
                int fromColumnIndex, int toColumnIndex
                )
        {
            ClassicAssert.AreEqual(fromColumnIndex, col.min - 1, "from column index"); // 1 based
            ClassicAssert.AreEqual(toColumnIndex, col.max - 1, "to column index"); // 1 based
            ClassicAssert.IsFalse(col.IsSetHidden(), "isSetHidden");
            ClassicAssert.IsTrue(col.IsSetCollapsed(), "isSetCollapsed"); //not necessarily set
        }
        /**
         * Verify that column groups were created correctly after Sheet.groupColumn
         *
         * @param col the column group xml bean
         * @param fromColumnIndex 0-indexed
         * @param toColumnIndex 0-indexed
         */
        private static void checkColumnGroupIsCollapsed(
                CT_Col col,
                int fromColumnIndex, int toColumnIndex
                )
        {
            ClassicAssert.AreEqual(fromColumnIndex, col.min - 1, "from column index"); // 1 based
            ClassicAssert.AreEqual(toColumnIndex, col.max - 1, "to column index"); // 1 based
            ClassicAssert.IsTrue(col.IsSetHidden(), "isSetHidden");
            ClassicAssert.IsTrue(col.IsSetCollapsed(), "isSetCollapsed");
            //assertTrue("getCollapsed", col.getCollapsed());
        }
        /**
         * Verify that column groups were created correctly after Sheet.groupColumn
         *
         * @param col the column group xml bean
         * @param fromColumnIndex 0-indexed
         * @param toColumnIndex 0-indexed
         */
        private static void checkColumnGroupIsExpanded(
                CT_Col col,
                int fromColumnIndex, int toColumnIndex
                )
        {
            ClassicAssert.AreEqual(fromColumnIndex, col.min - 1, "from column index"); // 1 based
            ClassicAssert.AreEqual(toColumnIndex, col.max - 1, "to column index"); // 1 based
            ClassicAssert.IsFalse(col.IsSetHidden(), "isSetHidden");
            ClassicAssert.IsTrue(col.IsSetCollapsed(), "isSetCollapsed");
            //assertTrue("isSetCollapsed", !col.isSetCollapsed() || !col.getCollapsed());
            //assertFalse("getCollapsed", col.getCollapsed());
        }

        /**
         * TODO - while this is internally consistent, I'm not
         *  completely clear in all cases what it's supposed to
         *  be doing... Someone who understands the goals a little
         *  better should really review this!
         */
        [Test]
        public void TestSetRowGroupCollapsed()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)wb1.CreateSheet();

            sheet1.GroupRow(5, 14);
            sheet1.GroupRow(7, 14);
            sheet1.GroupRow(16, 19);

            ClassicAssert.AreEqual(14, sheet1.PhysicalNumberOfRows);
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetHidden());

            //collapsed
            sheet1.SetRowGroupCollapsed(7, true);

            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetHidden());

            //expanded
            sheet1.SetRowGroupCollapsed(7, false);

            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetHidden());

            // Save and re-load
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            sheet1 = (XSSFSheet)wb2.GetSheetAt(0);

            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsTrue(((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetHidden());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetCollapsed());
            ClassicAssert.IsFalse(((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetHidden());

            wb2.Close();
        }

        /**
         * Get / Set column width and check the actual values of the underlying XML beans
         */
        [Test]
        public void TestColumnWidth_lowlevel()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet 1");
            sheet.SetColumnWidth(1, 22 * 256);
            ClassicAssert.AreEqual(22 * 256, sheet.GetColumnWidth(1));

            // Now check the low level stuff, and check that's all
            //  been Set correctly
            XSSFSheet xs = sheet;
            CT_Worksheet cts = xs.GetCTWorksheet();

            List<CT_Cols> cols_s = cts.GetColsList();
            ClassicAssert.AreEqual(1, cols_s.Count);
            CT_Cols cols = cols_s[0];
            ClassicAssert.AreEqual(1, cols.sizeOfColArray());
            CT_Col col = cols.GetColArray(0);

            // XML is 1 based, POI is 0 based
            ClassicAssert.AreEqual(2u, col.min);
            ClassicAssert.AreEqual(2u, col.max);
            ClassicAssert.AreEqual(22.0, col.width, 0.0);
            ClassicAssert.IsTrue(col.customWidth);

            // Now Set another
            sheet.SetColumnWidth(3, 33 * 256);

            cols_s = cts.GetColsList();
            ClassicAssert.AreEqual(1, cols_s.Count);
            cols = cols_s[0];
            ClassicAssert.AreEqual(2, cols.sizeOfColArray());

            col = cols.GetColArray(0);
            ClassicAssert.AreEqual(2u, col.min); // POI 1
            ClassicAssert.AreEqual(2u, col.max);
            ClassicAssert.AreEqual(22.0, col.width, 0.0);
            ClassicAssert.IsTrue(col.customWidth);

            col = cols.GetColArray(1);
            ClassicAssert.AreEqual(4u, col.min); // POI 3
            ClassicAssert.AreEqual(4u, col.max);
            ClassicAssert.AreEqual(33.0, col.width, 0.0);
            ClassicAssert.IsTrue(col.customWidth);

            workbook.Close();
        }

        /**
         * Setting width of a column included in a column span
         */
        [Test]
        public void Test47862()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("47862.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb1.GetSheetAt(0);
            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            //<cols>
            //  <col min="1" max="5" width="15.77734375" customWidth="1"/>
            //</cols>

            //a span of columns [1,5]
            ClassicAssert.AreEqual(5, cols.sizeOfColArray());
            CT_Col col = cols.GetColArray(0);
            ClassicAssert.AreEqual((uint)1, col.min);
            ClassicAssert.AreEqual((uint)1, col.max);
            col = cols.GetColArray(1);
            ClassicAssert.AreEqual((uint)2, col.min);
            ClassicAssert.AreEqual((uint)2, col.max);
            col = cols.GetColArray(2);
            ClassicAssert.AreEqual((uint)3, col.min);
            ClassicAssert.AreEqual((uint)3, col.max);
            col = cols.GetColArray(3);
            ClassicAssert.AreEqual((uint)4, col.min);
            ClassicAssert.AreEqual((uint)4, col.max);
            col = cols.GetColArray(4);
            ClassicAssert.AreEqual((uint)5, col.min);
            ClassicAssert.AreEqual((uint)5, col.max);
            double swidth = 15.77734375; //width of columns in the span
            ClassicAssert.AreEqual(swidth, col.width, 0.0);

            for (int i = 0; i < 5; i++)
            {
                ClassicAssert.AreEqual((int)(swidth * 256), sheet.GetColumnWidth(i));
            }

            int[] cw = new int[] { 10, 15, 20, 25, 30 };
            for (int i = 0; i < 5; i++)
            {
                sheet.SetColumnWidth(i, cw[i] * 256);
            }

            //the check below failed prior to fix of Bug #47862
            ColumnHelper.SortColumns(cols);
            //<cols>
            //  <col min="1" max="1" customWidth="true" width="10.0" />
            //  <col min="2" max="2" customWidth="true" width="15.0" />
            //  <col min="3" max="3" customWidth="true" width="20.0" />
            //  <col min="4" max="4" customWidth="true" width="25.0" />
            //  <col min="5" max="5" customWidth="true" width="30.0" />
            //</cols>

            //now the span is splitted into 5 individual columns
            ClassicAssert.AreEqual(5, cols.sizeOfColArray());
            for (int i = 0; i < 5; i++)
            {
                ClassicAssert.AreEqual(cw[i] * 256, sheet.GetColumnWidth(i));
                ClassicAssert.AreEqual(cw[i], cols.GetColArray(i).width, 0.0);
            }

            //serialize and check again
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = (XSSFSheet)wb2.GetSheetAt(0);
            cols = sheet.GetCTWorksheet().GetColsArray(0);
            ClassicAssert.AreEqual(5, cols.sizeOfColArray());
            for (int i = 0; i < 5; i++)
            {
                ClassicAssert.AreEqual(cw[i] * 256, sheet.GetColumnWidth(i));
                ClassicAssert.AreEqual(cw[i], cols.GetColArray(i).width, 0.0);
            }

            wb2.Close();
        }

        /**
         * Hiding a column included in a column span
         */
        [Test]
        public void Test47804()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("47804.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb1.GetSheetAt(0);
            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            ClassicAssert.AreEqual(4, cols.sizeOfColArray());
            CT_Col col;
            //<cols>
            //  <col min="2" max="4" width="12" customWidth="1"/>
            //  <col min="7" max="7" width="10.85546875" customWidth="1"/>
            //</cols>

            //a span of columns [2,4]
            col = cols.GetColArray(0);
            ClassicAssert.AreEqual((uint)2, col.min);
            ClassicAssert.AreEqual((uint)2, col.max);
            col = cols.GetColArray(1);
            ClassicAssert.AreEqual((uint)3, col.min);
            ClassicAssert.AreEqual((uint)3, col.max);
            col = cols.GetColArray(2);
            ClassicAssert.AreEqual((uint)4, col.min);
            ClassicAssert.AreEqual((uint)4, col.max);
            //individual column
            col = cols.GetColArray(3);
            ClassicAssert.AreEqual((uint)7, col.min);
            ClassicAssert.AreEqual((uint)7, col.max);

            sheet.SetColumnHidden(2, true); // Column C
            sheet.SetColumnHidden(6, true); // Column G

            ClassicAssert.IsTrue(sheet.IsColumnHidden(2));
            ClassicAssert.IsTrue(sheet.IsColumnHidden(6));

            //other columns but C and G are not hidden
            ClassicAssert.IsFalse(sheet.IsColumnHidden(1));
            ClassicAssert.IsFalse(sheet.IsColumnHidden(3));
            ClassicAssert.IsFalse(sheet.IsColumnHidden(4));
            ClassicAssert.IsFalse(sheet.IsColumnHidden(5));

            //the check below failed prior to fix of Bug #47804
            ColumnHelper.SortColumns(cols);
            //the span is now splitted into three parts
            //<cols>
            //  <col min="2" max="2" customWidth="true" width="12.0" />
            //  <col min="3" max="3" customWidth="true" width="12.0" hidden="true"/>
            //  <col min="4" max="4" customWidth="true" width="12.0"/>
            //  <col min="7" max="7" customWidth="true" width="10.85546875" hidden="true"/>
            //</cols>

            ClassicAssert.AreEqual(4, cols.sizeOfColArray());
            col = cols.GetColArray(0);
            ClassicAssert.AreEqual((uint)2, col.min);
            ClassicAssert.AreEqual((uint)2, col.max);
            col = cols.GetColArray(1);
            ClassicAssert.AreEqual((uint)3, col.min);
            ClassicAssert.AreEqual((uint)3, col.max);
            col = cols.GetColArray(2);
            ClassicAssert.AreEqual((uint)4, col.min);
            ClassicAssert.AreEqual((uint)4, col.max);
            col = cols.GetColArray(3);
            ClassicAssert.AreEqual((uint)7, col.min);
            ClassicAssert.AreEqual((uint)7, col.max);

            //serialize and check again
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = (XSSFSheet)wb2.GetSheetAt(0);
            ClassicAssert.IsTrue(sheet.IsColumnHidden(2));
            ClassicAssert.IsTrue(sheet.IsColumnHidden(6));
            ClassicAssert.IsFalse(sheet.IsColumnHidden(1));
            ClassicAssert.IsFalse(sheet.IsColumnHidden(3));
            ClassicAssert.IsFalse(sheet.IsColumnHidden(4));
            ClassicAssert.IsFalse(sheet.IsColumnHidden(5));

            wb2.Close();
        }
        [Test]
        public void TestCommentsTable()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)wb1.CreateSheet();
            CommentsTable comment1 = sheet1.GetCommentsTable(false);
            ClassicAssert.IsNull(comment1);

            comment1 = sheet1.GetCommentsTable(true);
            ClassicAssert.IsNotNull(comment1);
            ClassicAssert.AreEqual("/xl/comments1.xml", comment1.GetPackagePart().PartName.Name);

            ClassicAssert.AreSame(comment1, sheet1.GetCommentsTable(true));

            //second sheet
            XSSFSheet sheet2 = (XSSFSheet)wb1.CreateSheet();
            CommentsTable comment2 = sheet2.GetCommentsTable(false);
            ClassicAssert.IsNull(comment2);

            comment2 = sheet2.GetCommentsTable(true);
            ClassicAssert.IsNotNull(comment2);

            ClassicAssert.AreSame(comment2, sheet2.GetCommentsTable(true));
            ClassicAssert.AreEqual("/xl/comments2.xml", comment2.GetPackagePart().PartName.Name);

            //comment1 and  comment2 are different objects
            ClassicAssert.AreNotSame(comment1, comment2);
            wb1.Close();

            //now Test against a workbook Containing cell comments
            XSSFWorkbook wb2 = XSSFTestDataSamples.OpenSampleWorkbook("WithMoreVariousData.xlsx");
            sheet1 = (XSSFSheet)wb2.GetSheetAt(0);
            comment1 = sheet1.GetCommentsTable(true);
            ClassicAssert.IsNotNull(comment1);
            ClassicAssert.AreEqual("/xl/comments1.xml", comment1.GetPackagePart().PartName.Name);
            ClassicAssert.AreSame(comment1, sheet1.GetCommentsTable(true));

            wb2.Close();
        }

        /**
         * Rows and cells can be Created in random order,
         * but CTRows are kept in ascending order
         */
        [Test]
        public new void TestCreateRow()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb1.CreateSheet();
            CT_Worksheet wsh = sheet.GetCTWorksheet();
            CT_SheetData sheetData = wsh.sheetData;
            ClassicAssert.AreEqual(0, sheetData.SizeOfRowArray());

            XSSFRow row1 = (XSSFRow)sheet.CreateRow(2);
            row1.CreateCell(2);
            row1.CreateCell(1);

            XSSFRow row2 = (XSSFRow)sheet.CreateRow(1);
            row2.CreateCell(2);
            row2.CreateCell(1);
            row2.CreateCell(0);

            XSSFRow row3 = (XSSFRow)sheet.CreateRow(0);
            row3.CreateCell(3);
            row3.CreateCell(0);
            row3.CreateCell(2);
            row3.CreateCell(5);

            List<CT_Row> xrow = sheetData.row;
            ClassicAssert.AreEqual(3, xrow.Count);

            //rows are sorted: {0, 1, 2}
            ClassicAssert.AreEqual(4, xrow[0].SizeOfCArray());
            ClassicAssert.AreEqual(1u, xrow[0].r);
            ClassicAssert.IsTrue(xrow[0].Equals(row3.GetCTRow()));

            ClassicAssert.AreEqual(3, xrow[1].SizeOfCArray());
            ClassicAssert.AreEqual(2u, xrow[1].r);
            ClassicAssert.IsTrue(xrow[1].Equals(row2.GetCTRow()));

            ClassicAssert.AreEqual(2, xrow[2].SizeOfCArray());
            ClassicAssert.AreEqual(3u, xrow[2].r);
            ClassicAssert.IsTrue(xrow[2].Equals(row1.GetCTRow()));

            List<CT_Cell> xcell = xrow[0].c;
            ClassicAssert.AreEqual("D1", xcell[0].r);
            ClassicAssert.AreEqual("A1", xcell[1].r);
            ClassicAssert.AreEqual("C1", xcell[2].r);
            ClassicAssert.AreEqual("F1", xcell[3].r);

            //re-creating a row does NOT add extra data to the parent
            row2 = (XSSFRow)sheet.CreateRow(1);
            ClassicAssert.AreEqual(3, sheetData.SizeOfRowArray());
            //existing cells are invalidated
            ClassicAssert.AreEqual(0, sheetData.GetRowArray(1).SizeOfCArray());
            ClassicAssert.AreEqual(0, row2.PhysicalNumberOfCells);

            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = (XSSFSheet)wb2.GetSheetAt(0);
            wsh = sheet.GetCTWorksheet();
            xrow = sheetData.row;
            ClassicAssert.AreEqual(3, xrow.Count);

            //rows are sorted: {0, 1, 2}
            ClassicAssert.AreEqual(4, xrow[0].SizeOfCArray());
            ClassicAssert.AreEqual(1u, xrow[0].r);
            //cells are now sorted
            xcell = xrow[0].c;
            ClassicAssert.AreEqual("A1", xcell[0].r);
            ClassicAssert.AreEqual("C1", xcell[1].r);
            ClassicAssert.AreEqual("D1", xcell[2].r);
            ClassicAssert.AreEqual("F1", xcell[3].r);

            ClassicAssert.AreEqual(0, xrow[1].SizeOfCArray());
            ClassicAssert.AreEqual(2u, xrow[1].r);

            ClassicAssert.AreEqual(2, xrow[2].SizeOfCArray());
            ClassicAssert.AreEqual(3u, xrow[2].r);

            wb2.Close();
        }

        [Test]
        public void TestSetAutoFilter()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("new sheet");
            sheet.SetAutoFilter(CellRangeAddress.ValueOf("A1:D100"));

            ClassicAssert.AreEqual("A1:D100", sheet.GetCTWorksheet().autoFilter.@ref);

            // auto-filter must be registered in workboook.xml, see Bugzilla 50315
            XSSFName nm = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, 0);
            ClassicAssert.IsNotNull(nm);

            ClassicAssert.AreEqual(0u, nm.GetCTName().localSheetId);
            ClassicAssert.IsTrue(nm.GetCTName().hidden);
            ClassicAssert.AreEqual("_xlnm._FilterDatabase", nm.GetCTName().name);
            ClassicAssert.AreEqual("'new sheet'!$A$1:$D$100", nm.GetCTName().Value);

            wb.Close();
        }
        [Test]
        public void TestProtectSheet_lowlevel()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            CT_SheetProtection pr = sheet.GetCTWorksheet().sheetProtection;
            ClassicAssert.IsNull(pr, "CTSheetProtection should be null by default");
            String password = "Test";
            sheet.ProtectSheet(password);
            pr = sheet.GetCTWorksheet().sheetProtection;
            ClassicAssert.IsNotNull(pr, "CTSheetProtection should be not null");
            ClassicAssert.IsTrue(pr.sheet, "sheet protection should be on");
            ClassicAssert.IsTrue(pr.objects, "object protection should be on");
            ClassicAssert.IsTrue(pr.scenarios, "scenario protection should be on");
            int hashVal = CryptoFunctions.CreateXorVerifier1(password);
            int actualVal = int.Parse(pr.password, NumberStyles.HexNumber);
            ClassicAssert.AreEqual(hashVal, actualVal, "well known value for top secret hash should match");

            sheet.ProtectSheet(null);
            ClassicAssert.IsNull(sheet.GetCTWorksheet().sheetProtection, "protectSheet(null) should unset CTSheetProtection");

            wb.Close();
        }

        [Test]
        public void protectSheet_emptyPassword()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            CT_SheetProtection pr = sheet.GetCTWorksheet().sheetProtection;
            ClassicAssert.IsNull(pr, "CTSheetProtection should be null by default");
            String password = "";
            sheet.ProtectSheet(password);
            pr = sheet.GetCTWorksheet().sheetProtection;
            ClassicAssert.IsNotNull(pr, "CTSheetProtection should be not null");
            ClassicAssert.IsTrue(pr.IsSetSheet(), "sheet protection should be on");
            ClassicAssert.IsTrue(pr.IsSetObjects(), "object protection should be on");
            ClassicAssert.IsTrue(pr.IsSetScenarios(), "scenario protection should be on");
            int hashVal = CryptoFunctions.CreateXorVerifier1(password);
            ST_UnsignedshortHex xpassword = new ST_UnsignedshortHex() { StringValue = pr.password };
            int actualVal = int.Parse(xpassword.StringValue, NumberStyles.HexNumber);
            ClassicAssert.AreEqual(hashVal, actualVal, "well known value for top secret hash should match");
            sheet.ProtectSheet(null);
            ClassicAssert.IsNull(sheet.GetCTWorksheet().sheetProtection, "protectSheet(null) should unset CTSheetProtection");
            wb.Close();
        }

        [Test]
        public void ProtectSheet_lowlevel_2013()
        {
            String password = "test";
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet xs = wb1.CreateSheet() as XSSFSheet;
            xs.SetSheetPassword(password, HashAlgorithm.sha384);
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            ClassicAssert.IsTrue((wb2.GetSheetAt(0) as XSSFSheet).ValidateSheetPassword(password));
            wb2.Close();

            XSSFWorkbook wb3 = XSSFTestDataSamples.OpenSampleWorkbook("workbookProtection-sheet_password-2013.xlsx");
            ClassicAssert.IsTrue((wb3.GetSheetAt(0) as XSSFSheet).ValidateSheetPassword("pwd"));

            wb3.Close();
        }

        [Test]
        public void Test49966()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("49966.xlsx");
            CalculationChain calcChain = wb1.GetCalculationChain();
            ClassicAssert.IsNotNull(wb1.GetCalculationChain());
            ClassicAssert.AreEqual(3, calcChain.GetCTCalcChain().SizeOfCArray());

            ISheet sheet = wb1.GetSheetAt(0);
            IRow row = sheet.GetRow(0);

            sheet.RemoveRow(row);
            ClassicAssert.AreEqual(0, calcChain.GetCTCalcChain().SizeOfCArray(), "XSSFSheet#RemoveRow did not clear calcChain entries");

            //calcChain should be gone 
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            ClassicAssert.IsNull(wb2.GetCalculationChain());

            wb2.Close();
        }

        /**
         * See bug #50829
         */
        [Test]
        public void TestTables()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithTable.xlsx");
            ClassicAssert.AreEqual(3, wb.NumberOfSheets);

            // Check the table sheet
            XSSFSheet s1 = (XSSFSheet)wb.GetSheetAt(0);
            ClassicAssert.AreEqual("a", s1.GetRow(0).GetCell(0).RichStringCellValue.ToString());
            ClassicAssert.AreEqual(1.0, s1.GetRow(1).GetCell(0).NumericCellValue);

            List<XSSFTable> tables = s1.GetTables();
            ClassicAssert.IsNotNull(tables);
            ClassicAssert.AreEqual(1, tables.Count);

            XSSFTable table = tables[0];
            ClassicAssert.AreEqual("Tabella1", table.Name);
            ClassicAssert.AreEqual("Tabella1", table.DisplayName);

            // And the others
            XSSFSheet s2 = (XSSFSheet)wb.GetSheetAt(1);
            ClassicAssert.AreEqual(0, s2.GetTables().Count);
            XSSFSheet s3 = (XSSFSheet)wb.GetSheetAt(2);
            ClassicAssert.AreEqual(0, s3.GetTables().Count);

            wb.Close();
        }

        /**
     * Test to trigger OOXML-LITE generating to include org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetCalcPr
     */
        [Test]
        public void TestSetForceFormulaRecalculation()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb1.CreateSheet("Sheet 1");

            ClassicAssert.IsFalse(sheet.ForceFormulaRecalculation);

            // Set
            sheet.ForceFormulaRecalculation = true;
            ClassicAssert.AreEqual(true, sheet.ForceFormulaRecalculation);

            // calcMode="manual" is unset when forceFormulaRecalculation=true
            CT_CalcPr calcPr = wb1.GetCTWorkbook().AddNewCalcPr();
            calcPr.calcMode = ST_CalcMode.manual;
            sheet.ForceFormulaRecalculation = true;
            ClassicAssert.AreEqual(ST_CalcMode.auto, calcPr.calcMode);

            // Check
            sheet.ForceFormulaRecalculation = false;
            ClassicAssert.AreEqual(false, sheet.ForceFormulaRecalculation);

            // Save, re-load, and re-check
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = (XSSFSheet)wb2.GetSheet("Sheet 1");
            ClassicAssert.IsFalse(sheet.ForceFormulaRecalculation);

            wb2.Close();
        }
        [Test]
        public void Bug54607()
        {
            // run with the file provided in the Bug-Report
            runGetTopRow("54607.xlsx", true, 1, 0, 0);
            runGetLeftCol("54607.xlsx", true, 0, 0, 0);

            // run with some other flie to see
            runGetTopRow("54436.xlsx", true, 0);
            runGetLeftCol("54436.xlsx", true, 0);
            runGetTopRow("TwoSheetsNoneHidden.xlsx", true, 0, 0);
            runGetLeftCol("TwoSheetsNoneHidden.xlsx", true, 0, 0);
            runGetTopRow("TwoSheetsNoneHidden.xls", false, 0, 0);
            runGetLeftCol("TwoSheetsNoneHidden.xls", false, 0, 0);
        }

        private void runGetTopRow(String file, bool isXSSF, params int[] topRows)
        {
            IWorkbook wb = isXSSF ?
                wb = XSSFTestDataSamples.OpenSampleWorkbook(file) :
                wb = HSSFTestDataSamples.OpenSampleWorkbook(file);

            for (int si = 0; si < wb.NumberOfSheets; si++)
            {
                ISheet sh = wb.GetSheetAt(si);
                ClassicAssert.IsNotNull(sh.SheetName);
                ClassicAssert.AreEqual(topRows[si], sh.TopRow, "Did not match for sheet " + si);
            }

            Assert.Warn("test about SXSSFWorkbook was commented");
            // for XSSF also test with SXSSF
            if (isXSSF)
            {
                IWorkbook swb = new SXSSFWorkbook((XSSFWorkbook)wb);
                try
                {
                    for (int si = 0; si < swb.NumberOfSheets; si++)
                    {
                        ISheet sh = swb.GetSheetAt(si);
                        ClassicAssert.IsNotNull(sh.SheetName);
                        ClassicAssert.AreEqual(topRows[si], sh.TopRow, "Did not match for sheet " + si);
                    }
                }
                finally
                {
                    swb.Close();
                }
            }

            wb.Close();
        }

        private void runGetLeftCol(String file, bool isXSSF, params int[] topRows)
        {
            IWorkbook wb = isXSSF ?
                wb = XSSFTestDataSamples.OpenSampleWorkbook(file) :
                wb = HSSFTestDataSamples.OpenSampleWorkbook(file);

            for (int si = 0; si < wb.NumberOfSheets; si++)
            {
                ISheet sh = wb.GetSheetAt(si);
                ClassicAssert.IsNotNull(sh.SheetName);
                ClassicAssert.AreEqual(topRows[si], sh.LeftCol, "Did not match for sheet " + si);
            }

            Assert.Warn("test about SXSSFWorkbook was commented");
            // for XSSF also test with SXSSF
            if (isXSSF)
            {
                IWorkbook swb = new SXSSFWorkbook((XSSFWorkbook)wb);
                for (int si = 0; si < swb.NumberOfSheets; si++)
                {
                    ISheet sh = swb.GetSheetAt(si);
                    ClassicAssert.IsNotNull(sh.SheetName);
                    ClassicAssert.AreEqual(topRows[si], sh.LeftCol, "Did not match for sheet " + si);
                }

                swb.Close();
            }





            wb.Close();
        }

        [Test]
        public void Bug55745()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("55745.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            List<XSSFTable> tables = sheet.GetTables();
            /*System.out.println(tables.size());

            for(XSSFTable table : tables) {
                System.out.println("XPath: " + table.GetCommonXpath());
                System.out.println("Name: " + table.GetName());
                System.out.println("Mapped Cols: " + table.GetNumerOfMappedColumns());
                System.out.println("Rowcount: " + table.GetRowCount());
                System.out.println("End Cell: " + table.GetEndCellReference());
                System.out.println("Start Cell: " + table.GetStartCellReference());
            }*/
            ClassicAssert.AreEqual(8, tables.Count, "Sheet should contain 8 tables");
            ClassicAssert.IsNotNull(sheet.GetCommentsTable(false), "Sheet should contain a comments table");

            wb.Close();
        }

        [Test]
        public void Bug55723b()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            // stored with a special name
            ClassicAssert.IsNull(wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, 0));

            CellRangeAddress range = CellRangeAddress.ValueOf("A:B");
            IAutoFilter filter = sheet.SetAutoFilter(range);
            ClassicAssert.IsNotNull(filter);

            // stored with a special name
            XSSFName name = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, 0);
            ClassicAssert.IsNotNull(name);
            ClassicAssert.AreEqual("Sheet0!$A:$B", name.RefersToFormula);

            range = CellRangeAddress.ValueOf("B:C");
            filter = sheet.SetAutoFilter(range);
            ClassicAssert.IsNotNull(filter);

            // stored with a special name
            name = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, 0);
            ClassicAssert.IsNotNull(name);
            ClassicAssert.AreEqual("Sheet0!$B:$C", name.RefersToFormula);

            wb.Close();
        }

        [Test]
        public void Bug51585()
        {
            XSSFTestDataSamples.OpenSampleWorkbook("51585.xlsx");
        }

        private XSSFWorkbook SetupSheet()
        {
            //set up workbook
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

            IRow row1 = sheet.CreateRow(0);
            ICell cell = row1.CreateCell(0);
            cell.SetCellValue("Names");
            ICell cell2 = row1.CreateCell(1);
            cell2.SetCellValue("#");

            IRow row2 = sheet.CreateRow(1);
            ICell cell3 = row2.CreateCell(0);
            cell3.SetCellValue("Jane");
            ICell cell4 = row2.CreateCell(1);
            cell4.SetCellValue(3);

            IRow row3 = sheet.CreateRow(2);
            ICell cell5 = row3.CreateCell(0);
            cell5.SetCellValue("John");
            ICell cell6 = row3.CreateCell(1);
            cell6.SetCellValue(3);

            return wb;        }

        [Test]
        public void TestCreateTwoPivotTablesInOneSheet()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            ClassicAssert.IsNotNull(wb);
            ClassicAssert.IsNotNull(sheet);
            XSSFPivotTable pivotTable = sheet.CreatePivotTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007), new CellReference("H5"));
            ClassicAssert.IsNotNull(pivotTable);
            ClassicAssert.IsTrue(wb.PivotTables.Count > 0);
            XSSFPivotTable pivotTable2 = sheet.CreatePivotTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007), new CellReference("L5"), sheet);
            ClassicAssert.IsNotNull(pivotTable2);
            ClassicAssert.IsTrue(wb.PivotTables.Count > 1);

            wb.Close();
        }

        [Test]
        public void TestCreateTwoPivotTablesInTwoSheets()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            ClassicAssert.IsNotNull(wb);
            ClassicAssert.IsNotNull(sheet);
            XSSFPivotTable pivotTable = sheet.CreatePivotTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007), new CellReference("H5"));
            ClassicAssert.IsNotNull(pivotTable);
            ClassicAssert.IsTrue(wb.PivotTables.Count > 0);
            ClassicAssert.IsNotNull(wb);
            XSSFSheet sheet2 = wb.CreateSheet() as XSSFSheet;
            XSSFPivotTable pivotTable2 = sheet2.CreatePivotTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007), new CellReference("H5"), sheet);
            ClassicAssert.IsNotNull(pivotTable2);
            ClassicAssert.IsTrue(wb.PivotTables.Count > 1);

            wb.Close();
        }

        [Test]
        public void TestCreatePivotTable()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            ClassicAssert.IsNotNull(wb);
            ClassicAssert.IsNotNull(sheet);
            XSSFPivotTable pivotTable = sheet.CreatePivotTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007), new CellReference("H5"));
            ClassicAssert.IsNotNull(pivotTable);
            ClassicAssert.IsTrue(wb.PivotTables.Count > 0);

            wb.Close();
        }

        [Test]
        public void TestCreatePivotTableInOtherSheetThanDataSheet()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet1 = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sheet2 = wb.CreateSheet() as XSSFSheet;

            XSSFPivotTable pivotTable = sheet2.CreatePivotTable
                    (new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007), new CellReference("H5"), sheet1);
            ClassicAssert.AreEqual(0, pivotTable.GetRowLabelColumns().Count);

            ClassicAssert.AreEqual(1, wb.PivotTables.Count);
            ClassicAssert.AreEqual(0, sheet1.GetPivotTables().Count);
            ClassicAssert.AreEqual(1, sheet2.GetPivotTables().Count);

            wb.Close();
        }

        [Test]
        public void TestCreatePivotTableInOtherSheetThanDataSheetUsingAreaReference()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sheet2 = wb.CreateSheet("TEST") as XSSFSheet;

            XSSFPivotTable pivotTable = sheet2.CreatePivotTable(
                new AreaReference(sheet.SheetName + "!A$1:B$2", SpreadsheetVersion.EXCEL2007),
                new CellReference("H5"));
            ClassicAssert.AreEqual(0, pivotTable.GetRowLabelColumns().Count);

            wb.Close();
        }

        [Test]
        public void TestCreatePivotTableWithConflictingDataSheets()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sheet2 = wb.CreateSheet("TEST") as XSSFSheet;

            Assert.Throws<ArgumentException>(() =>
            {
                sheet2.CreatePivotTable(
                    new AreaReference(sheet.SheetName + "!A$1:B$2", SpreadsheetVersion.EXCEL2007),
                    new CellReference("H5"),
                    sheet2);
            });
            wb.Close();
        }

        [Test]
        public void TestReadFails()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

            Assert.Throws<POIXMLException>(() => { sheet.OnDocumentRead(); });

            wb.Close();
        }
        /** 
         * This would be better off as a testable example rather than a simple unit test
         * since Sheet.createComment() was deprecated and removed.
         * https://poi.apache.org/spreadsheet/quick-guide.html#CellComments
         * Feel free to relocated or delete this unit test if it doesn't belong here.
         */
        [Test]
        public void TestCreateComment()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            IClientAnchor anchor = wb.GetCreationHelper().CreateClientAnchor();

            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFComment comment = sheet.CreateDrawingPatriarch().CreateCellComment(anchor) as XSSFComment;
            ClassicAssert.IsNotNull(comment);

            wb.Close();
        }

        protected void testCopyOneRow(String copyRowsTestWorkbook)
        {
            double FLOAT_PRECISION = 1e-9;
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook(copyRowsTestWorkbook);
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            CellCopyPolicy defaultCopyPolicy = new CellCopyPolicy();
            sheet.CopyRows(1, 1, 6, defaultCopyPolicy);
            IRow srcRow = sheet.GetRow(1);
            IRow destRow = sheet.GetRow(6);
            int col = 0;
            ICell cell;
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual("Source row ->", cell.StringCellValue);
            // Style
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual("Red", cell.StringCellValue, "[Style] B7 cell value");
            ClassicAssert.AreEqual(CellUtil.GetCell(srcRow, 1).CellStyle, cell.CellStyle, "[Style] B7 cell style");
            // Blank
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Blank, cell.CellType, "[Blank] C7 cell type");
            // Error
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Error, cell.CellType, "[Error] D7 cell type");
            FormulaError error = FormulaError.ForInt(cell.ErrorCellValue);
            ClassicAssert.AreEqual(FormulaError.NA, error, "[Error] D7 cell value"); //FIXME: XSSFCell and HSSFCell expose different interfaces. getErrorCellString would be helpful here
                                                                              // Date
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Date] E7 cell type");
            DateTime date = new DateTime(2000, 1, 1);
            ClassicAssert.AreEqual(date, cell.DateCellValue, "[Date] E7 cell value");
            // Boolean
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Boolean, cell.CellType, "[Boolean] F7 cell type");
            ClassicAssert.IsTrue(cell.BooleanCellValue, "[Boolean] F7 cell value");
            // String
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.String, cell.CellType, "[String] G7 cell type");
            ClassicAssert.AreEqual("Hello", cell.StringCellValue, "[String] G7 cell value");

            // Int
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Int] H7 cell type");
            ClassicAssert.AreEqual(15, (int)cell.NumericCellValue, "[Int] H7 cell value");

            // Float
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Float] I7 cell type");
            ClassicAssert.AreEqual(12.5, cell.NumericCellValue, FLOAT_PRECISION, "[Float] I7 cell value");

            // Cell Formula
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual("J7", new CellReference(cell).FormatAsString());
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType, "[Cell Formula] J7 cell type");
            ClassicAssert.AreEqual("5+2", cell.CellFormula, "[Cell Formula] J7 cell formula");
            Console.WriteLine("Cell formula evaluation currently unsupported");

            // Cell Formula with Reference
            // Formula row references should be adjusted by destRowNum-srcRowNum
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual("K7", new CellReference(cell).FormatAsString());
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Reference] K7 cell type");
            ClassicAssert.AreEqual("J7+H$2", cell.CellFormula,
                "[Cell Formula with Reference] K7 cell formula");

            // Cell Formula with Reference spanning multiple rows
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Reference spanning multiple rows] L7 cell type");
            ClassicAssert.AreEqual("G7&\" \"&G8", cell.CellFormula,
                "[Cell Formula with Reference spanning multiple rows] L7 cell formula");

            // Cell Formula with Reference spanning multiple rows
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Area Reference] M7 cell type");
            ClassicAssert.AreEqual("SUM(H7:I8)", cell.CellFormula,
                "[Cell Formula with Area Reference] M7 cell formula");

            // Array Formula
            cell = CellUtil.GetCell(destRow, col++);
            Console.WriteLine("Array formulas currently unsupported");
            // FIXME: Array Formula set with Sheet.setArrayFormula() instead of cell.setFormula()
            /*
            ClassicAssert.AreEqual("[Array Formula] N7 cell type", CellType.Formula, cell.CellType);
            ClassicAssert.AreEqual("[Array Formula] N7 cell formula", "{SUM(H7:J7*{1,2,3})}", cell.CellFormula);
            */

            // Data Format
            cell = CellUtil.GetCell(destRow, col++);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Data Format] O7 cell type;");
            ClassicAssert.AreEqual(100.20, cell.NumericCellValue, FLOAT_PRECISION, "[Data Format] O7 cell value");
            //FIXME: currently Assert.Fails
            String moneyFormat = "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)";
            ClassicAssert.AreEqual(moneyFormat, cell.CellStyle.GetDataFormatString(), "[Data Format] O7 data format");

            // Merged
            cell = CellUtil.GetCell(destRow, col);
            ClassicAssert.AreEqual("Merged cells", cell.StringCellValue,
                "[Merged] P7:Q7 cell value");
            ClassicAssert.IsTrue(sheet.MergedRegions.Contains(CellRangeAddress.ValueOf("P7:Q7")),
                "[Merged] P7:Q7 merged region");

            // Merged across multiple rows
            // Microsoft Excel 2013 does not copy a merged region unless all rows of
            // the source merged region are selected
            // POI's behavior should match this behavior
            col += 2;
            cell = CellUtil.GetCell(destRow, col);
            // Note: this behavior deviates from Microsoft Excel,
            // which will not overwrite a cell in destination row if merged region extends beyond the copied row.
            // The Excel way would require:
            //ClassicAssert.AreEqual("[Merged across multiple rows] R7:S8 merged region", "Should NOT be overwritten", cell.StringCellValue);
            //ClassicAssert.IsFalse("[Merged across multiple rows] R7:S8 merged region", 
            //        sheet.MergedRegions.contains(CellRangeAddress.valueOf("R7:S8")));
            // As currently implemented, cell value is copied but merged region is not copied
            ClassicAssert.AreEqual("Merged cells across multiple rows", cell.StringCellValue,
                "[Merged across multiple rows] R7:S8 cell value");
            ClassicAssert.IsFalse(sheet.MergedRegions.Contains(CellRangeAddress.ValueOf("R7:S7")),
                "[Merged across multiple rows] R7:S7 merged region (one row)"); //shouldn't do 1-row merge
            ClassicAssert.IsFalse(sheet.MergedRegions.Contains(CellRangeAddress.ValueOf("R7:S8")),
                "[Merged across multiple rows] R7:S8 merged region"); //shouldn't do 2-row merge

            // Make sure other rows are blank (off-by-one errors)
            ClassicAssert.IsNull(sheet.GetRow(5));
            ClassicAssert.IsNull(sheet.GetRow(7));

            wb.Close();
        }

        protected void testCopyMultipleRows(String copyRowsTestWorkbook)
        {
            double FLOAT_PRECISION = 1e-9;
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook(copyRowsTestWorkbook);
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            CellCopyPolicy defaultCopyPolicy = new CellCopyPolicy();
            sheet.CopyRows(0, 3, 8, defaultCopyPolicy);
            IRow srcHeaderRow = sheet.GetRow(0);
            IRow srcRow1 = sheet.GetRow(1);
            IRow srcRow2 = sheet.GetRow(2);
            IRow srcRow3 = sheet.GetRow(3);
            IRow destHeaderRow = sheet.GetRow(8);
            IRow destRow1 = sheet.GetRow(9);
            IRow destRow2 = sheet.GetRow(10);
            IRow destRow3 = sheet.GetRow(11);
            int col = 0;
            ICell cell;

            // Header row should be copied
            ClassicAssert.IsNotNull(destHeaderRow);

            // Data rows
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual("Source row ->", cell.StringCellValue);

            // Style
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual("Red", cell.StringCellValue, "[Style] B10 cell value");
            ClassicAssert.AreEqual(CellUtil.GetCell(srcRow1, 1).CellStyle, cell.CellStyle, "[Style] B10 cell style");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual("Blue", cell.StringCellValue, "[Style] B11 cell value");
            ClassicAssert.AreEqual(CellUtil.GetCell(srcRow2, 1).CellStyle, cell.CellStyle, "[Style] B11 cell style");

            // Blank
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Blank, cell.CellType, "[Blank] C10 cell type");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Blank, cell.CellType, "[Blank] C11 cell type");

            // Error
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Error, cell.CellType, "[Error] D10 cell type");
            FormulaError error = FormulaError.ForInt(cell.ErrorCellValue);
            ClassicAssert.AreEqual(FormulaError.NA, error, "[Error] D10 cell value"); //FIXME: XSSFCell and HSSFCell expose different interfaces. getErrorCellString would be helpful here

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Error, cell.CellType, "[Error] D11 cell type");
            error = FormulaError.ForInt(cell.ErrorCellValue);
            ClassicAssert.AreEqual(FormulaError.NAME, error, "[Error] D11 cell value"); //FIXME: XSSFCell and HSSFCell expose different interfaces. getErrorCellString would be helpful here

            // Date
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Date] E10 cell type");
            DateTime date = new DateTime(2000, 1, 1);
            ClassicAssert.AreEqual(date, cell.DateCellValue, "[Date] E10 cell value");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Date] E11 cell type");
            date = new DateTime(2000, 1, 2);
            ClassicAssert.AreEqual(date, cell.DateCellValue, "[Date] E11 cell value");

            // Boolean
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Boolean, cell.CellType, "[Boolean] F10 cell type");
            ClassicAssert.IsTrue(cell.BooleanCellValue, "[Boolean] F10 cell value");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Boolean, cell.CellType, "[Boolean] F11 cell type");
            ClassicAssert.IsFalse(cell.BooleanCellValue, "[Boolean] F11 cell value");

            // String
            col++;            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.String, cell.CellType, "[String] G10 cell type");
            ClassicAssert.AreEqual("Hello", cell.StringCellValue, "[String] G10 cell value");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.String, cell.CellType, "[String] G11 cell type");
            ClassicAssert.AreEqual("World", cell.StringCellValue, "[String] G11 cell value");

            // Int
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Int] H10 cell type");
            ClassicAssert.AreEqual(15, (int)cell.NumericCellValue, "[Int] H10 cell value");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Int] H11 cell type");
            ClassicAssert.AreEqual(42, (int)cell.NumericCellValue, "[Int] H11 cell value");

            // Float
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Float] I10 cell type");
            ClassicAssert.AreEqual(12.5, cell.NumericCellValue, FLOAT_PRECISION, "[Float] I10 cell value");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Float] I11 cell type");
            ClassicAssert.AreEqual(5.5, cell.NumericCellValue, FLOAT_PRECISION, "[Float] I11 cell value");

            // Cell Formula
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType, "[Cell Formula] J10 cell type");
            ClassicAssert.AreEqual("5+2", cell.CellFormula, "[Cell Formula] J10 cell formula");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType, "[Cell Formula] J11 cell type");
            ClassicAssert.AreEqual("6+18", cell.CellFormula, "[Cell Formula] J11 cell formula");
            // Cell Formula with Reference
            col++;
            // Formula row references should be adjusted by destRowNum-srcRowNum
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Reference] K10 cell type");
            ClassicAssert.AreEqual("J10+H$2", cell.CellFormula,
                "[Cell Formula with Reference] K10 cell formula");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Reference] K11 cell type");
            ClassicAssert.AreEqual("J11+H$2", cell.CellFormula, "[Cell Formula with Reference] K11 cell formula");

            // Cell Formula with Reference spanning multiple rows
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Reference spanning multiple rows] L10 cell type");
            ClassicAssert.AreEqual("G10&\" \"&G11", cell.CellFormula,
                "[Cell Formula with Reference spanning multiple rows] L10 cell formula");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Reference spanning multiple rows] L11 cell type");
            ClassicAssert.AreEqual("G11&\" \"&G12", cell.CellFormula,
                "[Cell Formula with Reference spanning multiple rows] L11 cell formula");

            // Cell Formula with Area Reference
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Area Reference] M10 cell type");
            ClassicAssert.AreEqual("SUM(H10:I11)", cell.CellFormula,
                "[Cell Formula with Area Reference] M10 cell formula");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Formula, cell.CellType,
                "[Cell Formula with Area Reference] M11 cell type");
            ClassicAssert.AreEqual("SUM($H$3:I10)", cell.CellFormula,
                "[Cell Formula with Area Reference] M11 cell formula"); //Also acceptable: SUM($H10:I$3), but this AreaReference isn't in ascending order

            // Array Formula
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            //Console.WriteLine("Array formulas currently unsupported");
            /*
                // FIXME: Array Formula set with Sheet.setArrayFormula() instead of cell.setFormula()
                ClassicAssert.AreEqual("[Array Formula] N10 cell type", CellType.Formula, cell.CellType);
                ClassicAssert.AreEqual("[Array Formula] N10 cell formula", "{SUM(H10:J10*{1,2,3})}", cell.CellFormula);

                cell = CellUtil.GetCell(destRow2, col);
                // FIXME: Array Formula set with Sheet.setArrayFormula() instead of cell.setFormula() 
                ClassicAssert.AreEqual("[Array Formula] N11 cell type", CellType.Formula, cell.CellType);
                ClassicAssert.AreEqual("[Array Formula] N11 cell formula", "{SUM(H11:J11*{1,2,3})}", cell.CellFormula);
             */

            // Data Format
            col++;
            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual(CellType.Numeric, cell.CellType, "[Data Format] O10 cell type");
            ClassicAssert.AreEqual(100.20, cell.NumericCellValue, FLOAT_PRECISION, "[Data Format] O10 cell value");
            String moneyFormat = "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)";
            ClassicAssert.AreEqual(moneyFormat, cell.CellStyle.GetDataFormatString(), "[Data Format] O10 cell data format");

            // Merged
            col++;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual("Merged cells", cell.StringCellValue, "[Merged] P10:Q10 cell value");
            ClassicAssert.IsTrue(sheet.MergedRegions.Contains(CellRangeAddress.ValueOf("P10:Q10")),
                "[Merged] P10:Q10 merged region");

            cell = CellUtil.GetCell(destRow2, col);
            ClassicAssert.AreEqual("Merged cells", cell.StringCellValue,
                "[Merged] P11:Q11 cell value");
            ClassicAssert.IsTrue(sheet.MergedRegions.Contains(CellRangeAddress.ValueOf("P11:Q11")),
                "[Merged] P11:Q11 merged region");

            // Should Q10/Q11 be checked?

            // Merged across multiple rows
            // Microsoft Excel 2013 does not copy a merged region unless all rows of
            // the source merged region are selected
            // POI's behavior should match this behavior
            col += 2;
            cell = CellUtil.GetCell(destRow1, col);
            ClassicAssert.AreEqual("Merged cells across multiple rows", cell.StringCellValue,
                "[Merged across multiple rows] R10:S11 cell value");
            ClassicAssert.IsTrue(sheet.MergedRegions.Contains(CellRangeAddress.ValueOf("R10:S11")),
                "[Merged across multiple rows] R10:S11 merged region");

            // Row 3 (zero-based) was empty, so Row 11 (zero-based) should be empty too.
            if (srcRow3 == null)
            {
                ClassicAssert.IsNull(destRow3, "Row 3 was empty, so Row 11 should be empty");
            }

            // Make sure other rows are blank (off-by-one errors)
            ClassicAssert.IsNull(sheet.GetRow(7), "Off-by-one lower edge case"); //one row above destHeaderRow
            ClassicAssert.IsNull(sheet.GetRow(12), "Off-by-one upper edge case"); //one row below destRow3

            wb.Close();
        }

        [Test]
        public void TestCopyOneRow()
        {
            testCopyOneRow("XSSFSheet.copyRows.xlsx");
        }

        [Test]
        public void TestCopyMultipleRows()
        {
            testCopyMultipleRows("XSSFSheet.copyRows.xlsx");
        }

        [Test]
        public void TestIgnoredErrors()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.CreateSheet() as XSSFSheet;
            CellRangeAddress region = CellRangeAddress.ValueOf("B2:D4");
            sheet.AddIgnoredErrors(region, IgnoredErrorType.NumberStoredAsText);
            CT_IgnoredError ignoredError = sheet.GetCTWorksheet().ignoredErrors.GetIgnoredErrorArray(0);
            ClassicAssert.AreEqual(1, ignoredError.sqref.Count);
            ClassicAssert.AreEqual("B2:D4", ignoredError.sqref[0]);
            ClassicAssert.IsTrue(ignoredError.numberStoredAsText);

            Dictionary<IgnoredErrorType, ISet<CellRangeAddress>> ignoredErrors = sheet.GetIgnoredErrors();
            ClassicAssert.AreEqual(1, ignoredErrors.Count);
            ClassicAssert.AreEqual(1, ignoredErrors[IgnoredErrorType.NumberStoredAsText].Count);
            IEnumerator<CellRangeAddress> it = ignoredErrors[IgnoredErrorType.NumberStoredAsText].GetEnumerator();
            it.MoveNext();
            ClassicAssert.AreEqual("B2:D4", it.Current.FormatAsString());

            workbook.Close();
        }
        [Test]
        public void TestIgnoredErrorsMultipleTypes()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.CreateSheet() as XSSFSheet;
            CellRangeAddress region = CellRangeAddress.ValueOf("B2:D4");
            sheet.AddIgnoredErrors(region, IgnoredErrorType.Formula, IgnoredErrorType.EvaluationError);
            CT_IgnoredError ignoredError = sheet.GetCTWorksheet().ignoredErrors.GetIgnoredErrorArray(0);
            ClassicAssert.AreEqual(1, ignoredError.sqref.Count);
            ClassicAssert.AreEqual("B2:D4", ignoredError.sqref[0]);
            ClassicAssert.IsFalse(ignoredError.numberStoredAsText);
            ClassicAssert.IsTrue(ignoredError.formula);
            ClassicAssert.IsTrue(ignoredError.evalError);

            Dictionary<IgnoredErrorType, ISet<CellRangeAddress>> ignoredErrors = sheet.GetIgnoredErrors();
            ClassicAssert.AreEqual(2, ignoredErrors.Count);
            ClassicAssert.AreEqual(1, ignoredErrors[IgnoredErrorType.Formula].Count);
            IEnumerator<CellRangeAddress> it = ignoredErrors[IgnoredErrorType.Formula].GetEnumerator();
            it.MoveNext();
            ClassicAssert.AreEqual("B2:D4", it.Current.FormatAsString());
            ClassicAssert.AreEqual(1, ignoredErrors[IgnoredErrorType.EvaluationError].Count);
            it = ignoredErrors[IgnoredErrorType.EvaluationError].GetEnumerator();
            it.MoveNext();
            ClassicAssert.AreEqual("B2:D4", it.Current.FormatAsString());
            workbook.Close();
        }
        [Test]
        public void TestIgnoredErrorsMultipleCalls()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.CreateSheet() as XSSFSheet;
            CellRangeAddress region = CellRangeAddress.ValueOf("B2:D4");
            // Two calls means two elements, no clever collapsing just yet.
            sheet.AddIgnoredErrors(region, IgnoredErrorType.EvaluationError);
            sheet.AddIgnoredErrors(region, IgnoredErrorType.Formula);

            CT_IgnoredError ignoredError = sheet.GetCTWorksheet().ignoredErrors.GetIgnoredErrorArray(0);
            ClassicAssert.AreEqual(1, ignoredError.sqref.Count);
            ClassicAssert.AreEqual("B2:D4", ignoredError.sqref[0]);
            ClassicAssert.IsFalse(ignoredError.formula);
            ClassicAssert.IsTrue(ignoredError.evalError);

            ignoredError = sheet.GetCTWorksheet().ignoredErrors.GetIgnoredErrorArray(1);
            ClassicAssert.AreEqual(1, ignoredError.sqref.Count);
            ClassicAssert.AreEqual("B2:D4", ignoredError.sqref[0]);
            ClassicAssert.IsTrue(ignoredError.formula);
            ClassicAssert.IsFalse(ignoredError.evalError);

            Dictionary<IgnoredErrorType, ISet<CellRangeAddress>> ignoredErrors = sheet.GetIgnoredErrors();
            ClassicAssert.AreEqual(2, ignoredErrors.Count);
            ClassicAssert.AreEqual(1, ignoredErrors[IgnoredErrorType.Formula].Count);
            IEnumerator<CellRangeAddress> it = ignoredErrors[IgnoredErrorType.Formula].GetEnumerator();
            it.MoveNext();
            ClassicAssert.AreEqual("B2:D4", it.Current.FormatAsString());
            ClassicAssert.AreEqual(1, ignoredErrors[IgnoredErrorType.EvaluationError].Count);
            it = ignoredErrors[IgnoredErrorType.EvaluationError].GetEnumerator();
            it.MoveNext();
            ClassicAssert.AreEqual("B2:D4", it.Current.FormatAsString());
            workbook.Close();
        }

        [Test]
        public void SetTabColor()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sh = wb.CreateSheet() as XSSFSheet;
                ClassicAssert.IsTrue(sh.GetCTWorksheet().sheetPr == null || !sh.GetCTWorksheet().sheetPr.IsSetTabColor());
                sh.TabColor = new XSSFColor(IndexedColors.Red, null);
                ClassicAssert.IsTrue(sh.GetCTWorksheet().sheetPr.IsSetTabColor());
                ClassicAssert.AreEqual(IndexedColors.Red.Index,
                        sh.GetCTWorksheet().sheetPr.tabColor.indexed);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void GetTabColor()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sh = wb.CreateSheet() as XSSFSheet;
                ClassicAssert.IsTrue(sh.GetCTWorksheet().sheetPr == null || !sh.GetCTWorksheet().sheetPr.IsSetTabColor());
                ClassicAssert.IsNull(sh.TabColor);
                sh.TabColor = new XSSFColor(IndexedColors.Red, null);
                XSSFColor expected = new XSSFColor(IndexedColors.Red, null);
                ClassicAssert.AreEqual(expected, sh.TabColor);
            }
            finally
            {
                wb.Close();
            }
        }

        // Test using an existing workbook saved by Excel
        [Test]
        public void TabColor()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("SheetTabColors.xlsx");
            try
            {
                // non-colored sheets do not have a color
                ClassicAssert.IsNull((wb.GetSheet("default") as XSSFSheet).TabColor);

                // test indexed-colored sheet
                XSSFColor expected = new XSSFColor(IndexedColors.Red, null);
                ClassicAssert.AreEqual(expected, (wb.GetSheet("indexedRed") as XSSFSheet).TabColor);

                // test regular-colored (non-indexed, ARGB) sheet
                expected = new XSSFColor
                {
                    ARGBHex = "FF7F2700"
                };
                ClassicAssert.AreEqual(expected, (wb.GetSheet("customOrange") as XSSFSheet).TabColor);
            }
            finally
            {
                wb.Close();
            }
        }

        /**
         * See bug #52425
         */
        [Test]
        public void TestInsertCommentsToClonedSheet()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("52425.xlsx");
            ICreationHelper helper = wb.GetCreationHelper();
            ISheet sheet2 = wb.CreateSheet("Sheet 2");
            ISheet sheet3 = wb.CloneSheet(0);
            wb.SetSheetName(2, "Sheet 3");
            // Adding Comment to new created Sheet 2
            AddComments(helper, sheet2);
            // Adding Comment to cloned Sheet 3
            AddComments(helper, sheet3);
        }

        private void AddComments(ICreationHelper helper, ISheet sheet)
        {
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            for (int i = 0; i < 2; i++)
            {
                IClientAnchor anchor = helper.CreateClientAnchor();
                anchor.Col1 = 0;
                anchor.Row1 = 0 + i;
                anchor.Col2 = 2;
                anchor.Row2 = 3 + i;
                IComment comment = drawing.CreateCellComment(anchor);
                comment.String = helper.CreateRichTextString("BugTesting");
                IRow row = sheet.GetRow(0 + i);
                if (row == null)
                {
                    row = sheet.CreateRow(0 + i);
                }
                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                }
                cell.CellComment = comment;
            }
        }


        // bug 59687:  XSSFSheet.RemoveRow doesn't handle row gaps properly when removing row comments
        [Test]
        public void TestRemoveRowWithCommentAndGapAbove()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("59687.xlsx");
            ISheet sheet = wb.GetSheetAt(0);

            // comment exists
            CellAddress commentCellAddress = new CellAddress("A4");
            ClassicAssert.IsNotNull(sheet.GetCellComment(commentCellAddress));

            ClassicAssert.AreEqual(1, sheet.GetCellComments().Count, "Wrong starting # of comments");

            sheet.RemoveRow(sheet.GetRow(commentCellAddress.Row));

            ClassicAssert.AreEqual(0, sheet.GetCellComments().Count, "There should not be any comments left!");
        }

        //[Test]
        //public void TestCoordinate()
        //{
        //    Console.WriteLine("=====Caution=====");
        //    Console.WriteLine("Results may not be as expected if the default font of the original excel file is not installed in the OS");
        //    XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("coordinate.xlsm");
        //    XSSFSheet sheet = (XSSFSheet)wb.GetSheet("Sheet1");

        //    XSSFDrawing drawing = sheet.GetDrawingPatriarch();

        //    List<XSSFShape> shapes = drawing.GetShapes();
        //    XSSFClientAnchor anchor;

        //    foreach (var shape in shapes)
        //    {
        //        XSSFClientAnchor sa = (XSSFClientAnchor)shape.anchor;
        //        switch (shape.Name)
        //        {
        //            case "cxn1":
        //                anchor = new XSSFClientAnchor(sheet, Units.ToEMU(50), Units.ToEMU(75), Units.ToEMU(125), Units.ToEMU(150));
        //                break;
        //            case "cxn2":
        //                anchor = new XSSFClientAnchor(sheet, Units.ToEMU(75), Units.ToEMU(75), Units.ToEMU(150), Units.ToEMU(150));
        //                break;
        //            case "cxn3":
        //                anchor = new XSSFClientAnchor(sheet, Units.ToEMU(150), Units.ToEMU(75), Units.ToEMU(225), Units.ToEMU(150));
        //                break;
        //            case "cxn4":
        //                anchor = new XSSFClientAnchor(sheet, Units.ToEMU(225), Units.ToEMU(75), Units.ToEMU(300), Units.ToEMU(150));
        //                break;
        //            default:
        //                ClassicAssert.True(false, "Unexpected shape {0}", new object[] { shape.Name });
        //                return;
        //        }
        //        ClassicAssert.IsTrue(sa.From.col == anchor.From.col,       /**/"From.col   [{0}]({1}={2})", new object[] { shape.Name, sa.From.col, anchor.From.col });
        //        ClassicAssert.IsTrue(sa.From.colOff == anchor.From.colOff,  /**/"From.colOff[{0}]({1}={2})", new object[] { shape.Name, sa.From.colOff, anchor.From.colOff });
        //        ClassicAssert.IsTrue(sa.From.row == anchor.From.row,       /**/"From.row   [{0}]({1}={2})", new object[] { shape.Name, sa.From.row, anchor.From.row });
        //        ClassicAssert.IsTrue(sa.From.rowOff == anchor.From.rowOff, /**/"From.rowOff[{0}]({1}={2})", new object[] { shape.Name, sa.From.rowOff, anchor.From.rowOff });
        //        ClassicAssert.IsTrue(sa.To.col == anchor.To.col,           /**/"To.col     [{0}]({1}={2})", new object[] { shape.Name, sa.To.col, anchor.To.col });
        //        ClassicAssert.IsTrue(sa.To.colOff == anchor.To.colOff,     /**/"To.colOff  [{0}]({1}={2})", new object[] { shape.Name, sa.To.colOff, anchor.To.colOff });
        //        ClassicAssert.IsTrue(sa.To.row == anchor.To.row,           /**/"To.row     [{0}]({1}={2})", new object[] { shape.Name, sa.To.row, anchor.To.row });
        //        ClassicAssert.IsTrue(sa.To.rowOff == anchor.To.rowOff,    /**/"To.rowOff  [{0}]({1}={2})", new object[] { shape.Name, sa.To.rowOff, anchor.To.rowOff });
        //    }
        //}

        [Test]
        public void TestDefaultColumnWidth()
        {
            using (var book = new XSSFWorkbook())
            {
                var sheet = book.CreateSheet("Sheet1");
                var row = sheet.CreateRow(1);

                var cell = row.CreateCell(0);

                ClassicAssert.AreEqual(sheet.GetColumnWidth(0) / 256, sheet.DefaultColumnWidth);
                ClassicAssert.AreEqual(sheet.DefaultColumnWidth, 8.43);

                sheet.DefaultColumnWidth = 50.1;
                ClassicAssert.AreEqual(sheet.GetColumnWidth(0) / 256, sheet.DefaultColumnWidth);

                sheet.SetColumnWidth(0, 100);
                ClassicAssert.AreEqual(sheet.GetColumnWidth(0), 100);
                ClassicAssert.AreNotEqual(sheet.GetColumnWidth(0) / 256, sheet.DefaultColumnWidth);
            }
        }


        [Test]
        public void TestCopyRepeatingRowsAndColumns()
        {
            using (var book = new XSSFWorkbook())
            {
                var sheet = book.CreateSheet("Sheet1");
                
                var row1 = sheet.CreateRow(0);
                row1.CreateCell(0);

                var row2 = sheet.CreateRow(1);
                row2.CreateCell(0);

                sheet.RepeatingRows = CellRangeAddress.ValueOf("1:1");
                sheet.RepeatingColumns = CellRangeAddress.ValueOf("A1:B1");

                var clonedSheet = book.CloneSheet(0);

                ClassicAssert.IsNotNull(clonedSheet.RepeatingRows, "RepeatingRows is null");
                ClassicAssert.AreEqual(clonedSheet.RepeatingRows.FirstRow, sheet.RepeatingRows.FirstRow, "RepeatingRows.FirstRow are not equal");
                ClassicAssert.AreEqual(clonedSheet.RepeatingRows.LastRow, sheet.RepeatingRows.LastRow, "RepeatingRows.LastRow are not equal");

                ClassicAssert.IsNotNull(clonedSheet.RepeatingColumns, "RepeatingColumns is null");
                ClassicAssert.AreEqual(clonedSheet.RepeatingColumns.FirstColumn, sheet.RepeatingColumns.FirstColumn, "RepeatingColumns.FirstColumn are not equal");
                ClassicAssert.AreEqual(clonedSheet.RepeatingColumns.LastColumn, sheet.RepeatingColumns.LastColumn, "RepeatingColumns.LastColumn are not equal");

            }
        }
    }
}
