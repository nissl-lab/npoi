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

using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;

namespace TestCases.XSSF.UserModel
{

    /**
     * Tests for {@link XSSFPrintSetup}
     */
    [TestFixture]
    public class TestXSSFPrintSetup
    {
        [Test]
        public void TestSetGetPaperSize()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.paperSize = (9);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(PaperSize.A4, printSetup.GetPaperSizeEnum());
            ClassicAssert.AreEqual(9, printSetup.PaperSize);

            printSetup.SetPaperSize(PaperSize.A3);
            ClassicAssert.AreEqual((uint)8, pSetup.paperSize);
        }

        [Test]
        public void TestSetGetScale()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.scale = (uint)9;
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(9, printSetup.Scale);

            printSetup.Scale = ((short)100);
            ClassicAssert.AreEqual((uint)100, pSetup.scale);
        }
        [Test]
        public void TestSetGetPageStart()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.firstPageNumber = 9;
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(9, printSetup.PageStart);

            printSetup.PageStart = ((short)1);
            ClassicAssert.AreEqual((uint)1, pSetup.firstPageNumber);
        }

        [Test]
        public void TestSetGetFitWidthHeight()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.fitToWidth = (50);
            pSetup.fitToHeight = (99);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(50, printSetup.FitWidth);
            ClassicAssert.AreEqual(99, printSetup.FitHeight);

            printSetup.FitWidth = ((short)66);
            printSetup.FitHeight = ((short)80);
            ClassicAssert.AreEqual((uint)66, pSetup.fitToWidth);
            ClassicAssert.AreEqual((uint)80, pSetup.fitToHeight);

        }
        [Test]
        public void TestSetGetLeftToRight()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.pageOrder = (ST_PageOrder.downThenOver);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(false, printSetup.LeftToRight);

            printSetup.LeftToRight = (true);
            ClassicAssert.AreEqual(PageOrder.OVER_THEN_DOWN.Value, (int)pSetup.pageOrder);
        }
        [Test]
        public void TestSetGetOrientation()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.orientation = (ST_Orientation.portrait);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(PrintOrientation.PORTRAIT, printSetup.Orientation);
            ClassicAssert.AreEqual(false, printSetup.Landscape);
            ClassicAssert.AreEqual(false, printSetup.NoOrientation);

            printSetup.Orientation = (PrintOrientation.LANDSCAPE);
            ClassicAssert.AreEqual((int)pSetup.orientation, printSetup.Orientation.Value);
            ClassicAssert.AreEqual(true, printSetup.Landscape);
            ClassicAssert.AreEqual(false, printSetup.NoOrientation);
        }

        [Test]
        public void TestSetGetValidSettings()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.usePrinterDefaults = (false);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(false, printSetup.ValidSettings);

            printSetup.ValidSettings = (true);
            ClassicAssert.AreEqual(true, pSetup.usePrinterDefaults);
        }
        [Test]
        public void TestSetGetNoColor()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.blackAndWhite = (false);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(false, printSetup.NoColor);

            printSetup.NoColor = true;
            ClassicAssert.AreEqual(true, pSetup.blackAndWhite);
        }
        [Test]
        public void TestSetGetDraft()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.draft = (false);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(false, printSetup.Draft);

            printSetup.Draft = (true);
            ClassicAssert.AreEqual(true, pSetup.draft);
        }
        [Test]
        public void TestSetGetNotes()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.cellComments = ST_CellComments.none;
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(false, printSetup.Notes);

            printSetup.Notes = true;
            ClassicAssert.AreEqual(PrintCellComments.AS_DISPLAYED.Value, (int)pSetup.cellComments);
        }

        [Test]
        public void TestSetGetUsePage()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.useFirstPageNumber = (false);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(false, printSetup.UsePage);

            printSetup.UsePage = (true);
            ClassicAssert.AreEqual(true, pSetup.useFirstPageNumber);
        }
        [Test]
        public void TestSetGetHVResolution()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.horizontalDpi = (120);
            pSetup.verticalDpi = (100);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(120, printSetup.HResolution);
            ClassicAssert.AreEqual(100, printSetup.VResolution);

            printSetup.HResolution = ((short)150);
            printSetup.VResolution = ((short)130);
            ClassicAssert.AreEqual((uint)150, pSetup.horizontalDpi);
            ClassicAssert.AreEqual((uint)130, pSetup.verticalDpi);
        }
        [Test]
        public void TestSetGetHeaderFooterMargin()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageMargins pMargins = worksheet.AddNewPageMargins();
            pMargins.header = (1.5);
            pMargins.footer = (2);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(1.5, printSetup.HeaderMargin, 0.0);
            ClassicAssert.AreEqual(2.0, printSetup.FooterMargin, 0.0);

            printSetup.HeaderMargin = (5);
            printSetup.FooterMargin = (3.5);
            ClassicAssert.AreEqual(5.0, pMargins.header, 0.0);
            ClassicAssert.AreEqual(3.5, pMargins.footer, 0.0);
        }
        [Test]
        public void TestSetGetCopies()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_PageSetup pSetup = worksheet.AddNewPageSetup();
            pSetup.copies = (9);
            XSSFPrintSetup printSetup = new XSSFPrintSetup(worksheet);
            ClassicAssert.AreEqual(9, printSetup.Copies);

            printSetup.Copies = (short)15;
            ClassicAssert.AreEqual((uint)15, pSetup.copies);
        }
        [Test]
        public void TestSetSaveRead()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet s1 = (XSSFSheet)wb.CreateSheet();
            ClassicAssert.AreEqual(false, s1.GetCTWorksheet().IsSetPageSetup());
            ClassicAssert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());

            XSSFPrintSetup print = (XSSFPrintSetup)s1.PrintSetup;
            ClassicAssert.AreEqual(true, s1.GetCTWorksheet().IsSetPageSetup());
            ClassicAssert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());

            print.Copies = ((short)3);
            print.Landscape = (true);
            ClassicAssert.AreEqual(3, print.Copies);
            ClassicAssert.AreEqual(true, print.Landscape);

            XSSFSheet s2 = (XSSFSheet)wb.CreateSheet();
            ClassicAssert.AreEqual(false, s2.GetCTWorksheet().IsSetPageSetup());
            ClassicAssert.AreEqual(true, s2.GetCTWorksheet().IsSetPageMargins());

            // Round trip and check
            XSSFWorkbook wbBack = (XSSFWorkbook)XSSFITestDataProvider.instance.WriteOutAndReadBack(wb);

            s1 = (XSSFSheet)wbBack.GetSheetAt(0);
            s2 = (XSSFSheet)wbBack.GetSheetAt(1);

            ClassicAssert.AreEqual(true, s1.GetCTWorksheet().IsSetPageSetup());
            ClassicAssert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());
            ClassicAssert.AreEqual(false, s2.GetCTWorksheet().IsSetPageSetup());
            ClassicAssert.AreEqual(true, s2.GetCTWorksheet().IsSetPageMargins());

            print = (XSSFPrintSetup)s1.PrintSetup;
            ClassicAssert.AreEqual(3, print.Copies);
            ClassicAssert.AreEqual(true, print.Landscape);

            wb.Close();
        }
        [Test]
        public void TestSetLandscapeFalse()
        {
            XSSFPrintSetup ps = new XSSFPrintSetup(new CT_Worksheet());

            ClassicAssert.IsFalse(ps.Landscape);

            ps.Landscape = (true);
            ClassicAssert.IsTrue(ps.Landscape);

            ps.Landscape = (false);
            ClassicAssert.IsFalse(ps.Landscape);
        }

        [Test]
        public void TestSetLeftToRight()
        {
            XSSFPrintSetup ps = new XSSFPrintSetup(new CT_Worksheet());

            ClassicAssert.IsFalse(ps.LeftToRight);

            ps.LeftToRight = (true);
            ClassicAssert.IsTrue(ps.LeftToRight);

            ps.LeftToRight = (false);
            ClassicAssert.IsFalse(ps.LeftToRight);
        }
    }
}
