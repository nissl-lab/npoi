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

using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NPOI.XSSF.UserModel.Helpers;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Model;
using System.Collections.Generic;
using System;
using NPOI.XSSF;
using NPOI.Util;
using NPOI.HSSF.Record;
using TestCases.SS.UserModel;
using TestCases.HSSF;
namespace NPOI.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFSheet : BaseTestSheet
    {

        public TestXSSFSheet()
            : base(XSSFITestDataProvider.instance)
        {

        }

        //[Test]
        //TODO column styles are not yet supported by XSSF
        public override void TestDefaultColumnStyle()
        {
            base.TestDefaultColumnStyle();
        }
        [Test]
        public void TestTestGetSetMargin()
        {
            BaseTestGetSetMargin(new double[] { 0.7, 0.7, 0.75, 0.75, 0.3, 0.3 });
        }
        [Test]
        public void TestExistingHeaderFooter()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("45540_classic_Header.xlsx");
            XSSFOddHeader hdr;
            XSSFOddFooter ftr;

            // Sheet 1 has a header with center and right text
            XSSFSheet s1 = (XSSFSheet)workbook.GetSheetAt(0);
            Assert.IsNotNull(s1.Header);
            Assert.IsNotNull(s1.Footer);
            hdr = (XSSFOddHeader)s1.Header;
            ftr = (XSSFOddFooter)s1.Footer;

            Assert.AreEqual("&Ctestdoc&Rtest phrase", hdr.Text);
            Assert.AreEqual(null, ftr.Text);

            Assert.AreEqual("", hdr.Left);
            Assert.AreEqual("testdoc", hdr.Center);
            Assert.AreEqual("test phrase", hdr.Right);

            Assert.AreEqual("", ftr.Left);
            Assert.AreEqual("", ftr.Center);
            Assert.AreEqual("", ftr.Right);

            // Sheet 2 has a footer, but it's empty
            XSSFSheet s2 = (XSSFSheet)workbook.GetSheetAt(1);
            Assert.IsNotNull(s2.Header);
            Assert.IsNotNull(s2.Footer);
            hdr = (XSSFOddHeader)s2.Header;
            ftr = (XSSFOddFooter)s2.Footer;

            Assert.AreEqual(null, hdr.Text);
            Assert.AreEqual("&L&F", ftr.Text);

            Assert.AreEqual("", hdr.Left);
            Assert.AreEqual("", hdr.Center);
            Assert.AreEqual("", hdr.Right);

            Assert.AreEqual("&F", ftr.Left);
            Assert.AreEqual("", ftr.Center);
            Assert.AreEqual("", ftr.Right);

            // Save and reload
            IWorkbook wb = XSSFTestDataSamples.WriteOutAndReadBack(workbook);

            hdr = (XSSFOddHeader)wb.GetSheetAt(0).Header;
            ftr = (XSSFOddFooter)wb.GetSheetAt(0).Footer;

            Assert.AreEqual("", hdr.Left);
            Assert.AreEqual("testdoc", hdr.Center);
            Assert.AreEqual("test phrase", hdr.Right);

            Assert.AreEqual("", ftr.Left);
            Assert.AreEqual("", ftr.Center);
            Assert.AreEqual("", ftr.Right);
        }
        [Test]
        public void TestGetAllHeadersFooters()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet 1");
            Assert.IsNotNull(sheet.OddFooter);
            Assert.IsNotNull(sheet.EvenFooter);
            Assert.IsNotNull(sheet.FirstFooter);
            Assert.IsNotNull(sheet.OddHeader);
            Assert.IsNotNull(sheet.EvenHeader);
            Assert.IsNotNull(sheet.FirstHeader);

            Assert.AreEqual("", sheet.OddFooter.Left);
            sheet.OddFooter.Left = ("odd footer left");
            Assert.AreEqual("odd footer left", sheet.OddFooter.Left);

            Assert.AreEqual("", sheet.EvenFooter.Left);
            sheet.EvenFooter.Left = ("even footer left");
            Assert.AreEqual("even footer left", sheet.EvenFooter.Left);

            Assert.AreEqual("", sheet.FirstFooter.Left);
            sheet.FirstFooter.Left = ("first footer left");
            Assert.AreEqual("first footer left", sheet.FirstFooter.Left);

            Assert.AreEqual("", sheet.OddHeader.Left);
            sheet.OddHeader.Left = ("odd header left");
            Assert.AreEqual("odd header left", sheet.OddHeader.Left);

            Assert.AreEqual("", sheet.OddHeader.Right);
            sheet.OddHeader.Right = ("odd header right");
            Assert.AreEqual("odd header right", sheet.OddHeader.Right);

            Assert.AreEqual("", sheet.OddHeader.Center);
            sheet.OddHeader.Center = ("odd header center");
            Assert.AreEqual("odd header center", sheet.OddHeader.Center);

            // Defaults are odd
            Assert.AreEqual("odd footer left", sheet.Footer.Left);
            Assert.AreEqual("odd header center", sheet.Header.Center);
        }
        [Test]
        public void TestAutoSizeColumn()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet 1");
            sheet.CreateRow(0).CreateCell(13).SetCellValue("test");

            sheet.AutoSizeColumn(13);

            ColumnHelper columnHelper = sheet.GetColumnHelper();
            CT_Col col = columnHelper.GetColumn(13, false);
            Assert.IsTrue(col.bestFit);
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

            cell.CellComment = (comment);
            Assert.AreEqual("A1", ctComments.commentList.GetCommentArray(0).@ref);
            comment.Author = ("test A1 author");
            Assert.AreEqual("test A1 author", comments.GetAuthor((int)ctComments.commentList.GetCommentArray(0).authorId));
        }
        [Test]
        public void TestGetActiveCell()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            sheet.SetActiveCell("R5");

            Assert.AreEqual("R5", sheet.ActiveCell);

        }
        [Test]
        public void TestCreateFreezePane_XSSF()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CT_Worksheet ctWorksheet = sheet.GetCTWorksheet();

            sheet.CreateFreezePane(2, 4);
            Assert.AreEqual(2.0, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.xSplit);
            Assert.AreEqual(ST_Pane.bottomRight, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.activePane);
            sheet.CreateFreezePane(3, 6, 10, 10);
            Assert.AreEqual(3.0, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.xSplit);
            //Assert.AreEqual(10, sheet.TopRow);
            //Assert.AreEqual(10, sheet.LeftCol);
            sheet.CreateSplitPane(4, 8, 12, 12, PanePosition.LowerRight);
            Assert.AreEqual(8.0, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.ySplit);
            Assert.AreEqual(ST_Pane.bottomRight, ctWorksheet.sheetViews.GetSheetViewArray(0).pane.activePane);
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
            Assert.AreEqual("C3:D4", ctWorksheet.mergeCells.GetMergeCellArray(1).@ref);
            Assert.AreEqual(3, sheet.NumMergedRegions);
            sheet.RemoveMergedRegion(1);
            Assert.AreEqual("E5:F6", ctWorksheet.mergeCells.GetMergeCellArray(1).@ref);
            Assert.AreEqual(2, sheet.NumMergedRegions);
            sheet.RemoveMergedRegion(1);
            sheet.RemoveMergedRegion(0);
            Assert.AreEqual(0, sheet.NumMergedRegions);
            Assert.IsNull(sheet.GetCTWorksheet().mergeCells, " CTMergeCells should be deleted After removing the last merged " +
                    "region on the sheet.");
        }
        [Test]
        public void TestSetDefaultColumnStyle()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CT_Worksheet ctWorksheet = sheet.GetCTWorksheet();
            StylesTable stylesTable = workbook.GetStylesSource();
            XSSFFont font = new XSSFFont();
            font.FontName = ("Cambria");
            stylesTable.PutFont(font);
            CT_Xf cellStyleXf = new CT_Xf();
            cellStyleXf.fontId = (1);
            cellStyleXf.fillId = 0;
            cellStyleXf.borderId = 0;
            cellStyleXf.numFmtId = 0;
            stylesTable.PutCellStyleXf(cellStyleXf);
            CT_Xf cellXf = new CT_Xf();
            cellXf.xfId = (1);
            stylesTable.PutCellXf(cellXf);
            XSSFCellStyle cellStyle = new XSSFCellStyle(1, 1, stylesTable, null);
            Assert.AreEqual(1, cellStyle.FontIndex);

            sheet.SetDefaultColumnStyle(3, cellStyle);
            Assert.AreEqual(1u, ctWorksheet.GetColsArray(0).GetColArray(0).style);
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
            Assert.AreEqual(2, cols.sizeOfColArray());
            List<CT_Col> colArray = cols.GetColList();
            Assert.IsNotNull(colArray);
            Assert.AreEqual((uint)(2 + 1), colArray[0].min); // 1 based
            Assert.AreEqual((uint)(7 + 1), colArray[0].max); // 1 based
            Assert.AreEqual(1, colArray[0].outlineLevel);
            Assert.AreEqual(0, sheet.GetColumnOutlineLevel(0));

            //two level
            sheet.GroupColumn(1, 2);
            cols = sheet.GetCTWorksheet().GetColsArray(0);
            Assert.AreEqual(4, cols.sizeOfColArray());
            colArray = cols.GetColList();
            Assert.AreEqual(2, colArray[1].outlineLevel);

            //three level
            sheet.GroupColumn(6, 8);
            sheet.GroupColumn(2, 3);
            cols = sheet.GetCTWorksheet().GetColsArray(0);
            Assert.AreEqual(7, cols.sizeOfColArray());
            colArray = cols.GetColList();
            Assert.AreEqual(3, colArray[1].outlineLevel);
            Assert.AreEqual(3, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelCol);

            sheet.UngroupColumn(8, 10);
            colArray = cols.GetColList();
            //Assert.AreEqual(3, colArray[1].outlineLevel);

            sheet.UngroupColumn(4, 6);
            sheet.UngroupColumn(2, 2);
            colArray = cols.GetColList();
            Assert.AreEqual(4, colArray.Count);
            Assert.AreEqual(2, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelCol);
        }

        [Test]
        public void TestGroupUngroupRow()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();

            //one level
            sheet.GroupRow(9, 10);
            Assert.AreEqual(2, sheet.PhysicalNumberOfRows);
            CT_Row ctrow = ((XSSFRow)sheet.GetRow(9)).GetCTRow();

            Assert.IsNotNull(ctrow);
            Assert.AreEqual(10u, ctrow.r);
            Assert.AreEqual(1, ctrow.outlineLevel);
            Assert.AreEqual(1, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelRow);

            //two level
            sheet.GroupRow(10, 13);
            Assert.AreEqual(5, sheet.PhysicalNumberOfRows);
            ctrow = ((XSSFRow)sheet.GetRow(10)).GetCTRow();
            Assert.IsNotNull(ctrow);
            Assert.AreEqual(11u, ctrow.r);
            Assert.AreEqual(2, ctrow.outlineLevel);
            Assert.AreEqual(2, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelRow);


            sheet.UngroupRow(8, 10);
            Assert.AreEqual(4, sheet.PhysicalNumberOfRows);
            Assert.AreEqual(1, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelRow);

            sheet.UngroupRow(10, 10);
            Assert.AreEqual(3, sheet.PhysicalNumberOfRows);

            Assert.AreEqual(1, sheet.GetCTWorksheet().sheetFormatPr.outlineLevelRow);
        }
        [Test]
        public void TestSetZoom()
        {
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)workBook.CreateSheet("new sheet");
            sheet1.SetZoom(3, 4);   // 75 percent magnification
            long zoom = sheet1.GetCTWorksheet().sheetViews.GetSheetViewArray(0).zoomScale;
            Assert.AreEqual(zoom, 75);

            sheet1.SetZoom(200);
            zoom = sheet1.GetCTWorksheet().sheetViews.GetSheetViewArray(0).zoomScale;
            Assert.AreEqual(zoom, 200);

            try
            {
                sheet1.SetZoom(500);
                Assert.Fail("Expecting exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Valid scale values range from 10 to 400", e.Message);
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
            IWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)wb.CreateSheet();

            CT_Cols cols = sheet1.GetCTWorksheet().GetColsArray(0);
            Assert.AreEqual(0, cols.sizeOfColArray());

            sheet1.GroupColumn((short)4, (short)7);
            sheet1.GroupColumn((short)9, (short)12);

            Assert.AreEqual(2, cols.sizeOfColArray());

            Assert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(0).IsSetCollapsed());
            Assert.AreEqual(5, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(8, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(1).IsSetCollapsed());
            Assert.AreEqual(10, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(13, cols.GetColArray(1).max); // 1 based

            sheet1.GroupColumn((short)10, (short)11);
            Assert.AreEqual(4, cols.sizeOfColArray());

            Assert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(0).IsSetCollapsed());
            Assert.AreEqual(5, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(8, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(1).IsSetCollapsed());
            Assert.AreEqual(10, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(10, cols.GetColArray(1).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(2).IsSetCollapsed());
            Assert.AreEqual(11, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(12, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(3).IsSetCollapsed());
            Assert.AreEqual(13, cols.GetColArray(3).min); // 1 based
            Assert.AreEqual(13, cols.GetColArray(3).max); // 1 based

            // collapse columns - 1
            sheet1.SetColumnGroupCollapsed((short)5, true);
            Assert.AreEqual(5, cols.sizeOfColArray());

            Assert.AreEqual(true, cols.GetColArray(0).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(0).IsSetCollapsed());
            Assert.AreEqual(5, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(8, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(1).IsSetCollapsed());
            Assert.AreEqual(9, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(9, cols.GetColArray(1).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(2).IsSetCollapsed());
            Assert.AreEqual(10, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(10, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(3).IsSetCollapsed());
            Assert.AreEqual(11, cols.GetColArray(3).min); // 1 based
            Assert.AreEqual(12, cols.GetColArray(3).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            Assert.AreEqual(13, cols.GetColArray(4).min); // 1 based
            Assert.AreEqual(13, cols.GetColArray(4).max); // 1 based


            // expand columns - 1
            sheet1.SetColumnGroupCollapsed((short)5, false);

            Assert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(0).IsSetCollapsed());
            Assert.AreEqual(5, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(8, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            Assert.AreEqual(false, cols.GetColArray(1).IsSetCollapsed());
            Assert.AreEqual(9, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(9, cols.GetColArray(1).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(2).IsSetCollapsed());
            Assert.AreEqual(10, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(10, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(3).IsSetCollapsed());
            Assert.AreEqual(11, cols.GetColArray(3).min); // 1 based
            Assert.AreEqual(12, cols.GetColArray(3).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            Assert.AreEqual(13, cols.GetColArray(4).min); // 1 based
            Assert.AreEqual(13, cols.GetColArray(4).max); // 1 based


            //collapse - 2
            sheet1.SetColumnGroupCollapsed((short)9, true);
            Assert.AreEqual(6, cols.sizeOfColArray());
            Assert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(0).IsSetCollapsed());
            Assert.AreEqual(5, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(8, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(1).IsSetCollapsed());
            Assert.AreEqual(9, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(9, cols.GetColArray(1).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(2).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(2).IsSetCollapsed());
            Assert.AreEqual(10, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(10, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(3).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(3).IsSetCollapsed());
            Assert.AreEqual(11, cols.GetColArray(3).min); // 1 based
            Assert.AreEqual(12, cols.GetColArray(3).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(4).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            Assert.AreEqual(13, cols.GetColArray(4).min); // 1 based
            Assert.AreEqual(13, cols.GetColArray(4).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(5).IsSetCollapsed());
            Assert.AreEqual(14, cols.GetColArray(5).min); // 1 based
            Assert.AreEqual(14, cols.GetColArray(5).max); // 1 based


            //expand - 2
            sheet1.SetColumnGroupCollapsed((short)9, false);
            Assert.AreEqual(6, cols.sizeOfColArray());
            Assert.AreEqual(14, cols.GetColArray(5).min);

            //outline level 2: the line under ==> collapsed==True
            Assert.AreEqual(2, cols.GetColArray(3).outlineLevel);
            Assert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());

            Assert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(0).IsSetCollapsed());
            Assert.AreEqual(5, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(8, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(1).IsSetCollapsed());
            Assert.AreEqual(9, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(9, cols.GetColArray(1).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(2).IsSetCollapsed());
            Assert.AreEqual(10, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(10, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(3).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(3).IsSetCollapsed());
            Assert.AreEqual(11, cols.GetColArray(3).min); // 1 based
            Assert.AreEqual(12, cols.GetColArray(3).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            Assert.AreEqual(13, cols.GetColArray(4).min); // 1 based
            Assert.AreEqual(13, cols.GetColArray(4).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            Assert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            Assert.AreEqual(14, cols.GetColArray(5).min); // 1 based
            Assert.AreEqual(14, cols.GetColArray(5).max); // 1 based

            //DOCUMENTARE MEGLIO IL DISCORSO DEL LIVELLO
            //collapse - 3
            sheet1.SetColumnGroupCollapsed((short)10, true);
            Assert.AreEqual(6, cols.sizeOfColArray());
            Assert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(0).IsSetCollapsed());
            Assert.AreEqual(5, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(8, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(1).IsSetCollapsed());
            Assert.AreEqual(9, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(9, cols.GetColArray(1).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(2).IsSetCollapsed());
            Assert.AreEqual(10, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(10, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(3).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(3).IsSetCollapsed());
            Assert.AreEqual(11, cols.GetColArray(3).min); // 1 based
            Assert.AreEqual(12, cols.GetColArray(3).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(4).IsSetCollapsed());
            Assert.AreEqual(13, cols.GetColArray(4).min); // 1 based
            Assert.AreEqual(13, cols.GetColArray(4).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            Assert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            Assert.AreEqual(14, cols.GetColArray(5).min); // 1 based
            Assert.AreEqual(14, cols.GetColArray(5).max); // 1 based


            //expand - 3
            sheet1.SetColumnGroupCollapsed((short)10, false);
            Assert.AreEqual(6, cols.sizeOfColArray());
            Assert.AreEqual(false, cols.GetColArray(0).hidden);
            Assert.AreEqual(false, cols.GetColArray(5).hidden);
            Assert.AreEqual(false, cols.GetColArray(4).IsSetCollapsed());

            //      write out and give back
            // Save and re-load
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet1 = (XSSFSheet)wb.GetSheetAt(0);
            Assert.AreEqual(6, cols.sizeOfColArray());

            Assert.AreEqual(false, cols.GetColArray(0).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(0).IsSetCollapsed());
            Assert.AreEqual(5, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(8, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(1).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(1).IsSetCollapsed());
            Assert.AreEqual(9, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(9, cols.GetColArray(1).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(2).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(2).IsSetCollapsed());
            Assert.AreEqual(10, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(10, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(3).IsSetHidden());
            Assert.AreEqual(true, cols.GetColArray(3).IsSetCollapsed());
            Assert.AreEqual(11, cols.GetColArray(3).min); // 1 based
            Assert.AreEqual(12, cols.GetColArray(3).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(4).IsSetHidden());
            Assert.AreEqual(false, cols.GetColArray(4).IsSetCollapsed());
            Assert.AreEqual(13, cols.GetColArray(4).min); // 1 based
            Assert.AreEqual(13, cols.GetColArray(4).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(5).IsSetHidden());
            Assert.AreEqual(false, cols.GetColArray(5).IsSetCollapsed());
            Assert.AreEqual(14, cols.GetColArray(5).min); // 1 based
            Assert.AreEqual(14, cols.GetColArray(5).max); // 1 based

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
            IWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)wb.CreateSheet();

            sheet1.GroupRow(5, 14);
            sheet1.GroupRow(7, 14);
            sheet1.GroupRow(16, 19);

            Assert.AreEqual(14, sheet1.PhysicalNumberOfRows);
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetHidden());

            //collapsed
            sheet1.SetRowGroupCollapsed(7, true);

            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetHidden());

            //expanded
            sheet1.SetRowGroupCollapsed(7, false);

            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetHidden());


            // Save and re-load
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet1 = (XSSFSheet)wb.GetSheetAt(0);

            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(6)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(7)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(9)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(true, ((XSSFRow)sheet1.GetRow(14)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(16)).GetCTRow().IsSetHidden());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetCollapsed());
            Assert.AreEqual(false, ((XSSFRow)sheet1.GetRow(18)).GetCTRow().IsSetHidden());
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
            Assert.AreEqual(22 * 256, sheet.GetColumnWidth(1));

            // Now check the low level stuff, and check that's all
            //  been Set correctly
            XSSFSheet xs = sheet;
            CT_Worksheet cts = xs.GetCTWorksheet();

            List<CT_Cols> cols_s = cts.GetColsList();
            Assert.AreEqual(1, cols_s.Count);
            CT_Cols cols = cols_s[0];
            Assert.AreEqual(1, cols.sizeOfColArray());
            CT_Col col = cols.GetColArray(0);

            // XML is 1 based, POI is 0 based
            Assert.AreEqual(2u, col.min);
            Assert.AreEqual(2u, col.max);
            Assert.AreEqual(22.0, col.width, 0.0);
            Assert.IsTrue(col.customWidth);

            // Now Set another
            sheet.SetColumnWidth(3, 33 * 256);

            cols_s = cts.GetColsList();
            Assert.AreEqual(1, cols_s.Count);
            cols = cols_s[0];
            Assert.AreEqual(2, cols.sizeOfColArray());

            col = cols.GetColArray(0);
            Assert.AreEqual(2u, col.min); // POI 1
            Assert.AreEqual(2u, col.max);
            Assert.AreEqual(22.0, col.width, 0.0);
            Assert.IsTrue(col.customWidth);

            col = cols.GetColArray(1);
            Assert.AreEqual(4u, col.min); // POI 3
            Assert.AreEqual(4u, col.max);
            Assert.AreEqual(33.0, col.width, 0.0);
            Assert.IsTrue(col.customWidth);
        }

        /**
         * Setting width of a column included in a column span
         */
        [Test]
        public void Test47862()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("47862.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            //<cols>
            //  <col min="1" max="5" width="15.77734375" customWidth="1"/>
            //</cols>

            //a span of columns [1,5]
            Assert.AreEqual(1, cols.sizeOfColArray());
            CT_Col col = cols.GetColArray(0);
            Assert.AreEqual((uint)1, col.min);
            Assert.AreEqual((uint)5, col.max);
            double swidth = 15.77734375; //width of columns in the span
            Assert.AreEqual(swidth, col.width, 0.0);

            for (int i = 0; i < 5; i++)
            {
                Assert.AreEqual((int)(swidth * 256), sheet.GetColumnWidth(i));
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
            Assert.AreEqual(5, cols.sizeOfColArray());
            for (int i = 0; i < 5; i++)
            {
                Assert.AreEqual(cw[i] * 256, sheet.GetColumnWidth(i));
                Assert.AreEqual(cw[i], cols.GetColArray(i).width, 0.0);
            }

            //serialize and check again
            wb = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = (XSSFSheet)wb.GetSheetAt(0);
            cols = sheet.GetCTWorksheet().GetColsArray(0);
            Assert.AreEqual(5, cols.sizeOfColArray());
            for (int i = 0; i < 5; i++)
            {
                Assert.AreEqual(cw[i] * 256, sheet.GetColumnWidth(i));
                Assert.AreEqual(cw[i], cols.GetColArray(i).width, 0.0);
            }
        }

        /**
         * Hiding a column included in a column span
         */
        [Test]
        public void Test47804()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("47804.xlsx");
            XSSFSheet sheet = (XSSFSheet)wb.GetSheetAt(0);
            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            Assert.AreEqual(2, cols.sizeOfColArray());
            CT_Col col;
            //<cols>
            //  <col min="2" max="4" width="12" customWidth="1"/>
            //  <col min="7" max="7" width="10.85546875" customWidth="1"/>
            //</cols>

            //a span of columns [2,4]
            col = cols.GetColArray(0);
            Assert.AreEqual((uint)2, col.min);
            Assert.AreEqual((uint)4, col.max);
            //individual column
            col = cols.GetColArray(1);
            Assert.AreEqual((uint)7, col.min);
            Assert.AreEqual((uint)7, col.max);

            sheet.SetColumnHidden(2, true); // Column C
            sheet.SetColumnHidden(6, true); // Column G

            Assert.IsTrue(sheet.IsColumnHidden(2));
            Assert.IsTrue(sheet.IsColumnHidden(6));

            //other columns but C and G are not hidden
            Assert.IsFalse(sheet.IsColumnHidden(1));
            Assert.IsFalse(sheet.IsColumnHidden(3));
            Assert.IsFalse(sheet.IsColumnHidden(4));
            Assert.IsFalse(sheet.IsColumnHidden(5));

            //the check below failed prior to fix of Bug #47804
            ColumnHelper.SortColumns(cols);
            //the span is now splitted into three parts
            //<cols>
            //  <col min="2" max="2" customWidth="true" width="12.0" />
            //  <col min="3" max="3" customWidth="true" width="12.0" hidden="true"/>
            //  <col min="4" max="4" customWidth="true" width="12.0"/>
            //  <col min="7" max="7" customWidth="true" width="10.85546875" hidden="true"/>
            //</cols>

            Assert.AreEqual(4, cols.sizeOfColArray());
            col = cols.GetColArray(0);
            Assert.AreEqual((uint)2, col.min);
            Assert.AreEqual((uint)2, col.max);
            col = cols.GetColArray(1);
            Assert.AreEqual((uint)3, col.min);
            Assert.AreEqual((uint)3, col.max);
            col = cols.GetColArray(2);
            Assert.AreEqual((uint)4, col.min);
            Assert.AreEqual((uint)4, col.max);
            col = cols.GetColArray(3);
            Assert.AreEqual((uint)7, col.min);
            Assert.AreEqual((uint)7, col.max);

            //serialize and check again
            wb = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = (XSSFSheet)wb.GetSheetAt(0);
            Assert.IsTrue(sheet.IsColumnHidden(2));
            Assert.IsTrue(sheet.IsColumnHidden(6));
            Assert.IsFalse(sheet.IsColumnHidden(1));
            Assert.IsFalse(sheet.IsColumnHidden(3));
            Assert.IsFalse(sheet.IsColumnHidden(4));
            Assert.IsFalse(sheet.IsColumnHidden(5));
        }
        [Test]
        public void TestCommentsTable()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)workbook.CreateSheet();
            CommentsTable comment1 = sheet1.GetCommentsTable(false);
            Assert.IsNull(comment1);

            comment1 = sheet1.GetCommentsTable(true);
            Assert.IsNotNull(comment1);
            Assert.AreEqual("/xl/comments1.xml", comment1.GetPackageRelationship().TargetUri.ToString());

            Assert.AreSame(comment1, sheet1.GetCommentsTable(true));

            //second sheet
            XSSFSheet sheet2 = (XSSFSheet)workbook.CreateSheet();
            CommentsTable comment2 = sheet2.GetCommentsTable(false);
            Assert.IsNull(comment2);

            comment2 = sheet2.GetCommentsTable(true);
            Assert.IsNotNull(comment2);

            Assert.AreSame(comment2, sheet2.GetCommentsTable(true));
            Assert.AreEqual("/xl/comments2.xml", comment2.GetPackageRelationship().TargetUri.ToString());

            //comment1 and  comment2 are different objects
            Assert.AreNotSame(comment1, comment2);

            //now Test against a workbook Containing cell comments
            workbook = XSSFTestDataSamples.OpenSampleWorkbook("WithMoreVariousData.xlsx");
            sheet1 = (XSSFSheet)workbook.GetSheetAt(0);
            comment1 = sheet1.GetCommentsTable(true);
            Assert.IsNotNull(comment1);
            Assert.AreEqual("/xl/comments1.xml", comment1.GetPackageRelationship().TargetUri.ToString());
            Assert.AreSame(comment1, sheet1.GetCommentsTable(true));
        }

        /**
         * Rows and cells can be Created in random order,
         * but CTRows are kept in ascending order
         */
        [Test]
        public new void TestCreateRow()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CT_Worksheet wsh = sheet.GetCTWorksheet();
            CT_SheetData sheetData = wsh.sheetData;
            Assert.AreEqual(0, sheetData.SizeOfRowArray());

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
            Assert.AreEqual(3, xrow.Count);

            //rows are sorted: {0, 1, 2}
            Assert.AreEqual(4, xrow[0].SizeOfCArray());
            Assert.AreEqual(1u, xrow[0].r);
            Assert.IsTrue(xrow[0].Equals(row3.GetCTRow()));

            Assert.AreEqual(3, xrow[1].SizeOfCArray());
            Assert.AreEqual(2u, xrow[1].r);
            Assert.IsTrue(xrow[1].Equals(row2.GetCTRow()));

            Assert.AreEqual(2, xrow[2].SizeOfCArray());
            Assert.AreEqual(3u, xrow[2].r);
            Assert.IsTrue(xrow[2].Equals(row1.GetCTRow()));

            List<CT_Cell> xcell = xrow[0].c;
            Assert.AreEqual("D1", xcell[0].r);
            Assert.AreEqual("A1", xcell[1].r);
            Assert.AreEqual("C1", xcell[2].r);
            Assert.AreEqual("F1", xcell[3].r);

            //re-creating a row does NOT add extra data to the parent
            row2 = (XSSFRow)sheet.CreateRow(1);
            Assert.AreEqual(3, sheetData.SizeOfRowArray());
            //existing cells are invalidated
            Assert.AreEqual(0, sheetData.GetRowArray(1).SizeOfCArray());
            Assert.AreEqual(0, row2.PhysicalNumberOfCells);

            workbook = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            sheet = (XSSFSheet)workbook.GetSheetAt(0);
            wsh = sheet.GetCTWorksheet();
            xrow = sheetData.row;
            Assert.AreEqual(3, xrow.Count);

            //rows are sorted: {0, 1, 2}
            Assert.AreEqual(4, xrow[0].SizeOfCArray());
            Assert.AreEqual(1u, xrow[0].r);
            //cells are now sorted
            xcell = xrow[0].c;
            Assert.AreEqual("A1", xcell[0].r);
            Assert.AreEqual("C1", xcell[1].r);
            Assert.AreEqual("D1", xcell[2].r);
            Assert.AreEqual("F1", xcell[3].r);


            Assert.AreEqual(0, xrow[1].SizeOfCArray());
            Assert.AreEqual(2u, xrow[1].r);

            Assert.AreEqual(2, xrow[2].SizeOfCArray());
            Assert.AreEqual(3u, xrow[2].r);

        }
        [Test]
        public void TestSetAutoFilter()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("new sheet");
            sheet.SetAutoFilter(CellRangeAddress.ValueOf("A1:D100"));

            Assert.AreEqual("A1:D100", sheet.GetCTWorksheet().autoFilter.@ref);

            // auto-filter must be registered in workboook.xml, see Bugzilla 50315
            XSSFName nm = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, 0);
            Assert.IsNotNull(nm);

            Assert.AreEqual(0u, nm.GetCTName().localSheetId);
            Assert.AreEqual(true, nm.GetCTName().hidden);
            Assert.AreEqual("_xlnm._FilterDatabase", nm.GetCTName().name);
            Assert.AreEqual("'new sheet'!$A$1:$D$100", nm.GetCTName().Value);
        }
        //[Test]
        //public void TestProtectSheet_lowlevel()
        //{

        //    XSSFWorkbook wb = new XSSFWorkbook();
        //    XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
        //    CT_SheetProtection pr = sheet.GetCTWorksheet().sheetProtection;
        //    Assert.IsNull(pr, "CTSheetProtection should be null by default");
        //    String password = "Test";
        //    sheet.ProtectSheet(password);
        //    pr = sheet.GetCTWorksheet().sheetProtection;
        //    Assert.IsNotNull(pr, "CTSheetProtection should be not null");
        //    Assert.IsTrue(pr.sheet, "sheet protection should be on");
        //    Assert.IsTrue(pr.objects, "object protection should be on");
        //    Assert.IsTrue(pr.scenarios, "scenario protection should be on");
        //    String hash = HexDump.ShortToHex(PasswordRecord.HashPassword(password)).ToString().Substring(2);
        //    Assert.AreEqual(hash, pr.xgetPassword().StringValue, "well known value for top secret hash should be " + hash);

        //    sheet.ProtectSheet(null);
        //    Assert.IsNull(sheet.GetCTWorksheet().sheetProtection, "protectSheet(null) should unset CTSheetProtection");
        //}

        [Test]
        public void Test49966()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("49966.xlsx");
            CalculationChain calcChain = wb.GetCalculationChain();
            Assert.IsNotNull(wb.GetCalculationChain());
            Assert.AreEqual(3, calcChain.GetCTCalcChain().SizeOfCArray());

            ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(0);

            sheet.RemoveRow(row);
            Assert.AreEqual(0, calcChain.GetCTCalcChain().SizeOfCArray(), "XSSFSheet#RemoveRow did not clear calcChain entries");

            //calcChain should be gone 
            wb = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNull(wb.GetCalculationChain());
        }

        /**
         * See bug #50829
         */
        [Test]
        public void TestTables()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithTable.xlsx");
            Assert.AreEqual(3, wb.NumberOfSheets);

            // Check the table sheet
            XSSFSheet s1 = (XSSFSheet)wb.GetSheetAt(0);
            Assert.AreEqual("a", s1.GetRow(0).GetCell(0).RichStringCellValue.ToString());
            Assert.AreEqual(1.0, s1.GetRow(1).GetCell(0).NumericCellValue);

            List<XSSFTable> tables = s1.GetTables();
            Assert.IsNotNull(tables);
            Assert.AreEqual(1, tables.Count);

            XSSFTable table = tables[0];
            Assert.AreEqual("Tabella1", table.Name);
            Assert.AreEqual("Tabella1", table.DisplayName);

            // And the others
            XSSFSheet s2 = (XSSFSheet)wb.GetSheetAt(1);
            Assert.AreEqual(0, s2.GetTables().Count);
            XSSFSheet s3 = (XSSFSheet)wb.GetSheetAt(2);
            Assert.AreEqual(0, s3.GetTables().Count);
        }

        /**
     * Test to trigger OOXML-LITE generating to include org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetCalcPr
     */
        [Test]
        public void TestSetForceFormulaRecalculation()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet 1");

            Assert.IsFalse(sheet.ForceFormulaRecalculation);

            // Set
            sheet.ForceFormulaRecalculation = (true);
            Assert.AreEqual(true, sheet.ForceFormulaRecalculation);

            // calcMode="manual" is unset when forceFormulaRecalculation=true
            CT_CalcPr calcPr = workbook.GetCTWorkbook().AddNewCalcPr();
            calcPr.calcMode = (ST_CalcMode.manual);
            sheet.ForceFormulaRecalculation = (true);
            Assert.AreEqual(ST_CalcMode.auto, calcPr.calcMode);

            // Check
            sheet.ForceFormulaRecalculation = (false);
            Assert.AreEqual(false, sheet.ForceFormulaRecalculation);


            // Save, re-load, and re-check
            workbook = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            sheet = (XSSFSheet)workbook.GetSheet("Sheet 1");
            Assert.AreEqual(false, sheet.ForceFormulaRecalculation);
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
            IWorkbook wb;
            if (isXSSF)
            {
                wb = XSSFTestDataSamples.OpenSampleWorkbook(file);
            }
            else
            {
                wb = HSSFTestDataSamples.OpenSampleWorkbook(file);
            }
            for (int si = 0; si < wb.NumberOfSheets; si++)
            {
                ISheet sh = wb.GetSheetAt(si);
                Assert.IsNotNull(sh.SheetName);
                Assert.AreEqual(topRows[si], sh.TopRow, "Did not match for sheet " + si);
            }

            // for XSSF also test with SXSSF
            //if (isXSSF)
            //{
            //    Workbook swb = new SXSSFWorkbook((XSSFWorkbook)wb);
            //    try
            //    {
            //        for (int si = 0; si < swb.getNumberOfSheets(); si++)
            //        {
            //            Sheet sh = swb.getSheetAt(si);
            //            assertNotNull(sh.getSheetName());
            //            assertEquals("Did not match for sheet " + si, topRows[si], sh.getTopRow());
            //        }
            //    }
            //    finally
            //    {
            //        swb.close();
            //    }
            //}
        }

        private void runGetLeftCol(String file, bool isXSSF, params int[] topRows)
        {
            IWorkbook wb;
            if (isXSSF)
            {
                wb = XSSFTestDataSamples.OpenSampleWorkbook(file);
            }
            else
            {
                wb = HSSFTestDataSamples.OpenSampleWorkbook(file);
            }
            for (int si = 0; si < wb.NumberOfSheets; si++)
            {
                ISheet sh = wb.GetSheetAt(si);
                Assert.IsNotNull(sh.SheetName);
                Assert.AreEqual(topRows[si], sh.LeftCol, "Did not match for sheet " + si);
            }

            // for XSSF also test with SXSSF
            //if (isXSSF)
            //{
            //    IWorkbook swb = new SXSSFWorkbook((XSSFWorkbook)wb);
            //    for (int si = 0; si < swb.NumberOfSheets; si++)
            //    {
            //        ISheet sh = swb.GetSheetAt(si);
            //        Assert.IsNotNull(sh.SheetName);
            //        Assert.AreEqual("Did not match for sheet " + si, topRows[si], sh.GetLeftCol());
            //    }
            //    swb.Close();
            //}
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
            Assert.AreEqual(8, tables.Count, "Sheet should contain 8 tables");
            Assert.IsNotNull(sheet.GetCommentsTable(false), "Sheet should contain a comments table");
        }

        [Test]
        public void Bug55723b()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            // stored with a special name
            Assert.IsNull(wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, 0));

            CellRangeAddress range = CellRangeAddress.ValueOf("A:B");
            IAutoFilter filter = sheet.SetAutoFilter(range);
            Assert.IsNotNull(filter);

            // stored with a special name
            XSSFName name = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, 0);
            Assert.IsNotNull(name);
            Assert.AreEqual("Sheet0!$A:$B", name.RefersToFormula);

            range = CellRangeAddress.ValueOf("B:C");
            filter = sheet.SetAutoFilter(range);
            Assert.IsNotNull(filter);

            // stored with a special name
            name = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, 0);
            Assert.IsNotNull(name);
            Assert.AreEqual("Sheet0!$B:$C", name.RefersToFormula);
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

            IRow row1 = sheet.CreateRow((short)0);
            ICell cell = row1.CreateCell((short)0);
            cell.SetCellValue("Names");
            ICell cell2 = row1.CreateCell((short)1);
            cell2.SetCellValue("#");

            IRow row2 = sheet.CreateRow((short)1);
            ICell cell3 = row2.CreateCell((short)0);
            cell3.SetCellValue("Jane");
            ICell cell4 = row2.CreateCell((short)1);
            cell4.SetCellValue(3);

            IRow row3 = sheet.CreateRow((short)2);
            ICell cell5 = row3.CreateCell((short)0);
            cell5.SetCellValue("John");
            ICell cell6 = row3.CreateCell((short)1);
            cell6.SetCellValue(3);

            return wb;
        }

        [Test]
        public void TestCreateTwoPivotTablesInOneSheet()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            Assert.IsNotNull(wb);
            Assert.IsNotNull(sheet);
            XSSFPivotTable pivotTable = sheet.CreatePivotTable(new AreaReference("A1:B2"), new CellReference("H5"));
            Assert.IsNotNull(pivotTable);
            Assert.IsTrue(wb.PivotTables.Count > 0);
            XSSFPivotTable pivotTable2 = sheet.CreatePivotTable(new AreaReference("A1:B2"), new CellReference("L5"), sheet);
            Assert.IsNotNull(pivotTable2);
            Assert.IsTrue(wb.PivotTables.Count > 1);
        }

        [Test]
        public void TestCreateTwoPivotTablesInTwoSheets()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            Assert.IsNotNull(wb);
            Assert.IsNotNull(sheet);
            XSSFPivotTable pivotTable = sheet.CreatePivotTable(new AreaReference("A1:B2"), new CellReference("H5"));
            Assert.IsNotNull(pivotTable);
            Assert.IsTrue(wb.PivotTables.Count > 0);
            Assert.IsNotNull(wb);
            XSSFSheet sheet2 = wb.CreateSheet() as XSSFSheet;
            XSSFPivotTable pivotTable2 = sheet2.CreatePivotTable(new AreaReference("A1:B2"), new CellReference("H5"), sheet);
            Assert.IsNotNull(pivotTable2);
            Assert.IsTrue(wb.PivotTables.Count > 1);
        }

        [Test]
        public void TestCreatePivotTable()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            Assert.IsNotNull(wb);
            Assert.IsNotNull(sheet);
            XSSFPivotTable pivotTable = sheet.CreatePivotTable(new AreaReference("A1:B2"), new CellReference("H5"));
            Assert.IsNotNull(pivotTable);
            Assert.IsTrue(wb.PivotTables.Count > 0);
        }

        [Test]
        public void TestCreatePivotTableInOtherSheetThanDataSheet()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet1 = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sheet2 = wb.CreateSheet() as XSSFSheet;

            XSSFPivotTable pivotTable = sheet2.CreatePivotTable
                    (new AreaReference("A1:B2"), new CellReference("H5"), sheet1);
            Assert.AreEqual(0, pivotTable.GetRowLabelColumns().Count);

            Assert.AreEqual(1, wb.PivotTables.Count);
            Assert.AreEqual(0, sheet1.GetPivotTables().Count);
            Assert.AreEqual(1, sheet2.GetPivotTables().Count);
        }

        [Test]
        public void TestCreatePivotTableInOtherSheetThanDataSheetUsingAreaReference()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sheet2 = wb.CreateSheet() as XSSFSheet;

            XSSFPivotTable pivotTable = sheet2.CreatePivotTable
                    (new AreaReference(sheet.SheetName + "!A$1:B$2"), new CellReference("H5"));
            Assert.AreEqual(0, pivotTable.GetRowLabelColumns().Count);
        }

        [Test]
        public void TestCreatePivotTableWithConflictingDataSheets()
        {
            XSSFWorkbook wb = SetupSheet();
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sheet2 = wb.CreateSheet() as XSSFSheet;

            try
            {
                sheet2.CreatePivotTable(new AreaReference(sheet.SheetName + "!A$1:B$2"), new CellReference("H5"), sheet2);
            }
            catch (ArgumentException)
            {
                return;
            }
            Assert.Fail();
        }

        [Test]
        public void TestReadFails()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

            try
            {
                sheet.OnDocumentRead();
                Assert.Fail("Throws exception because we cannot read here");
            }
            catch (POIXMLException e)
            {
                // expected here
            }
        }

        [Test]
        public void TestCreateComment()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            Assert.IsNotNull(sheet.CreateComment());
        }

    }
}