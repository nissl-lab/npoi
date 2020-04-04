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

namespace TestCases.SS.Util
{
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;
    using TestCases.SS;


    public abstract class BaseTestCellUtil
    {
        protected ITestDataProvider _testDataProvider;
        //public BaseTestCellUtil()
        //    : this(HSSFITestDataProvider.Instance)
        //{

        //}
        protected BaseTestCellUtil(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }

        [Test]
        public void SetCellStyleProperty()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0);

            // Add a border should create a new style
            int styCnt1 = wb.NumCellStyles;
            CellUtil.SetCellStyleProperty(c, CellUtil.BORDER_BOTTOM, BorderStyle.Thin);
            int styCnt2 = wb.NumCellStyles;
            Assert.AreEqual(styCnt1 + 1, styCnt2);

            // Add same border to another cell, should not create another style
            c = r.CreateCell(1);
            CellUtil.SetCellStyleProperty(c, CellUtil.BORDER_BOTTOM, BorderStyle.Thin);
            int styCnt3 = wb.NumCellStyles;
            Assert.AreEqual(styCnt2, styCnt3);

            wb.Close();
        }

        [Test]//(expected=RuntimeException.class)
        public void SetCellStylePropertyWithInvalidValue()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0);
            // An invalid BorderStyle constant
            CellUtil.SetCellStyleProperty(c, CellUtil.BORDER_BOTTOM, 42);

            wb.Close();
        }

        [Test]
        public void SetCellStylePropertyBorderWithShortAndEnum()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0);
            // A valid BorderStyle constant, as a Short
            CellUtil.SetCellStyleProperty(c, CellUtil.BORDER_BOTTOM, (short)BorderStyle.DashDot);
            Assert.AreEqual(BorderStyle.DashDot, c.CellStyle.BorderBottom);

            // A valid BorderStyle constant, as an Enum
            CellUtil.SetCellStyleProperty(c, CellUtil.BORDER_TOP, BorderStyle.MediumDashDot);
            Assert.AreEqual(BorderStyle.MediumDashDot, c.CellStyle.BorderTop);

            wb.Close();
        }


        [Test]
        public void SetCellStyleProperties()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0);

            // Add multiple border properties to cell should create a single new style
            int styCnt1 = wb.NumCellStyles;
            Dictionary<String, Object> props = new Dictionary<String, Object>();
            props.Add(CellUtil.BORDER_TOP, BorderStyle.Thin);
            props.Add(CellUtil.BORDER_BOTTOM, BorderStyle.Thin);
            props.Add(CellUtil.BORDER_LEFT, BorderStyle.Thin);
            props.Add(CellUtil.BORDER_RIGHT, BorderStyle.Thin);
            props.Add(CellUtil.ALIGNMENT, HorizontalAlignment.Center); // try it both with a Short (deprecated)
            props.Add(CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.Center); // and with an enum
            CellUtil.SetCellStyleProperties(c, props);
            int styCnt2 = wb.NumCellStyles;
            Assert.AreEqual(styCnt1 + 1, styCnt2, "Only one additional style should have been created");

            // Add same border another to same cell, should not create another style
            c = r.CreateCell(1);
            CellUtil.SetCellStyleProperties(c, props);
            int styCnt3 = wb.NumCellStyles;
            Assert.AreEqual(styCnt3, styCnt2, "No additional styles should have been created");

            wb.Close();
        }

        [Test]
        public void GetRow()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row1 = sh.CreateRow(0);

            // Get row that already exists
            IRow r1 = CellUtil.GetRow(0, sh);
            Assert.IsNotNull(r1);
            Assert.AreSame(row1, r1, "An existing row should not be reCreated");

            // Get row that does not exist yet
            Assert.IsNotNull(CellUtil.GetRow(1, sh));

            wb.Close();
        }

        [Test]
        public void GetCell()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row = sh.CreateRow(0);
            ICell A1 = row.CreateCell(0);

            // Get cell that already exists
            ICell a1 = CellUtil.GetCell(row, 0);
            Assert.IsNotNull(a1);
            Assert.AreSame(A1, a1, "An existing cell should not be reCreated");

            // Get cell that does not exist yet
            Assert.IsNotNull(CellUtil.GetCell(row, 1));

            wb.Close();
        }

        [Test]
        public void CreateCell()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row = sh.CreateRow(0);

            ICellStyle style = wb.CreateCellStyle();
            style.WrapText = (/*setter*/true);

            // calling CreateCell on a non-existing cell should create a cell and Set the cell value and style.
            ICell F1 = CellUtil.CreateCell(row, 5, "Cell Value", style);

            Assert.AreSame(row.GetCell(5), F1);
            Assert.AreEqual("Cell Value", F1.StringCellValue);
            Assert.AreEqual(style, F1.CellStyle);
            // should be Assert.AreSame, but a new HSSFCellStyle is returned for each GetCellStyle() call.
            // HSSFCellStyle wraps an underlying style record, and the underlying
            // style record is the same between multiple GetCellStyle() calls.

            // calling CreateCell on an existing cell should return the existing cell and modify the cell value and style.
            ICell f1 = CellUtil.CreateCell(row, 5, "Overwritten cell value", null);
            Assert.AreSame(row.GetCell(5), f1);
            Assert.AreSame(F1, f1);
            Assert.AreEqual("Overwritten cell value", f1.StringCellValue);
            Assert.AreEqual("Overwritten cell value", F1.StringCellValue);
            Assert.AreEqual(style, f1.CellStyle, "cell style should be unChanged with CreateCell(..., null)");
            Assert.AreEqual(style, F1.CellStyle, "cell style should be unChanged with CreateCell(..., null)");

            // test CreateCell(row, column, value) (no CellStyle)
            f1 = CellUtil.CreateCell(row, 5, "Overwritten cell with default style");
            Assert.AreSame(F1, f1);

            wb.Close();

        }

        [Test]
        public void SetAlignmentEnum()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row = sh.CreateRow(0);
            ICell A1 = row.CreateCell(0);
            ICell B1 = row.CreateCell(1);
            // Assumptions
            Assert.AreEqual(A1.CellStyle, B1.CellStyle);
            // should be assertSame, but a new HSSFCellStyle is returned for each getCellStyle() call. 
            // HSSFCellStyle wraps an underlying style record, and the underlying
            // style record is the same between multiple getCellStyle() calls.
            Assert.AreEqual(HorizontalAlignment.General, A1.CellStyle.Alignment);
            Assert.AreEqual(HorizontalAlignment.General, B1.CellStyle.Alignment);
            // get/set alignment modifies the cell's style
            CellUtil.SetAlignment(A1, HorizontalAlignment.Right);
            Assert.AreEqual(HorizontalAlignment.Right, A1.CellStyle.Alignment);
            // get/set alignment doesn't affect the style of cells with
            // the same style prior to modifying the style
            Assert.AreNotEqual(A1.CellStyle, B1.CellStyle);
            Assert.AreEqual(HorizontalAlignment.General, B1.CellStyle.Alignment);
            wb.Close();
        }

        [Test]
        public void SetVerticalAlignmentEnum()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row = sh.CreateRow(0);
            ICell A1 = row.CreateCell(0);
            ICell B1 = row.CreateCell(1);
            // Assumptions
            Assert.AreEqual(A1.CellStyle, B1.CellStyle);
            // should be assertSame, but a new HSSFCellStyle is returned for each getCellStyle() call. 
            // HSSFCellStyle wraps an underlying style record, and the underlying
            // style record is the same between multiple getCellStyle() calls.
            Assert.AreEqual(VerticalAlignment.Bottom, A1.CellStyle.VerticalAlignment);
            Assert.AreEqual(VerticalAlignment.Bottom, B1.CellStyle.VerticalAlignment);
            // get/set alignment modifies the cell's style
            CellUtil.SetVerticalAlignment(A1, VerticalAlignment.Top);
            Assert.AreEqual(VerticalAlignment.Top, A1.CellStyle.VerticalAlignment);
            // get/set alignment doesn't affect the style of cells with
            // the same style prior to modifying the style
            Assert.AreNotEqual(A1.CellStyle, B1.CellStyle);
            Assert.AreEqual(VerticalAlignment.Bottom, B1.CellStyle.VerticalAlignment);
            wb.Close();
        }


        [Test]
        public void SetFont()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row = sh.CreateRow(0);
            ICell A1 = row.CreateCell(0);
            ICell B1 = row.CreateCell(1);
            short defaultFontIndex = 0;
            IFont font = wb.CreateFont();
            font.IsItalic = true;
            short customFontIndex = font.Index;

            // Assumptions
            Assert.AreNotEqual(defaultFontIndex, customFontIndex);
            Assert.AreEqual(A1.CellStyle, B1.CellStyle);
            // should be Assert.AreSame, but a new HSSFCellStyle is returned for each GetCellStyle() call. 
            // HSSFCellStyle wraps an underlying style record, and the underlying
            // style record is the same between multiple GetCellStyle() calls.
            Assert.AreEqual(defaultFontIndex, A1.CellStyle.FontIndex);
            Assert.AreEqual(defaultFontIndex, B1.CellStyle.FontIndex);

            // Get/set alignment modifies the cell's style
            CellUtil.SetFont(A1, font);
            Assert.AreEqual(customFontIndex, A1.CellStyle.FontIndex);

            // Get/set alignment doesn't affect the style of cells with
            // the same style prior to modifying the style
            Assert.AreNotEqual(A1.CellStyle, B1.CellStyle);
            Assert.AreEqual(defaultFontIndex, B1.CellStyle.FontIndex);

            wb.Close();
        }

        [Test]
        public void SetFontFromDifferentWorkbook()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            IWorkbook wb2 = _testDataProvider.CreateWorkbook();
            IFont font1 = wb1.CreateFont();
            IFont font2 = wb2.CreateFont();
            // do something to make font1 and font2 different
            // so they are not same or Equal.
            font1.IsItalic = true;
            ICell A1 = wb1.CreateSheet().CreateRow(0).CreateCell(0);

            // okay
            CellUtil.SetFont(A1, font1);

            // font belongs to different workbook
            try
            {
                CellUtil.SetFont(A1, font2);
                Assert.Fail("setFont not allowed if font belongs to a different workbook");
            }
            catch (ArgumentException e)
            {
                if (e.Message.StartsWith("Font does not belong to this workbook"))
                {
                    // expected
                }
                else
                {
                    throw e;
                }
            }
            finally
            {
                wb1.Close();
                wb2.Close();
            }
        }

        /**
         * bug 55555
         * @since POI 3.15 beta 3
         */
        [Test]
        public void SetFillForegroundColorBeforeFillBackgroundColorEnum()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ICell A1 = wb1.CreateSheet().CreateRow(0).CreateCell(0);
            Dictionary<String, Object> properties = new Dictionary<String, Object>();
            // FIXME: Use FillPattern.BRICKS enum
            properties.Add(CellUtil.FILL_PATTERN, FillPattern.Bricks);
            properties.Add(CellUtil.FILL_FOREGROUND_COLOR, IndexedColors.Blue.Index);
            properties.Add(CellUtil.FILL_BACKGROUND_COLOR, IndexedColors.Red.Index);

            CellUtil.SetCellStyleProperties(A1, properties);
            ICellStyle style = A1.CellStyle;
            // FIXME: Use FillPattern.BRICKS enum
            Assert.AreEqual(FillPattern.Bricks, style.FillPattern, "fill pattern");
            Assert.AreEqual(IndexedColors.Blue, IndexedColors.FromInt(style.FillForegroundColor), "fill foreground color");
            Assert.AreEqual(IndexedColors.Red, IndexedColors.FromInt(style.FillBackgroundColor), "fill background color");
        }

    }

}