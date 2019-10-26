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

namespace NPOI.SS.Util
{
    using System;
    using System.Collections.Generic;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;

    /**
* Tests Spreadsheet CellUtil
*
* @see NPOI.SS.Util.CellUtil
*/
    [TestFixture]
    public class TestCellUtil
    {
        [Test]
        public void SetCellStyleProperty()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(0);

            // Add a border should create a new style
            int styCnt1 = wb.NumCellStyles;
            CellUtil.SetCellStyleProperty(c, wb, CellUtil.BORDER_BOTTOM, BorderStyle.Thin);
            int styCnt2 = wb.NumCellStyles;
            Assert.AreEqual(styCnt2, styCnt1 + 1);

            // Add same border to another cell, should not create another style
            c = r.CreateCell(1);
            CellUtil.SetCellStyleProperty(c, wb, CellUtil.BORDER_BOTTOM, BorderStyle.Thin);
            int styCnt3 = wb.NumCellStyles;
            Assert.AreEqual(styCnt3, styCnt2);

            wb.Close();
        }

        [Test]
        public void SetCellStyleProperties()
        {
            IWorkbook wb = new HSSFWorkbook();
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
            CellUtil.SetCellStyleProperties(c, props);
            int styCnt2 = wb.NumCellStyles;
            Assert.AreEqual(styCnt2, styCnt1 + 1);

            // Add same border another to same cell, should not create another style
            c = r.CreateCell(1);
            CellUtil.SetCellStyleProperties(c, props);
            int styCnt3 = wb.NumCellStyles;
            Assert.AreEqual(styCnt3, styCnt2);

            wb.Close();
        }

        [Test]
        public void GetRow()
        {
            IWorkbook wb = new HSSFWorkbook();
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
            IWorkbook wb = new HSSFWorkbook();
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
            IWorkbook wb = new HSSFWorkbook();
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
        public void SetAlignment()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row = sh.CreateRow(0);
            ICell A1 = row.CreateCell(0);
            ICell B1 = row.CreateCell(1);

            // Assumptions
            Assert.AreEqual(A1.CellStyle, B1.CellStyle);
            // should be Assert.AreSame, but a new HSSFCellStyle is returned for each GetCellStyle() call. 
            // HSSFCellStyle wraps an underlying style record, and the underlying
            // style record is the same between multiple GetCellStyle() calls.
            Assert.AreEqual(HorizontalAlignment.General, A1.CellStyle.Alignment);
            Assert.AreEqual(HorizontalAlignment.General, B1.CellStyle.Alignment);

            // Get/set alignment modifies the cell's style
            CellUtil.SetAlignment(A1, wb, (int)HorizontalAlignment.Right);
            Assert.AreEqual(HorizontalAlignment.Right, A1.CellStyle.Alignment);

            // Get/set alignment doesn't affect the style of cells with
            // the same style prior to modifying the style
            Assert.AreNotEqual(A1.CellStyle, B1.CellStyle);
            Assert.AreEqual(HorizontalAlignment.General, B1.CellStyle.Alignment);

            wb.Close();
        }

        [Test]
        public void SetFont()
        {
            IWorkbook wb = new HSSFWorkbook();
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
            CellUtil.SetFont(A1, wb, font);
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
            IWorkbook wb1 = new HSSFWorkbook();
            IWorkbook wb2 = new HSSFWorkbook();
            IFont font1 = wb1.CreateFont();
            IFont font2 = wb2.CreateFont();
            // do something to make font1 and font2 different
            // so they are not same or Equal.
            font1.IsItalic = true;
            ICell A1 = wb1.CreateSheet().CreateRow(0).CreateCell(0);

            // okay
            CellUtil.SetFont(A1, wb1, font1);

            // font belongs to different workbook
            try
            {
                CellUtil.SetFont(A1, wb1, font2);
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

            // cell belongs to different workbook
            try
            {
                CellUtil.SetFont(A1, wb2, font2);
                Assert.Fail("setFont not allowed if cell belongs to a different workbook");
            }
            catch (ArgumentException e)
            {
                if (e.Message.StartsWith("Cannot Set cell style property. Cell does not belong to workbook."))
                {
                    // expected
                }
                else
                {
                    throw e;
                }
            }
        }
    }

}