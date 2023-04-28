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

using NUnit.Framework;
using NPOI.SS.UserModel;
using System;
using NPOI.HSSF.Record.CF;
using NPOI.HSSF.Util;
namespace TestCases.SS.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    public class BaseTestFont
    {
        protected ITestDataProvider _testDataProvider;

        public BaseTestFont()
            : this(TestCases.HSSF.HSSFITestDataProvider.Instance)
        {
        }
        /**
         * @param testDataProvider an object that provides test data in HSSF / XSSF specific way
         */
        protected BaseTestFont(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }
        public void BaseTestDefaultFont(String defaultName, short defaultSize, short defaultColor)
        {
            //get default font and check against default value
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            IFont fontFind = workbook.FindFont(false, defaultColor, defaultSize, defaultName, false, false, FontSuperScript.None, FontUnderlineType.None);
            Assert.IsNotNull(fontFind);

            //get default font, then change 2 values and check against different values (height Changes)
            IFont font = workbook.CreateFont();
            font.IsBold = true;
            Assert.IsTrue(font.IsBold);
            font.Underline = FontUnderlineType.Double;
            Assert.AreEqual(FontUnderlineType.Double, font.Underline);
            font.FontHeightInPoints = ((short)15);
            Assert.AreEqual(15 * 20, font.FontHeight);
            Assert.AreEqual(15, font.FontHeightInPoints);
            fontFind = workbook.FindFont(true, defaultColor, (short)(15 * 20), defaultName, false, false, FontSuperScript.None, FontUnderlineType.Double);
            Assert.IsNotNull(fontFind);
        }
        [Test]
        public void TestNumberOfFonts()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            int num0 = wb.NumberOfFonts;

            IFont f1 = wb.CreateFont();
            f1.IsBold = true;
            short idx1 = f1.Index;
            wb.CreateCellStyle().SetFont(f1);

            IFont f2 = wb.CreateFont();
            f2.Underline = FontUnderlineType.Double;
            short idx2 = f2.Index;
            wb.CreateCellStyle().SetFont(f2);

            IFont f3 = wb.CreateFont();
            f3.FontHeightInPoints = ((short)23);
            short idx3 = f3.Index;
            wb.CreateCellStyle().SetFont(f3);

            Assert.AreEqual(num0 + 3, wb.NumberOfFonts);
            Assert.IsTrue(wb.GetFontAt(idx1).IsBold);
            Assert.AreEqual(FontUnderlineType.Double, wb.GetFontAt(idx2).Underline);
            Assert.AreEqual(23, wb.GetFontAt(idx3).FontHeightInPoints);
        }

        /**
         * Tests that we can define fonts to a new
         *  file, save, load, and still see them
         */
        [Test]
        public void TestCreateSave()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s1 = wb.CreateSheet();
            IRow r1 = s1.CreateRow(0);
            ICell r1c1 = r1.CreateCell(0);
            r1c1.SetCellValue(2.2);

            int num0 = wb.NumberOfFonts;

            IFont font = wb.CreateFont();
            font.IsBold = true;
            font.IsStrikeout = (true);
            font.Color = (HSSFColor.Yellow.Index);
            font.FontName = ("Courier");
            short font1Idx = font.Index;
            wb.CreateCellStyle().SetFont(font);
            Assert.AreEqual(num0 + 1, wb.NumberOfFonts);

            ICellStyle cellStyleTitle = wb.CreateCellStyle();
            cellStyleTitle.SetFont(font);
            r1c1.CellStyle = (cellStyleTitle);

            // Save and re-load
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            s1 = wb.GetSheetAt(0);

            Assert.AreEqual(num0 + 1, wb.NumberOfFonts);
            short idx = s1.GetRow(0).GetCell(0).CellStyle.FontIndex;
            IFont fnt = wb.GetFontAt(idx);
            Assert.IsNotNull(fnt);
            Assert.AreEqual(HSSFColor.Yellow.Index, fnt.Color);
            Assert.AreEqual("Courier", fnt.FontName);

            // Now add an orphaned one
            IFont font2 = wb.CreateFont();
            font2.IsItalic = (true);
            font2.FontHeightInPoints = (short)15;
            short font2Idx = font2.Index;
            wb.CreateCellStyle().SetFont(font2);
            Assert.AreEqual(num0 + 2, wb.NumberOfFonts);

            // Save and re-load
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            s1 = wb.GetSheetAt(0);

            Assert.AreEqual(num0 + 2, wb.NumberOfFonts);
            Assert.IsNotNull(wb.GetFontAt(font1Idx));
            Assert.IsNotNull(wb.GetFontAt(font2Idx));

            Assert.AreEqual(15, wb.GetFontAt(font2Idx).FontHeightInPoints);
            Assert.IsTrue(wb.GetFontAt(font2Idx).IsItalic);
        }

        /**
         * Test that fonts get Added properly
         *
         * @see NPOI.HSSF.usermodel.TestBugs#test45338()
         */
        [Test]
        public void Test45338()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            int num0 = wb.NumberOfFonts;

            ISheet s = wb.CreateSheet();
            s.CreateRow(0);
            s.CreateRow(1);
            s.GetRow(0).CreateCell(0);
            s.GetRow(1).CreateCell(0);

            //default font
            IFont f1 = wb.GetFontAt((short)0);
            Assert.IsFalse(f1.IsBold);

            // Check that asking for the same font
            //  multiple times gives you the same thing.
            // Otherwise, our Tests wouldn't work!
            Assert.AreSame(wb.GetFontAt((short)0), wb.GetFontAt((short)0));

            // Look for a new font we have
            //  yet to add
            Assert.IsNull(
                wb.FindFont(
                    true, (short)123, (short)(22 * 20),
                    "Thingy", false, true, FontSuperScript.Sub, FontUnderlineType.Double
                )
            );

            IFont nf = wb.CreateFont();
            short nfIdx = nf.Index;
            Assert.AreEqual(num0 + 1, wb.NumberOfFonts);

            Assert.AreEqual(nf, wb.GetFontAt(nfIdx));

            nf.IsBold = true;
            nf.Color = (short)123;
            nf.FontHeightInPoints = (short)22;
            nf.FontName = ("Thingy");
            nf.IsItalic = (false);
            nf.IsStrikeout = (true);
            nf.TypeOffset = FontSuperScript.Sub;
            nf.Underline = FontUnderlineType.Double;

            Assert.AreEqual(num0 + 1, wb.NumberOfFonts);
            Assert.AreEqual(nf, wb.GetFontAt(nfIdx));

            Assert.AreEqual(wb.GetFontAt(nfIdx), wb.GetFontAt(nfIdx));
            Assert.IsTrue(wb.GetFontAt((short)0) != wb.GetFontAt(nfIdx));

            // Find it now
            Assert.IsNotNull(
                wb.FindFont(
                    true, (short)123, (short)(22 * 20),
                    "Thingy", false, true, FontSuperScript.Sub, FontUnderlineType.Double
                )
            );
            Assert.AreSame(nf,
                   wb.FindFont(
                       true, (short)123, (short)(22 * 20),
                       "Thingy", false, true, FontSuperScript.Sub, FontUnderlineType.Double
                   )
            );
        }
    }

}



