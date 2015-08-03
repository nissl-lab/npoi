/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using NUnit.Framework;

    using NPOI.SS.UserModel;

    [TestFixture]
    public class TestHSSFOptimiser
    {
        [Test]
        public void TestDoesNoHarmIfNothingToDo()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            // New files start with 4 built in fonts, and 21 built in styles
            Assert.AreEqual(4, wb.NumberOfFonts);
            Assert.AreEqual(21, wb.NumCellStyles);

            // Create a test font and style, and use them
            IFont f = wb.CreateFont();
            f.FontName = ("Testing");
            NPOI.SS.UserModel.ICellStyle s = wb.CreateCellStyle();
            s.SetFont(f);

            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet();
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);
            row.CreateCell(0).CellStyle = (s);

            // Should have one more than the default of each
            Assert.AreEqual(5, wb.NumberOfFonts);
            Assert.AreEqual(22, wb.NumCellStyles);

            // Optimise fonts
            HSSFOptimiser.OptimiseFonts(wb);

            Assert.AreEqual(5, wb.NumberOfFonts);
            Assert.AreEqual(22, wb.NumCellStyles);

            Assert.AreEqual(f, s.GetFont(wb));

            // Optimise styles
            HSSFOptimiser.OptimiseCellStyles(wb);

            Assert.AreEqual(5, wb.NumberOfFonts);
            Assert.AreEqual(22, wb.NumCellStyles);

            Assert.AreEqual(f, s.GetFont(wb));
        }
        [Test]
        public void TestOptimiseFonts()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // Add 6 fonts, some duplicates
            IFont f1 = wb.CreateFont();
            f1.FontHeight = ((short)11);
            f1.FontName = ("Testing");

            IFont f2 = wb.CreateFont();
            f2.FontHeight = ((short)22);
            f2.FontName = ("Also Testing");

            IFont f3 = wb.CreateFont();
            f3.FontHeight = ((short)33);
            f3.FontName = ("Unique");

            IFont f4 = wb.CreateFont();
            f4.FontHeight = ((short)11);
            f4.FontName = ("Testing");

            IFont f5 = wb.CreateFont();
            f5.FontHeight = ((short)22);
            f5.FontName = ("Also Testing");

            IFont f6 = wb.CreateFont();
            f6.FontHeight = ((short)66);
            f6.FontName = ("Also Unique");



            // Use all three of the four in cell styles
            Assert.AreEqual(21, wb.NumCellStyles);

            NPOI.SS.UserModel.ICellStyle cs1 = wb.CreateCellStyle();
            cs1.SetFont(f1);
            Assert.AreEqual(5, cs1.FontIndex);

            NPOI.SS.UserModel.ICellStyle cs2 = wb.CreateCellStyle();
            cs2.SetFont(f4);
            Assert.AreEqual(8, cs2.FontIndex);

            NPOI.SS.UserModel.ICellStyle cs3 = wb.CreateCellStyle();
            cs3.SetFont(f5);
            Assert.AreEqual(9, cs3.FontIndex);

            NPOI.SS.UserModel.ICellStyle cs4 = wb.CreateCellStyle();
            cs4.SetFont(f6);
            Assert.AreEqual(10, cs4.FontIndex);

            Assert.AreEqual(25, wb.NumCellStyles);


            // And three in rich text
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);

            HSSFRichTextString rtr1 = new HSSFRichTextString("Test");
            rtr1.ApplyFont(0, 2, f1);
            rtr1.ApplyFont(3, 4, f2);
            r.CreateCell(0).SetCellValue(rtr1);

            HSSFRichTextString rtr2 = new HSSFRichTextString("AlsoTest");
            rtr2.ApplyFont(0, 2, f3);
            rtr2.ApplyFont(3, 5, f5);
            rtr2.ApplyFont(6, 8, f6);
            r.CreateCell(1).SetCellValue(rtr2);


            // Check what we have now
            Assert.AreEqual(10, wb.NumberOfFonts);
            Assert.AreEqual(25, wb.NumCellStyles);

            // Optimise
            HSSFOptimiser.OptimiseFonts(wb);

            // Check font count
            Assert.AreEqual(8, wb.NumberOfFonts);
            Assert.AreEqual(25, wb.NumCellStyles);

            // Check font use in cell styles
            Assert.AreEqual(5, cs1.FontIndex);
            Assert.AreEqual(5, cs2.FontIndex); // duplicate of 1
            Assert.AreEqual(6, cs3.FontIndex); // duplicate of 2
            Assert.AreEqual(8, cs4.FontIndex); // two have gone


            // And in rich text

            // RTR 1 had f1 and f2, unchanged 
            Assert.AreEqual(5, (r.GetCell(0).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(0));
            Assert.AreEqual(5, (r.GetCell(0).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(1));
            Assert.AreEqual(6, (r.GetCell(0).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(3));
            Assert.AreEqual(6, (r.GetCell(0).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(4));

            // RTR 2 had f3 (unchanged), f5 (=f2) and f6 (moved down)
            Assert.AreEqual(7, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(0));
            Assert.AreEqual(7, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(1));
            Assert.AreEqual(6, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(3));
            Assert.AreEqual(6, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(4));
            Assert.AreEqual(8, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(6));
            Assert.AreEqual(8, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(7));
        }
        [Test]
        public void TestOptimiseStyles()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // Two fonts
            Assert.AreEqual(4, wb.NumberOfFonts);

            IFont f1 = wb.CreateFont();
            f1.FontHeight = ((short)11);
            f1.FontName = ("Testing");

            IFont f2 = wb.CreateFont();
            f2.FontHeight = ((short)22);
            f2.FontName = ("Also Testing");

            Assert.AreEqual(6, wb.NumberOfFonts);


            // Several styles
            Assert.AreEqual(21, wb.NumCellStyles);

            NPOI.SS.UserModel.ICellStyle cs1 = wb.CreateCellStyle();
            cs1.SetFont(f1);

            NPOI.SS.UserModel.ICellStyle cs2 = wb.CreateCellStyle();
            cs2.SetFont(f2);

            NPOI.SS.UserModel.ICellStyle cs3 = wb.CreateCellStyle();
            cs3.SetFont(f1);

            NPOI.SS.UserModel.ICellStyle cs4 = wb.CreateCellStyle();
            cs4.SetFont(f1);
            cs4.Alignment = HorizontalAlignment.CenterSelection;// ((short)22);

            NPOI.SS.UserModel.ICellStyle cs5 = wb.CreateCellStyle();
            cs5.SetFont(f2);
            cs5.Alignment = HorizontalAlignment.Fill; //((short)111);

            NPOI.SS.UserModel.ICellStyle cs6 = wb.CreateCellStyle();
            cs6.SetFont(f2);

            Assert.AreEqual(27, wb.NumCellStyles);


            // Use them
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);

            r.CreateCell(0).CellStyle = (cs1);
            r.CreateCell(1).CellStyle = (cs2);
            r.CreateCell(2).CellStyle = (cs3);
            r.CreateCell(3).CellStyle = (cs4);
            r.CreateCell(4).CellStyle = (cs5);
            r.CreateCell(5).CellStyle = (cs6);
            r.CreateCell(6).CellStyle = (cs1);
            r.CreateCell(7).CellStyle = (cs2);

            Assert.AreEqual(21, ((HSSFCell)r.GetCell(0)).CellValueRecord.XFIndex);
            Assert.AreEqual(26, ((HSSFCell)r.GetCell(5)).CellValueRecord.XFIndex);
            Assert.AreEqual(21, ((HSSFCell)r.GetCell(6)).CellValueRecord.XFIndex);


            // Optimise
            HSSFOptimiser.OptimiseCellStyles(wb);


            // Check
            Assert.AreEqual(6, wb.NumberOfFonts);
            Assert.AreEqual(25, wb.NumCellStyles);

            // cs1 -> 21
            Assert.AreEqual(21, ((HSSFCell)r.GetCell(0)).CellValueRecord.XFIndex);
            // cs2 -> 22
            Assert.AreEqual(22, ((HSSFCell)r.GetCell(1)).CellValueRecord.XFIndex);
            Assert.AreEqual(22, r.GetCell(1).CellStyle.GetFont(wb).FontHeight);
            // cs3 = cs1 -> 21
            Assert.AreEqual(21, ((HSSFCell)r.GetCell(2)).CellValueRecord.XFIndex);
            // cs4 --> 24 -> 23
            Assert.AreEqual(23, ((HSSFCell)r.GetCell(3)).CellValueRecord.XFIndex);
            // cs5 --> 25 -> 24
            Assert.AreEqual(24, ((HSSFCell)r.GetCell(4)).CellValueRecord.XFIndex);
            // cs6 = cs2 -> 22
            Assert.AreEqual(22, ((HSSFCell)r.GetCell(5)).CellValueRecord.XFIndex);
            // cs1 -> 21
            Assert.AreEqual(21, ((HSSFCell)r.GetCell(6)).CellValueRecord.XFIndex);
            // cs2 -> 22
            Assert.AreEqual(22, ((HSSFCell)r.GetCell(7)).CellValueRecord.XFIndex);


            // Add a new duplicate, and two that aren't used
            HSSFCellStyle csD = (HSSFCellStyle)wb.CreateCellStyle();
            csD.SetFont(f1);
            r.CreateCell(8).CellStyle=(csD);

            HSSFFont f3 = (HSSFFont)wb.CreateFont();
            f3.FontHeight=((short)23);
            f3.FontName=("Testing 3");
            HSSFFont f4 = (HSSFFont)wb.CreateFont();
            f4.FontHeight=((short)24);
            f4.FontName=("Testing 4");

            HSSFCellStyle csU1 = (HSSFCellStyle)wb.CreateCellStyle();
            csU1.SetFont(f3);
            HSSFCellStyle csU2 = (HSSFCellStyle)wb.CreateCellStyle();
            csU2.SetFont(f4);

            // Check before the optimise
            Assert.AreEqual(8, wb.NumberOfFonts);
            Assert.AreEqual(28, wb.NumCellStyles);

            // Optimise, should remove the two un-used ones and the one duplicate
            HSSFOptimiser.OptimiseCellStyles(wb);

            // Check
            Assert.AreEqual(8, wb.NumberOfFonts);
            Assert.AreEqual(25, wb.NumCellStyles);

            // csD -> cs1 -> 21
            Assert.AreEqual(21, ((HSSFCell)r.GetCell(8)).CellValueRecord.XFIndex);
        }
        [Test]
        public void TestOptimiseStylesCheckActualStyles()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // Several styles
            Assert.AreEqual(21, wb.NumCellStyles);

            HSSFCellStyle cs1 = (HSSFCellStyle)wb.CreateCellStyle();
            cs1.BorderBottom=(BorderStyle.Thick);

            HSSFCellStyle cs2 = (HSSFCellStyle)wb.CreateCellStyle();
            cs2.BorderBottom=(BorderStyle.DashDot);

            HSSFCellStyle cs3 = (HSSFCellStyle)wb.CreateCellStyle(); // = cs1
            cs3.BorderBottom=(BorderStyle.Thick);

            Assert.AreEqual(24, wb.NumCellStyles);

            // Use them
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            HSSFRow r = (HSSFRow)s.CreateRow(0);

            r.CreateCell(0).CellStyle=(cs1);
            r.CreateCell(1).CellStyle=(cs2);
            r.CreateCell(2).CellStyle=(cs3);

            Assert.AreEqual(21, ((HSSFCell)r.GetCell(0)).CellValueRecord.XFIndex);
            Assert.AreEqual(22, ((HSSFCell)r.GetCell(1)).CellValueRecord.XFIndex);
            Assert.AreEqual(23, ((HSSFCell)r.GetCell(2)).CellValueRecord.XFIndex);

            // Optimise
            HSSFOptimiser.OptimiseCellStyles(wb);

            // Check
            Assert.AreEqual(23, wb.NumCellStyles);

            Assert.AreEqual(BorderStyle.Thick, r.GetCell(0).CellStyle.BorderBottom);
            Assert.AreEqual(BorderStyle.DashDot, r.GetCell(1).CellStyle.BorderBottom);
            Assert.AreEqual(BorderStyle.Thick, r.GetCell(2).CellStyle.BorderBottom);
        }
    }
}