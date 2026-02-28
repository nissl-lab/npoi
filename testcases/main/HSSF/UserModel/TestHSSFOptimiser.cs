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
    using NUnit.Framework;using NUnit.Framework.Legacy;

    using NPOI.SS.UserModel;

    [TestFixture]
    public class TestHSSFOptimiser
    {
        [Test]
        public void TestDoesNoHarmIfNothingToDo()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            // New files start with 4 built in fonts, and 21 built in styles
            ClassicAssert.AreEqual(4, wb.NumberOfFonts);
            ClassicAssert.AreEqual(21, wb.NumCellStyles);

            // Create a test font and style, and use them
            IFont f = wb.CreateFont();
            f.FontName = ("Testing");
            NPOI.SS.UserModel.ICellStyle s = wb.CreateCellStyle();
            s.SetFont(f);

            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet();
            HSSFRow row = (HSSFRow)sheet.CreateRow(0);
            row.CreateCell(0).CellStyle = (s);

            // Should have one more than the default of each
            ClassicAssert.AreEqual(5, wb.NumberOfFonts);
            ClassicAssert.AreEqual(22, wb.NumCellStyles);

            // Optimise fonts
            HSSFOptimiser.OptimiseFonts(wb);

            ClassicAssert.AreEqual(5, wb.NumberOfFonts);
            ClassicAssert.AreEqual(22, wb.NumCellStyles);

            ClassicAssert.AreEqual(f, s.GetFont(wb));

            // Optimise styles
            HSSFOptimiser.OptimiseCellStyles(wb);

            ClassicAssert.AreEqual(5, wb.NumberOfFonts);
            ClassicAssert.AreEqual(22, wb.NumCellStyles);

            ClassicAssert.AreEqual(f, s.GetFont(wb));
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
            ClassicAssert.AreEqual(21, wb.NumCellStyles);

            NPOI.SS.UserModel.ICellStyle cs1 = wb.CreateCellStyle();
            cs1.SetFont(f1);
            ClassicAssert.AreEqual(5, cs1.FontIndex);

            NPOI.SS.UserModel.ICellStyle cs2 = wb.CreateCellStyle();
            cs2.SetFont(f4);
            ClassicAssert.AreEqual(8, cs2.FontIndex);

            NPOI.SS.UserModel.ICellStyle cs3 = wb.CreateCellStyle();
            cs3.SetFont(f5);
            ClassicAssert.AreEqual(9, cs3.FontIndex);

            NPOI.SS.UserModel.ICellStyle cs4 = wb.CreateCellStyle();
            cs4.SetFont(f6);
            ClassicAssert.AreEqual(10, cs4.FontIndex);

            ClassicAssert.AreEqual(25, wb.NumCellStyles);


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
            ClassicAssert.AreEqual(10, wb.NumberOfFonts);
            ClassicAssert.AreEqual(25, wb.NumCellStyles);

            // Optimise
            HSSFOptimiser.OptimiseFonts(wb);

            // Check font count
            ClassicAssert.AreEqual(8, wb.NumberOfFonts);
            ClassicAssert.AreEqual(25, wb.NumCellStyles);

            // Check font use in cell styles
            ClassicAssert.AreEqual(5, cs1.FontIndex);
            ClassicAssert.AreEqual(5, cs2.FontIndex); // duplicate of 1
            ClassicAssert.AreEqual(6, cs3.FontIndex); // duplicate of 2
            ClassicAssert.AreEqual(8, cs4.FontIndex); // two have gone


            // And in rich text

            // RTR 1 had f1 and f2, unchanged 
            ClassicAssert.AreEqual(5, (r.GetCell(0).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(0));
            ClassicAssert.AreEqual(5, (r.GetCell(0).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(1));
            ClassicAssert.AreEqual(6, (r.GetCell(0).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(3));
            ClassicAssert.AreEqual(6, (r.GetCell(0).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(4));

            // RTR 2 had f3 (unchanged), f5 (=f2) and f6 (moved down)
            ClassicAssert.AreEqual(7, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(0));
            ClassicAssert.AreEqual(7, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(1));
            ClassicAssert.AreEqual(6, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(3));
            ClassicAssert.AreEqual(6, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(4));
            ClassicAssert.AreEqual(8, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(6));
            ClassicAssert.AreEqual(8, (r.GetCell(1).RichStringCellValue as HSSFRichTextString).GetFontAtIndex(7));
        }
        [Test]
        public void TestOptimiseStyles()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // Two fonts
            ClassicAssert.AreEqual(4, wb.NumberOfFonts);

            IFont f1 = wb.CreateFont();
            f1.FontHeight = ((short)11);
            f1.FontName = ("Testing");

            IFont f2 = wb.CreateFont();
            f2.FontHeight = ((short)22);
            f2.FontName = ("Also Testing");

            ClassicAssert.AreEqual(6, wb.NumberOfFonts);


            // Several styles
            ClassicAssert.AreEqual(21, wb.NumCellStyles);

            NPOI.SS.UserModel.ICellStyle cs1 = wb.CreateCellStyle();
            cs1.SetFont(f1);

            NPOI.SS.UserModel.ICellStyle cs2 = wb.CreateCellStyle();
            cs2.SetFont(f2);

            NPOI.SS.UserModel.ICellStyle cs3 = wb.CreateCellStyle();
            cs3.SetFont(f1);

            NPOI.SS.UserModel.ICellStyle cs4 = wb.CreateCellStyle();
            cs4.SetFont(f1);
            cs4.Alignment = HorizontalAlignment.Center;// ((short)22);

            NPOI.SS.UserModel.ICellStyle cs5 = wb.CreateCellStyle();
            cs5.SetFont(f2);
            cs5.Alignment = HorizontalAlignment.Fill; //((short)111);

            NPOI.SS.UserModel.ICellStyle cs6 = wb.CreateCellStyle();
            cs6.SetFont(f2);

            ClassicAssert.AreEqual(27, wb.NumCellStyles);


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

            ClassicAssert.AreEqual(21, ((HSSFCell)r.GetCell(0)).CellValueRecord.XFIndex);
            ClassicAssert.AreEqual(26, ((HSSFCell)r.GetCell(5)).CellValueRecord.XFIndex);
            ClassicAssert.AreEqual(21, ((HSSFCell)r.GetCell(6)).CellValueRecord.XFIndex);


            // Optimise
            HSSFOptimiser.OptimiseCellStyles(wb);


            // Check
            ClassicAssert.AreEqual(6, wb.NumberOfFonts);
            ClassicAssert.AreEqual(25, wb.NumCellStyles);

            // cs1 -> 21
            ClassicAssert.AreEqual(21, ((HSSFCell)r.GetCell(0)).CellValueRecord.XFIndex);
            // cs2 -> 22
            ClassicAssert.AreEqual(22, ((HSSFCell)r.GetCell(1)).CellValueRecord.XFIndex);
            ClassicAssert.AreEqual(22, r.GetCell(1).CellStyle.GetFont(wb).FontHeight);
            // cs3 = cs1 -> 21
            ClassicAssert.AreEqual(21, ((HSSFCell)r.GetCell(2)).CellValueRecord.XFIndex);
            // cs4 --> 24 -> 23
            ClassicAssert.AreEqual(23, ((HSSFCell)r.GetCell(3)).CellValueRecord.XFIndex);
            // cs5 --> 25 -> 24
            ClassicAssert.AreEqual(24, ((HSSFCell)r.GetCell(4)).CellValueRecord.XFIndex);
            // cs6 = cs2 -> 22
            ClassicAssert.AreEqual(22, ((HSSFCell)r.GetCell(5)).CellValueRecord.XFIndex);
            // cs1 -> 21
            ClassicAssert.AreEqual(21, ((HSSFCell)r.GetCell(6)).CellValueRecord.XFIndex);
            // cs2 -> 22
            ClassicAssert.AreEqual(22, ((HSSFCell)r.GetCell(7)).CellValueRecord.XFIndex);


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
            ClassicAssert.AreEqual(8, wb.NumberOfFonts);
            ClassicAssert.AreEqual(28, wb.NumCellStyles);

            // Optimise, should remove the two un-used ones and the one duplicate
            HSSFOptimiser.OptimiseCellStyles(wb);

            // Check
            ClassicAssert.AreEqual(8, wb.NumberOfFonts);
            ClassicAssert.AreEqual(25, wb.NumCellStyles);

            // csD -> cs1 -> 21
            ClassicAssert.AreEqual(21, ((HSSFCell)r.GetCell(8)).CellValueRecord.XFIndex);
        }
        [Test]
        public void TestOptimiseStylesCheckActualStyles()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // Several styles
            ClassicAssert.AreEqual(21, wb.NumCellStyles);

            HSSFCellStyle cs1 = (HSSFCellStyle)wb.CreateCellStyle();
            cs1.BorderBottom=(BorderStyle.Thick);

            HSSFCellStyle cs2 = (HSSFCellStyle)wb.CreateCellStyle();
            cs2.BorderBottom=(BorderStyle.DashDot);

            HSSFCellStyle cs3 = (HSSFCellStyle)wb.CreateCellStyle(); // = cs1
            cs3.BorderBottom=(BorderStyle.Thick);

            ClassicAssert.AreEqual(24, wb.NumCellStyles);

            // Use them
            HSSFSheet s = (HSSFSheet)wb.CreateSheet();
            HSSFRow r = (HSSFRow)s.CreateRow(0);

            r.CreateCell(0).CellStyle=(cs1);
            r.CreateCell(1).CellStyle=(cs2);
            r.CreateCell(2).CellStyle=(cs3);

            ClassicAssert.AreEqual(21, ((HSSFCell)r.GetCell(0)).CellValueRecord.XFIndex);
            ClassicAssert.AreEqual(22, ((HSSFCell)r.GetCell(1)).CellValueRecord.XFIndex);
            ClassicAssert.AreEqual(23, ((HSSFCell)r.GetCell(2)).CellValueRecord.XFIndex);

            // Optimise
            HSSFOptimiser.OptimiseCellStyles(wb);

            // Check
            ClassicAssert.AreEqual(23, wb.NumCellStyles);

            ClassicAssert.AreEqual(BorderStyle.Thick, r.GetCell(0).CellStyle.BorderBottom);
            ClassicAssert.AreEqual(BorderStyle.DashDot, r.GetCell(1).CellStyle.BorderBottom);
            ClassicAssert.AreEqual(BorderStyle.Thick, r.GetCell(2).CellStyle.BorderBottom);
        }

        [Test]
        public void TestColumnAndRowStyles()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ClassicAssert.AreEqual(21, wb.NumCellStyles,
                "Usually we have 21 pre-defined styles in a newly created Workbook, see InternalWorkbook.CreateWorkbook()");

            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0);
            row.CreateCell(1);
            row.RowStyle = CreateColorStyle(wb, IndexedColors.Red);

            row = sheet.CreateRow(1);
            row.CreateCell(0);
            row.CreateCell(1);
            row.RowStyle = CreateColorStyle(wb, IndexedColors.Red);

            sheet.SetDefaultColumnStyle(0, CreateColorStyle(wb, IndexedColors.Red));
            sheet.SetDefaultColumnStyle(1, CreateColorStyle(wb, IndexedColors.Red));

            // now the color should be equal for those two columns and rows
            CheckColumnStyles(sheet, 0, 1, false);
            CheckRowStyles(sheet, 0, 1, false);

            // Optimise styles
            HSSFOptimiser.OptimiseCellStyles(wb);

            // We should have the same style-objects for these two columns and rows
            CheckColumnStyles(sheet, 0, 1, true);
            CheckRowStyles(sheet, 0, 1, true);
        }

        [Test]
        public void TestUnusedStyle()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ClassicAssert.AreEqual(21, wb.NumCellStyles,
                "Usually we have 21 pre-defined styles in a newly created Workbook, see InternalWorkbook.CreateWorkbook()");

            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0);
            row.CreateCell(1).CellStyle = CreateColorStyle(wb, IndexedColors.Green);


            row = sheet.CreateRow(1);
            row.CreateCell(0);
            row.CreateCell(1).CellStyle = CreateColorStyle(wb, IndexedColors.Red);


            // Create style. But don't use it.
            for(int i = 0; i < 3; i++)
            {
                // Set Cell Color : AQUA
                CreateColorStyle(wb, IndexedColors.Aqua);
            }

            ClassicAssert.AreEqual(21 + 2 + 3, wb.NumCellStyles);
            ClassicAssert.AreEqual(IndexedColors.Green.Index, sheet.GetRow(0).GetCell(1).CellStyle.FillForegroundColor);
            ClassicAssert.AreEqual(IndexedColors.Red.Index, sheet.GetRow(1).GetCell(1).CellStyle.FillForegroundColor);

            // Optimise styles
            HSSFOptimiser.OptimiseCellStyles(wb);

            ClassicAssert.AreEqual(21 + 2, wb.NumCellStyles);
            ClassicAssert.AreEqual(IndexedColors.Green.Index, sheet.GetRow(0).GetCell(1).CellStyle.FillForegroundColor);
            ClassicAssert.AreEqual(IndexedColors.Red.Index, sheet.GetRow(1).GetCell(1).CellStyle.FillForegroundColor);
        }

        [Test]
        public void TestUnusedStyleOneUsed()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ClassicAssert.AreEqual(21, wb.NumCellStyles,
                "Usually we have 21 pre-defined styles in a newly created Workbook, see InternalWorkbook.CreateWorkbook()");

            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            IRow row = sheet.CreateRow(0);
            row.CreateCell(0);
            row.CreateCell(1).CellStyle = CreateColorStyle(wb, IndexedColors.Green);

            // Create style. But don't use it.
            for(int i = 0; i < 3; i++)
            {
                // Set Cell Color : AQUA
                CreateColorStyle(wb, IndexedColors.Aqua);
            }

            row = sheet.CreateRow(1);
            row.CreateCell(0).CellStyle = CreateColorStyle(wb, IndexedColors.Aqua);
            row.CreateCell(1).CellStyle = CreateColorStyle(wb, IndexedColors.Red);

            ClassicAssert.AreEqual(21 + 3 + 3, wb.NumCellStyles);
            ClassicAssert.AreEqual(IndexedColors.Green.Index, sheet.GetRow(0).GetCell(1).CellStyle.FillForegroundColor);
            ClassicAssert.AreEqual(IndexedColors.Aqua.Index, sheet.GetRow(1).GetCell(0).CellStyle.FillForegroundColor);
            ClassicAssert.AreEqual(IndexedColors.Red.Index, sheet.GetRow(1).GetCell(1).CellStyle.FillForegroundColor);

            // Optimise styles
            HSSFOptimiser.OptimiseCellStyles(wb);

            ClassicAssert.AreEqual(21 + 3, wb.NumCellStyles);
            ClassicAssert.AreEqual(IndexedColors.Green.Index, sheet.GetRow(0).GetCell(1).CellStyle.FillForegroundColor);
            ClassicAssert.AreEqual(IndexedColors.Aqua.Index, sheet.GetRow(1).GetCell(0).CellStyle.FillForegroundColor);
            ClassicAssert.AreEqual(IndexedColors.Red.Index, sheet.GetRow(1).GetCell(1).CellStyle.FillForegroundColor);
        }

        [Test]
        public void TestDefaultColumnStyleWitoutCell()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ClassicAssert.AreEqual(21, wb.NumCellStyles,
                "Usually we have 21 pre-defined styles in a newly created Workbook, see InternalWorkbook.CreateWorkbook()");

            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            //Set CellStyle and RowStyle and ColumnStyle
            for(int i = 0; i < 2; i++)
            {
                sheet.CreateRow(i);
            }

            // Create a test font and style, and use them
            int obj_cnt = wb.NumCellStyles;
            int cnt = wb.NumCellStyles;

            // Set Column Color : Red
            sheet.SetDefaultColumnStyle(3,
                    CreateColorStyle(wb, IndexedColors.Red));
            obj_cnt++;

            // Set Column Color : Red
            sheet.SetDefaultColumnStyle(4,
                    CreateColorStyle(wb, IndexedColors.Red));
            obj_cnt++;

            ClassicAssert.AreEqual(obj_cnt, wb.NumCellStyles);

            // now the color should be equal for those two columns and rows
            CheckColumnStyles(sheet, 3, 4, false);

            // Optimise styles
            HSSFOptimiser.OptimiseCellStyles(wb);

            // We should have the same style-objects for these two columns and rows
            CheckColumnStyles(sheet, 3, 4, true);

            // (GREEN + RED + BLUE + CORAL) + YELLOW(2*2)
            ClassicAssert.AreEqual(cnt + 1, wb.NumCellStyles);
        }

        [Test]
        public void TestUserDefinedStylesAreNeverOptimizedAway()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ClassicAssert.AreEqual(21, wb.NumCellStyles,
                "Usually we have 21 pre-defined styles in a newly created Workbook, see InternalWorkbook.CreateWorkbook()");

            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            //Set CellStyle and RowStyle and ColumnStyle
            for(int i = 0; i < 2; i++)
            {
                sheet.CreateRow(i);
            }

            // Create a test font and style, and use them
            int obj_cnt = wb.NumCellStyles;
            int cnt = wb.NumCellStyles;
            for(int i = 0; i < 3; i++)
            {
                HSSFCellStyle s1 = null;
                if(i == 0)
                {
                    // Set cell color : +2(user style + proxy of it)
                    s1 = (HSSFCellStyle) CreateColorStyle(wb, IndexedColors.Yellow);
                    s1.UserStyleName = "user define";
                    obj_cnt += 2;
                }

                HSSFRow row = sheet.GetRow(1) as HSSFRow;
                row.CreateCell(i).CellStyle = s1;
            }

            // Create style. But don't use it.
            for(int i = 3; i < 6; i++)
            {
                // Set Cell Color : AQUA
                CreateColorStyle(wb, IndexedColors.Aqua);
                obj_cnt++;
            }

            // Set cell color : +2(user style + proxy of it)
            HSSFCellStyle s = (HSSFCellStyle) CreateColorStyle(wb,IndexedColors.Yellow);
            s.UserStyleName = "user define2";
            obj_cnt += 2;

            sheet.CreateRow(10).CreateCell(0).CellStyle = s;

            ClassicAssert.AreEqual(obj_cnt, wb.NumCellStyles);

            // Confirm user style name
            CheckUserStyles(sheet);

            // Optimise styles
            HSSFOptimiser.OptimiseCellStyles(wb);

            // Confirm user style name
            CheckUserStyles(sheet);

            // (GREEN + RED + BLUE + CORAL) + YELLOW(2*2)
            ClassicAssert.AreEqual(cnt + 2 * 2, wb.NumCellStyles);
        }

        [Test]
        public void TestBug57517()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ClassicAssert.AreEqual(21, wb.NumCellStyles,
                "Usually we have 21 pre-defined styles in a newly created Workbook, see InternalWorkbook.CreateWorkbook()");

            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            //Set CellStyle and RowStyle and ColumnStyle
            for(int i = 0; i < 2; i++)
            {
                sheet.CreateRow(i);
            }

            // Create a test font and style, and use them
            int obj_cnt = wb.NumCellStyles;
            int cnt = wb.NumCellStyles;
            for(int i = 0; i < 3; i++)
            {
                // Set Cell Color : GREEN
                HSSFRow row = sheet.GetRow(0) as HSSFRow;
                row.CreateCell(i).CellStyle = CreateColorStyle(wb, IndexedColors.Green);
                obj_cnt++;

                // Set Column Color : Red
                sheet.SetDefaultColumnStyle(i + 3, CreateColorStyle(wb, IndexedColors.Red));
                obj_cnt++;

                // Set Row Color : Blue
                row = sheet.CreateRow(i + 3) as HSSFRow;
                row.RowStyle = CreateColorStyle(wb, IndexedColors.Blue);
                obj_cnt++;

                HSSFCellStyle s1 = null;
                if(i == 0)
                {
                    // Set cell color : +2(user style + proxy of it)
                    s1 = (HSSFCellStyle) CreateColorStyle(wb, IndexedColors.Yellow);
                    s1.UserStyleName = "user define";
                    obj_cnt += 2;
                }

                row = sheet.GetRow(1) as HSSFRow;
                row.CreateCell(i).CellStyle = s1;

            }

            // Create style. But don't use it.
            for(int i = 3; i < 6; i++)
            {
                // Set Cell Color : AQUA
                CreateColorStyle(wb, IndexedColors.Aqua);
                obj_cnt++;
            }

            // Set CellStyle and RowStyle and ColumnStyle
            for(int i = 9; i < 11; i++)
            {
                sheet.CreateRow(i);
            }

            //Set 0 or 255 index of ColumnStyle.
            HSSFCellStyle s = (HSSFCellStyle) CreateColorStyle(wb, IndexedColors.Coral);
            obj_cnt++;
            sheet.SetDefaultColumnStyle(0, s);
            sheet.SetDefaultColumnStyle(255, s);

            // Create a test font and style, and use them
            for(int i = 3; i < 6; i++)
            {
                // Set Cell Color : GREEN
                HSSFRow row = sheet.GetRow(0 + 9) as HSSFRow;
                row.CreateCell(i - 3).CellStyle = CreateColorStyle(wb, IndexedColors.Green);
                obj_cnt++;

                // Set Column Color : Red
                sheet.SetDefaultColumnStyle(i + 3,
                        CreateColorStyle(wb, IndexedColors.Red));
                obj_cnt++;

                // Set Row Color : Blue
                row = sheet.CreateRow(i + 3) as HSSFRow;
                row.RowStyle = CreateColorStyle(wb, IndexedColors.Blue);
                obj_cnt++;

                if(i == 3)
                {
                    // Set cell color : +2(user style + proxy of it)
                    s = (HSSFCellStyle) CreateColorStyle(wb,
                            IndexedColors.Yellow);
                    s.UserStyleName = "user define2";
                    obj_cnt += 2;
                }

                row = sheet.GetRow(1 + 9) as HSSFRow;
                row.CreateCell(i - 3).CellStyle = s;
            }

            ClassicAssert.AreEqual(obj_cnt, wb.NumCellStyles);

            // now the color should be equal for those two columns and rows
            CheckColumnStyles(sheet, 3, 4, false);
            CheckRowStyles(sheet, 3, 4, false);

            // Confirm user style name
            CheckUserStyles(sheet);

            //        out = new FileOutputStream(new File(tmpDirName, "out.xls"));
            //        wb.write(out);
            //        out.Close();

            // Optimise styles
            HSSFOptimiser.OptimiseCellStyles(wb);

            //        out = new FileOutputStream(new File(tmpDirName, "out_optimised.xls"));
            //        wb.write(out);
            //        out.Close();

            // We should have the same style-objects for these two columns and rows
            CheckColumnStyles(sheet, 3, 4, true);
            CheckRowStyles(sheet, 3, 4, true);

            // Confirm user style name
            CheckUserStyles(sheet);

            // (GREEN + RED + BLUE + CORAL) + YELLOW(2*2)
            ClassicAssert.AreEqual(cnt + 4 + 2 * 2, wb.NumCellStyles);
        }

        private void CheckUserStyles(HSSFSheet sheet)
        {
            HSSFCellStyle parentStyle1 = (sheet.GetRow(1).GetCell(0).CellStyle as HSSFCellStyle).ParentStyle;
            ClassicAssert.IsNotNull(parentStyle1);
            ClassicAssert.AreEqual(parentStyle1.UserStyleName, "user define");

            HSSFCellStyle parentStyle10 = (sheet.GetRow(10).GetCell(0).CellStyle as HSSFCellStyle).ParentStyle;
            ClassicAssert.IsNotNull(parentStyle10);
            ClassicAssert.AreEqual(parentStyle10.UserStyleName, "user define2");
        }

        private void CheckColumnStyles(HSSFSheet sheet, int col1, int col2, bool checkEquals)
        {
            // we should have the same color for the column styles
            HSSFCellStyle columnStyle1 = sheet.GetColumnStyle(col1) as HSSFCellStyle;
            ClassicAssert.IsNotNull(columnStyle1);
            HSSFCellStyle columnStyle2 = sheet.GetColumnStyle(col2) as HSSFCellStyle;
            ClassicAssert.IsNotNull(columnStyle2);
            ClassicAssert.AreEqual(columnStyle1.FillForegroundColor, columnStyle2.FillForegroundColor);
            if(checkEquals)
            {
                ClassicAssert.AreEqual(columnStyle1.Index, columnStyle2.Index);
                ClassicAssert.AreEqual(columnStyle1, columnStyle2);
            }
        }

        private void CheckRowStyles(HSSFSheet sheet, int row1, int row2, bool checkEquals)
        {
            // we should have the same color for the row styles
            HSSFCellStyle rowStyle1 = sheet.GetRow(row1).RowStyle as HSSFCellStyle;
            ClassicAssert.IsNotNull(rowStyle1);
            HSSFCellStyle rowStyle2 = sheet.GetRow(row2).RowStyle as HSSFCellStyle;
            ClassicAssert.IsNotNull(rowStyle2);
            ClassicAssert.AreEqual(rowStyle1.FillForegroundColor, rowStyle2.FillForegroundColor);
            if(checkEquals)
            {
                ClassicAssert.AreEqual(rowStyle1.Index, rowStyle2.Index);
                ClassicAssert.AreEqual(rowStyle1, rowStyle2);
            }
        }

        private ICellStyle CreateColorStyle(IWorkbook wb, IndexedColors c)
        {
            ICellStyle cs = wb.CreateCellStyle();
            cs.FillPattern = FillPattern.SolidForeground;
            cs.FillForegroundColor = c.Index;
            return cs;
        }
    }
}