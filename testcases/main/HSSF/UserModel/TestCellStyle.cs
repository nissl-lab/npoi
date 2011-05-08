
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
        

/*
 * TestCellStyle.java
 *
 * Created on December 11, 2001, 5:51 PM
 */
namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.Util;
    using NPOI.HSSF.UserModel;


    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;

    /**
     * Class to Test cell styling functionality
     *
     * @author Andrew C. Oliver
     */
    [TestClass]
    public class TestCellStyle
    {

        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        /** Creates a new instance of TestCellStyle */

        public TestCellStyle()
        {

        }

        /**
         * TEST NAME:  Test Write Sheet Font <P>
         * OBJECTIVE:  Test that HSSF can Create a simple spreadsheet with numeric and string values and styled with fonts.<P>
         * SUCCESS:    HSSF Creates a sheet.  Filesize Matches a known good.  NPOI.SS.UserModel.Sheet objects
         *             Last row, first row is Tested against the correct values (99,0).<P>
         * FAILURE:    HSSF does not Create a sheet or excepts.  Filesize does not Match the known good.
         *             NPOI.SS.UserModel.Sheet last row or first row is incorrect.             <P>
         *
         */
        [TestMethod]
        public void TestWriteSheetFont()
    {
        string             filepath = TempFile.GetTempFilePath("TestWriteSheetFont",
                                                    ".xls");
        FileStream out1  = new FileStream(filepath,FileMode.OpenOrCreate);
        HSSFWorkbook     wb   = new HSSFWorkbook();
        NPOI.SS.UserModel.Sheet        s    = wb.CreateSheet();
        Row          r    = null;
        Cell         c    = null;
        Font         fnt  = wb.CreateFont();
        NPOI.SS.UserModel.CellStyle    cs   = wb.CreateCellStyle();

        fnt.Color=(NPOI.HSSF.Util.HSSFColor.RED.index);
        fnt.Boldweight= (short)FontBoldWeight.BOLD;
        cs.SetFont(fnt);
        for (short rownum = ( short ) 0; rownum < 100; rownum++)
        {
            r = s.CreateRow(rownum);

            // r.SetRowNum(( short ) rownum);
            for (short cellnum = ( short ) 0; cellnum < 50; cellnum += 2)
            {
                c = r.CreateCell(cellnum);
                c.SetCellValue(rownum * 10000 + cellnum
                               + ((( double ) rownum / 1000)
                                  + (( double ) cellnum / 10000)));
                c = r.CreateCell(cellnum + 1);
                c.SetCellValue("TEST");
                c.CellStyle = (cs);
            }
        }
        wb.Write(out1);
        out1.Close();
        SanityChecker sanityChecker = new SanityChecker();
        sanityChecker.CheckHSSFWorkbook(wb);
        Assert.AreEqual(99, s.LastRowNum, "LAST ROW == 99");
        Assert.AreEqual(0, s.FirstRowNum, "FIRST ROW == 0");

        // assert((s.LastRowNum == 99));
    }

        /**
         * Tests that is creating a file with a date or an calendar works correctly.
         */
        [TestMethod]
        public void TestDataStyle()
            
    {
        string             filepath = TempFile.GetTempFilePath("TestWriteSheetStyleDate",
                                                    ".xls");
        FileStream out1  = new FileStream(filepath,FileMode.OpenOrCreate);
        HSSFWorkbook     wb   = new HSSFWorkbook();
        NPOI.SS.UserModel.Sheet        s    = wb.CreateSheet();
        NPOI.SS.UserModel.CellStyle    cs   = wb.CreateCellStyle();
        Row row = s.CreateRow(0);

        // with Date:
        Cell cell = row.CreateCell(1);
        cs.DataFormat=(HSSFDataFormat.GetBuiltinFormat("m/d/yy"));
        cell.CellStyle = (cs);
        cell.SetCellValue(DateTime.Now);

        // with Calendar:
        cell = row.CreateCell(2);
        cs.DataFormat=(HSSFDataFormat.GetBuiltinFormat("m/d/yy"));
        cell.CellStyle = (cs);
        cell.SetCellValue(DateTime.Now);

        wb.Write(out1);
        out1.Close();
        SanityChecker sanityChecker = new SanityChecker();
        sanityChecker.CheckHSSFWorkbook(wb);

        Assert.AreEqual(0, s.LastRowNum, "LAST ROW ");
        Assert.AreEqual(0, s.FirstRowNum,"FIRST ROW ");

    }
        [TestMethod]
        public void TestHashEquals()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.Sheet s = wb.CreateSheet();
            NPOI.SS.UserModel.CellStyle cs1 = wb.CreateCellStyle();
            NPOI.SS.UserModel.CellStyle cs2 = wb.CreateCellStyle();
            Row row = s.CreateRow(0);
            Cell cell1 = row.CreateCell(1);
            Cell cell2 = row.CreateCell(2);

            cs1.DataFormat = (HSSFDataFormat.GetBuiltinFormat("m/d/yy"));
            cs2.DataFormat = (HSSFDataFormat.GetBuiltinFormat("m/dd/yy"));

            cell1.CellStyle = (cs1);
            cell1.SetCellValue(DateTime.Now);

            cell2.CellStyle = (cs2);
            cell2.SetCellValue(DateTime.Now);

            Assert.AreEqual(cs1.GetHashCode(), cs1.GetHashCode());
            Assert.AreEqual(cs2.GetHashCode(), cs2.GetHashCode());
            Assert.IsTrue(cs1.Equals(cs1));
            Assert.IsTrue(cs2.Equals(cs2));

            // Change cs1, hash will alter
            int hash1 = cs1.GetHashCode();
            cs1.DataFormat = (HSSFDataFormat.GetBuiltinFormat("m/dd/yy"));
            Assert.IsFalse(hash1 == cs1.GetHashCode());
        }

        /**
         * TEST NAME:  Test Write Sheet Style <P>
         * OBJECTIVE:  Test that HSSF can Create a simple spreadsheet with numeric and string values and styled with colors
         *             and borders.<P>
         * SUCCESS:    HSSF Creates a sheet.  Filesize Matches a known good.  NPOI.SS.UserModel.Sheet objects
         *             Last row, first row is Tested against the correct values (99,0).<P>
         * FAILURE:    HSSF does not Create a sheet or excepts.  Filesize does not Match the known good.
         *             NPOI.SS.UserModel.Sheet last row or first row is incorrect.             <P>
         *
         */
        [TestMethod]
        public void TestWriteSheetStyle()
    {
        string             filepath = TempFile.GetTempFilePath("TestWriteSheetStyle",
                                                    ".xls");
        FileStream out1  = new FileStream(filepath,FileMode.OpenOrCreate);
        HSSFWorkbook     wb   = new HSSFWorkbook();
        NPOI.SS.UserModel.Sheet        s    = wb.CreateSheet();
        Row          r    = null;
        Cell         c    = null;
        Font         fnt  = wb.CreateFont();
        NPOI.SS.UserModel.CellStyle cs = wb.CreateCellStyle();
        NPOI.SS.UserModel.CellStyle cs2 = wb.CreateCellStyle();

        cs.BorderBottom= (CellBorderType.THIN);
        cs.BorderLeft= (CellBorderType.THIN);
        cs.BorderRight= (CellBorderType.THIN);
        cs.BorderTop= (CellBorderType.THIN);
        cs.FillForegroundColor= ( short ) 0xA;
        cs.FillPattern = FillPatternType.SOLID_FOREGROUND;
        fnt.Color= ( short ) 0xf;
        fnt.IsItalic= (true);
        cs2.FillForegroundColor= ( short ) 0x0;
        cs2.FillPattern= FillPatternType.SOLID_FOREGROUND;
        cs2.SetFont(fnt);
        for (short rownum = ( short ) 0; rownum < 100; rownum++)
        {
            r = s.CreateRow(rownum);

            // r.SetRowNum(( short ) rownum);
            for (short cellnum = ( short ) 0; cellnum < 50; cellnum += 2)
            {
                c = r.CreateCell(cellnum);
                c.SetCellValue(rownum * 10000 + cellnum
                               + ((( double ) rownum / 1000)
                                  + (( double ) cellnum / 10000)));
                c.CellStyle = (cs);
                c = r.CreateCell(cellnum + 1);
                c.SetCellValue("TEST");
                c.CellStyle = (cs2);
            }
        }
        wb.Write(out1);
        out1.Close();
        SanityChecker sanityChecker = new SanityChecker();
        sanityChecker.CheckHSSFWorkbook(wb);
        Assert.AreEqual(99, s.LastRowNum, "LAST ROW == 99");
        Assert.AreEqual(0, s.FirstRowNum, "FIRST ROW == 0");

        // assert((s.LastRowNum == 99));
    }

        /**
         * Cloning one NPOI.SS.UserModel.CellType onto Another, same
         *  HSSFWorkbook
         */
        [TestMethod]
        public void TestCloneStyleSameWB()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Font fnt = wb.CreateFont();
            fnt.FontName=("TestingFont");
            Assert.AreEqual(5, wb.NumberOfFonts);

            NPOI.SS.UserModel.CellStyle orig = wb.CreateCellStyle();
            orig.Alignment=(HorizontalAlignment.RIGHT);
            orig.SetFont(fnt);
            orig.DataFormat=((short)18);

            Assert.AreEqual(HorizontalAlignment.RIGHT,orig.Alignment);
            Assert.AreEqual(fnt,orig.GetFont(wb));
            Assert.AreEqual(18,orig.DataFormat);

            NPOI.SS.UserModel.CellStyle clone = wb.CreateCellStyle();
            Assert.AreNotEqual(HorizontalAlignment.RIGHT , clone.Alignment);
            Assert.AreNotEqual(fnt, clone.GetFont(wb));
            Assert.AreNotEqual(18, clone.DataFormat);

            clone.CloneStyleFrom(orig);
            Assert.AreEqual(HorizontalAlignment.RIGHT, clone.Alignment);
            Assert.AreEqual(fnt , clone.GetFont(wb));
            Assert.AreEqual(18,clone.DataFormat);
            Assert.AreEqual(5, wb.NumberOfFonts);
        }

        /**
         * Cloning one NPOI.SS.UserModel.CellType onto Another, across
         *  two different HSSFWorkbooks
         */
        [TestMethod]
        public void TestCloneStyleDiffWB()
        {
            HSSFWorkbook wbOrig = new HSSFWorkbook();

            Font fnt = wbOrig.CreateFont();
            fnt.FontName=("TestingFont");
            Assert.AreEqual(5, wbOrig.NumberOfFonts);

            DataFormat fmt = wbOrig.CreateDataFormat();
            fmt.GetFormat("MadeUpOne");
            fmt.GetFormat("MadeUpTwo");

            NPOI.SS.UserModel.CellStyle orig = wbOrig.CreateCellStyle();
            orig.Alignment = (HorizontalAlignment.RIGHT);
            orig.SetFont(fnt);
            orig.DataFormat=(fmt.GetFormat("Test##"));

            Assert.AreEqual(HorizontalAlignment.RIGHT, orig.Alignment);
            Assert.AreEqual(fnt,orig.GetFont(wbOrig));
            Assert.AreEqual(fmt.GetFormat("Test##") , orig.DataFormat);

            // Now a style on another workbook
            HSSFWorkbook wbClone = new HSSFWorkbook();
            Assert.AreEqual(4, wbClone.NumberOfFonts);
            DataFormat fmtClone = wbClone.CreateDataFormat();

            NPOI.SS.UserModel.CellStyle clone = wbClone.CreateCellStyle();
            Assert.AreEqual(4, wbClone.NumberOfFonts);

            Assert.AreNotEqual(HorizontalAlignment.RIGHT,clone.Alignment);
            Assert.AreNotEqual("TestingFont", clone.GetFont(wbClone).FontName);

            clone.CloneStyleFrom(orig);
            Assert.AreEqual(HorizontalAlignment.RIGHT, clone.Alignment);
            Assert.AreEqual("TestingFont" ,clone.GetFont(wbClone).FontName);
            Assert.AreEqual(fmtClone.GetFormat("Test##"),clone.DataFormat);
            Assert.AreNotEqual(fmtClone.GetFormat("Test##") , fmt.GetFormat("Test##"));
            Assert.AreEqual(5, wbClone.NumberOfFonts);
        }
        [TestMethod]
        public void TestStyleNames()
        {
            HSSFWorkbook wb = OpenSample("WithExtendedStyles.xls");
            NPOI.SS.UserModel.Sheet s = wb.GetSheetAt(0);
            Cell c1 = s.GetRow(0).GetCell(0);
            Cell c2 = s.GetRow(1).GetCell(0);
            Cell c3 = s.GetRow(2).GetCell(0);

            HSSFCellStyle cs1 = (HSSFCellStyle)c1.CellStyle;
            HSSFCellStyle cs2 = (HSSFCellStyle)c2.CellStyle;
            HSSFCellStyle cs3 = (HSSFCellStyle)c3.CellStyle;

            Assert.IsNotNull(cs1);
            Assert.IsNotNull(cs2);
            Assert.IsNotNull(cs3);

            // Check we got the styles we'd expect
            Assert.AreEqual(10, cs1.GetFont(wb).FontHeightInPoints);
            Assert.AreEqual(9, cs2.GetFont(wb).FontHeightInPoints);
            Assert.AreEqual(12, cs3.GetFont(wb).FontHeightInPoints);

            Assert.AreEqual(15, cs1.Index);
            Assert.AreEqual(23, cs2.Index);
            Assert.AreEqual(24, cs3.Index);

            Assert.IsNull(cs1.ParentStyle);
            Assert.IsNotNull(cs2.ParentStyle);
            Assert.IsNotNull(cs3.ParentStyle);

            Assert.AreEqual(21, cs2.ParentStyle.Index);
            Assert.AreEqual(22, cs3.ParentStyle.Index);

            // Now Check we can get style records for 
            //  the parent ones
            Assert.IsNull(wb.Workbook.GetStyleRecord(15));
            Assert.IsNull(wb.Workbook.GetStyleRecord(23));
            Assert.IsNull(wb.Workbook.GetStyleRecord(24));

            Assert.IsNotNull(wb.Workbook.GetStyleRecord(21));
            Assert.IsNotNull(wb.Workbook.GetStyleRecord(22));

            // Now Check the style names
            Assert.AreEqual(null, cs1.UserStyleName);
            Assert.AreEqual(null, cs2.UserStyleName);
            Assert.AreEqual(null, cs3.UserStyleName);
            Assert.AreEqual("style1", cs2.ParentStyle.UserStyleName);
            Assert.AreEqual("style2", cs3.ParentStyle.UserStyleName);
        }
    }
}
