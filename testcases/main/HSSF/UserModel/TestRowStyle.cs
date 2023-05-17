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


/*
 * TestRowStyle.java
 *
 * Created on May 20, 2005
 */
namespace TestCases.HSSF.UserModel
{
    using System.IO;
    using System;

    using NUnit.Framework;

    using TestCases.HSSF;
    using NPOI.Util;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Class to Test row styling functionality
     *
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     */

    [TestFixture]
    public class TestRowStyle
    {

        /** Creates a new instance of TestCellStyle */

        public TestRowStyle()
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
        [Test]
        public void TestWriteSheetFont()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();
            IRow r = null;
            //ICell c = null;
            IFont fnt = wb.CreateFont();
            NPOI.SS.UserModel.ICellStyle cs = wb.CreateCellStyle();

            fnt.Color = (NPOI.HSSF.Util.HSSFColor.Red.Index);
            fnt.IsBold = true;
            cs.SetFont(fnt);
            for (short rownum = (short)0; rownum < 100; rownum++)
            {
                r = s.CreateRow(rownum);
                r.RowStyle = (cs);
                r.CreateCell(0);
            }
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            SanityChecker sanityChecker = new SanityChecker();
            sanityChecker.CheckHSSFWorkbook(wb);
            Assert.AreEqual(99, s.LastRowNum, "LAST ROW == 99");
            Assert.AreEqual(0, s.FirstRowNum, "FIRST ROW == 0");
        }

        /**
         * Tests that is creating a file with a date or an calendar works correctly.
         */
        [Test]
        public void TestDataStyle()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();
            NPOI.SS.UserModel.ICellStyle cs = wb.CreateCellStyle();
            IRow row = s.CreateRow((short)0);

            // with Date:
            cs.DataFormat = (HSSFDataFormat.GetBuiltinFormat("m/d/yy"));
            row.RowStyle = (cs);
            row.CreateCell(0);


            // with Calendar:
            row = s.CreateRow((short)1);
            cs.DataFormat = (HSSFDataFormat.GetBuiltinFormat("m/d/yy"));
            row.RowStyle = (cs);
            row.CreateCell(0);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            SanityChecker sanityChecker = new SanityChecker();
            sanityChecker.CheckHSSFWorkbook(wb);

            Assert.AreEqual(1, s.LastRowNum, "LAST ROW ");
            Assert.AreEqual(0, s.FirstRowNum, "FIRST ROW ");

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
        [Test]
        public void TestWriteSheetStyle()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet();
            IRow r = null;
            IFont fnt = wb.CreateFont();
            NPOI.SS.UserModel.ICellStyle cs = wb.CreateCellStyle();
            NPOI.SS.UserModel.ICellStyle cs2 = wb.CreateCellStyle();

            cs.BorderBottom = (BorderStyle.Thin);
            cs.BorderLeft = (BorderStyle.Thin);
            cs.BorderRight = (BorderStyle.Thin);
            cs.BorderTop = (BorderStyle.Thin);
            cs.FillForegroundColor = ((short)0xA);
            cs.FillPattern = FillPattern.SolidForeground;
            fnt.Color = ((short)0xf);
            fnt.IsItalic = (true);
            cs2.FillForegroundColor = ((short)0x0);
            cs2.FillPattern = FillPattern.SolidForeground;
            cs2.SetFont(fnt);
            for (short rownum = (short)0; rownum < 100; rownum++)
            {
                r = s.CreateRow(rownum);
                r.RowStyle = (cs);
                r.CreateCell(0);

                rownum++;
                if (rownum >= 100)
                    break; // I feel too lazy to Check if this isreqd :-/ 

                r = s.CreateRow(rownum);
                r.RowStyle = (cs2);
                r.CreateCell(0);
            }
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            SanityChecker sanityChecker = new SanityChecker();
            sanityChecker.CheckHSSFWorkbook(wb);
            Assert.AreEqual(99, s.LastRowNum, "LAST ROW == 99");
            Assert.AreEqual(0, s.FirstRowNum, "FIRST ROW == 0");

            s = wb.GetSheetAt(0);
            Assert.IsNotNull(s, "Sheet is not null");

            for (short rownum = (short)0; rownum < 100; rownum++)
            {
                r = s.GetRow(rownum);
                Assert.IsNotNull(r, "Row is not null");

                cs = r.RowStyle;
                Assert.AreEqual(cs.BorderBottom, BorderStyle.Thin, "Bottom Border Style for row: ");
                Assert.AreEqual(cs.BorderLeft, BorderStyle.Thin, "Left Border Style for row: ");
                Assert.AreEqual(cs.BorderRight, BorderStyle.Thin, "Right Border Style for row: ");
                Assert.AreEqual(cs.BorderTop, BorderStyle.Thin, "Top Border Style for row: ");
                Assert.AreEqual(0xA, cs.FillForegroundColor, "FillForegroundColor for row: ");
                Assert.AreEqual((short)0x1, (short)cs.FillPattern, "FillPattern for row: ");

                rownum++;
                if (rownum >= 100)
                    break; // I feel too lazy to Check if this isreqd :-/ 

                r = s.GetRow(rownum);
                Assert.IsNotNull(r, "Row is not null");
                cs2 = r.RowStyle;
                Assert.AreEqual(cs2.FillForegroundColor, (short)0x0, "FillForegroundColor for row: ");
                Assert.AreEqual((short)cs2.FillPattern, (short)0x1, "FillPattern for row: ");
            }
        }
    }
}