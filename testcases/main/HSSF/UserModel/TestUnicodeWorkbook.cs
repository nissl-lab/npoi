/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.HSSF.UserModel;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using NUnit.Framework;

    [TestFixture]
    public class TestUnicodeWorkbook
    {

        public TestUnicodeWorkbook()
        {

        }

        /** Tests that all of the unicode capable string fields can be set, written and then read back
         * 
         *
         */
        //[Test]
        //public void TestUnicodeInAll()
        //{
        //    HSSFWorkbook wb = new HSSFWorkbook();
        //    //Create a unicode dataformat (contains euro symbol)
        //    DataFormat df = wb.CreateDataFormat();
        //    String formatStr = "_([$\u20ac-2]\\\\\\ * #,##0.00_);_([$\u20ac-2]\\\\\\ * \\\\\\(#,##0.00\\\\\\);_([$\u20ac-2]\\\\\\ *\\\"\\-\\\\\"??_);_(@_)";
        //    short fmt = df.GetFormat(formatStr);

        //    //Create a unicode sheet name (euro symbol)
        //    NPOI.SS.UserModel.Sheet s = wb.CreateSheet("\u20ac");

        //    //Set a unicode header (you guessed it the euro symbol)
        //    HSSFHeader h = s.Header;
        //    h.Center = ("\u20ac");
        //    h.Left = ("\u20ac");
        //    h.Right = ("\u20ac");

        //    //Set a unicode footer
        //    HSSFFooter f = s.Footer;
        //    f.Center = ("\u20ac");
        //    f.Left = ("\u20ac");
        //    f.Right = ("\u20ac");

        //    Row r = s.CreateRow(0);
        //    Cell c = r.CreateCell(1);
        //    c.SetCellValue(12.34);
        //    c.CellStyle.DataFormat = (fmt);

        //    Cell c2 = r.CreateCell(2);
        //    c.SetCellValue(new HSSFRichTextString("\u20ac"));

        //    Cell c3 = r.CreateCell(3);
        //    String formulaString = "TEXT(12.34,\"\u20ac###,##\")";
        //    c3.CellFormula = (formulaString);


        //    string path = NPOI.Util.TempFile.GetTempFilePath("unicode", "Test.xls");
        //    FileStream tempFile = File.Create(path);
        //    wb.Write(tempFile);
        //    wb = null;
        //    tempFile.Close();
        //    FileStream in1 = new FileStream(path,FileMode.Open);
        //    wb = new HSSFWorkbook(in1);

        //    //Test the sheetname
        //    s = wb.GetSheet("\u20ac");
        //    Assert.IsNotNull(s);

        //    //Test the header
        //    h = s.Header;
        //    Assert.AreEqual(h.Center, "\u20ac");
        //    Assert.AreEqual(h.Left, "\u20ac");
        //    Assert.AreEqual(h.Right, "\u20ac");

        //    //Test the footer
        //    f = s.Footer;
        //    Assert.AreEqual(f.Center, "\u20ac");
        //    Assert.AreEqual(f.Left, "\u20ac");
        //    Assert.AreEqual(f.Right, "\u20ac");

        //    //Test the dataformat
        //    r = s.GetRow(0);
        //    c = r.GetCell(1);
        //    df = wb.CreateDataFormat();
        //    Assert.AreEqual(formatStr, df.GetFormat(c.CellStyle.DataFormat));

        //    //Test the cell string value
        //    c2 = r.GetCell(2);
        //    Assert.AreEqual(c.RichStringCellValue.String, "\u20ac");

        //    //Test the cell formula
        //    c3 = r.GetCell(3);
        //    Assert.AreEqual(c3.CellFormula, formulaString);
        //}

        /** Tests Bug38230
         *  That a Umlat is written  and then read back.
         *  It should have been written as a compressed unicode.
         * 
         * 
         *
         */
        [Test]
        public void TestUmlatReadWrite()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            //Create a unicode sheet name (euro symbol)
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet("Test");

            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(1);
            c.SetCellValue(new HSSFRichTextString("\u00e4"));

            //Confirm that the sring will be compressed
            Assert.AreEqual(((HSSFRichTextString)c.RichStringCellValue).UnicodeString.OptionFlags, 0);

            string path = NPOI.Util.TempFile.GetTempFilePath("umlat", "Test.xls");
            FileStream tempFile = File.Create(path);

            wb.Write(tempFile);

            tempFile.Close();
            wb = null;
 
            FileStream in1 = File.Open(path,FileMode.Open);
            wb = new HSSFWorkbook(in1);
            in1.Close();
            //Test the sheetname
            s = wb.GetSheet("Test");
            Assert.IsNotNull(s);

            c = r.GetCell(1);
            Assert.AreEqual(c.RichStringCellValue.String, "\u00e4");
        }

    }
}