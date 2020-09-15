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

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Drawing;

namespace TestCases.XSSF.UserModel
{
    /**
     * Tests functionality of the XSSFRichTextRun object
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFRichTextString
    {

        [Test]
        public void TestCreate()
        {

            XSSFRichTextString rt = new XSSFRichTextString("Apache POI");
            Assert.AreEqual("Apache POI", rt.String);
            Assert.AreEqual(false, rt.HasFormatting());

            CT_Rst st = rt.GetCTRst();
            Assert.IsTrue(st.IsSetT());
            Assert.AreEqual("Apache POI", st.t);
            Assert.AreEqual(false, rt.HasFormatting());

            rt.Append(" is cool stuff");
            Assert.AreEqual(2, st.sizeOfRArray());
            Assert.IsFalse(st.IsSetT());

            Assert.AreEqual("Apache POI is cool stuff", rt.String);
            Assert.AreEqual(false, rt.HasFormatting());
        }
        [Test]
        public void TestEmpty()
        {
            XSSFRichTextString rt = new XSSFRichTextString();
            Assert.AreEqual(0, rt.GetIndexOfFormattingRun(9999));
            Assert.AreEqual(-1, rt.GetLengthOfFormattingRun(9999));
            Assert.IsNull(rt.GetFontAtIndex(9999));
        }

        [Test]
        public void TestApplyFont()
        {

            XSSFRichTextString rt = new XSSFRichTextString();
            rt.Append("123");
            rt.Append("4567");
            rt.Append("89");

            Assert.AreEqual("123456789", rt.String);
            Assert.AreEqual(false, rt.HasFormatting());

            XSSFFont font1 = new XSSFFont();
            font1.IsBold = (true);

            rt.ApplyFont(2, 5, font1);
            Assert.AreEqual(true, rt.HasFormatting());

            Assert.AreEqual(4, rt.NumFormattingRuns);
            Assert.AreEqual(0, rt.GetIndexOfFormattingRun(0));
            Assert.AreEqual("12", rt.GetCTRst().GetRArray(0).t);

            Assert.AreEqual(2, rt.GetIndexOfFormattingRun(1));
            Assert.AreEqual("345", rt.GetCTRst().GetRArray(1).t);

            Assert.AreEqual(5, rt.GetIndexOfFormattingRun(2));
            Assert.AreEqual(2, rt.GetLengthOfFormattingRun(2));
            Assert.AreEqual("67", rt.GetCTRst().GetRArray(2).t);

            Assert.AreEqual(7, rt.GetIndexOfFormattingRun(3));
            Assert.AreEqual(2, rt.GetLengthOfFormattingRun(3));
            Assert.AreEqual("89", rt.GetCTRst().GetRArray(3).t);

            Assert.AreEqual(-1, rt.GetIndexOfFormattingRun(9999));
            Assert.AreEqual(-1, rt.GetLengthOfFormattingRun(9999));
            Assert.IsNull(rt.GetFontAtIndex(9999));

        }

        [Test]
        public void TestApplyFontIndex()
        {
            XSSFRichTextString rt = new XSSFRichTextString("Apache POI");
            rt.ApplyFont(0, 10, (short)1);

            rt.ApplyFont((short)1);

            Assert.IsNotNull(rt.GetFontAtIndex(0));
        }

        [Test]
        public void TestApplyFontWithStyles()
        {
            XSSFRichTextString rt = new XSSFRichTextString("Apache POI");

            StylesTable tbl = new StylesTable();
            rt.SetStylesTableReference(tbl);

            try
            {
                rt.ApplyFont(0, 10, (short)1);
                Assert.Fail("Fails without styles in the table");
            }
            catch (ArgumentOutOfRangeException e)
            {
                // expected
            }

            tbl.PutFont(new XSSFFont());
            rt.ApplyFont(0, 10, (short)1);
            rt.ApplyFont((short)1);
        }

        [Test]
        public void TestApplyFontException()
        {
            XSSFRichTextString rt = new XSSFRichTextString("Apache POI");

            rt.ApplyFont(0, 0, (short)1);

            try
            {
                rt.ApplyFont(11, 10, (short)1);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.Contains("11"));
            }

            try
            {
                rt.ApplyFont(-1, 10, (short)1);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.Contains("-1"));
            }

            try
            {
                rt.ApplyFont(0, 555, (short)1);
                Assert.Fail("Should catch Exception here");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.Contains("555"));
            }
        }

        [Test]
        public void TestClearFormatting()
        {

            XSSFRichTextString rt = new XSSFRichTextString("Apache POI");
            Assert.AreEqual("Apache POI", rt.String);
            Assert.AreEqual(false, rt.HasFormatting());

            rt.ClearFormatting();

            CT_Rst st = rt.GetCTRst();
            Assert.IsTrue(st.IsSetT());
            Assert.AreEqual("Apache POI", rt.String);
            Assert.AreEqual(0, rt.NumFormattingRuns);
            Assert.AreEqual(false, rt.HasFormatting());

            XSSFFont font = new XSSFFont();
            font.IsBold = true;

            rt.ApplyFont(7, 10, font);
            Assert.AreEqual(2, rt.NumFormattingRuns);
            Assert.AreEqual(true, rt.HasFormatting());

            rt.ClearFormatting();

            Assert.AreEqual("Apache POI", rt.String);
            Assert.AreEqual(0, rt.NumFormattingRuns);
            Assert.AreEqual(false, rt.HasFormatting());
        }
        [Test]
        public void TestGetFonts()
        {

            XSSFRichTextString rt = new XSSFRichTextString();

            XSSFFont font1 = new XSSFFont();
            font1.FontName = ("Arial");
            font1.IsItalic = (true);
            rt.Append("The quick", font1);

            XSSFFont font1FR = (XSSFFont)rt.GetFontOfFormattingRun(0);
            Assert.AreEqual(font1.IsItalic, font1FR.IsItalic);
            Assert.AreEqual(font1.FontName, font1FR.FontName);

            XSSFFont font2 = new XSSFFont();
            font2.FontName = ("Courier");
            font2.IsBold = (true);
            rt.Append(" brown fox", font2);

            XSSFFont font2FR = (XSSFFont)rt.GetFontOfFormattingRun(1);
            Assert.AreEqual(font2.IsBold, font2FR.IsBold);
            Assert.AreEqual(font2.FontName, font2FR.FontName);
        }
        [Test]
        /**
         * make sure we insert xml:space="preserve" attribute
         * if a string has leading or trailing white spaces
         */
        public void TestPreserveSpaces()
        {
            XSSFRichTextString rt = new XSSFRichTextString("Apache");
            CT_Rst ct = rt.GetCTRst();
            string t=ct.t;
            
            Assert.AreEqual("<t>Apache</t>", ct.XmlText);
            rt.String = "  Apache";
            Assert.AreEqual("<t xml:space=\"preserve\">  Apache</t>", ct.XmlText);
            rt.Append(" POI");
            rt.Append(" ");
            Assert.AreEqual("  Apache POI ", rt.String);
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">  Apache</t>", rt.GetCTRst().GetRArray(0).xmlText());
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\"> POI</xml-fragment>", rt.getCTRst().getRArray(1).xgetT().xmlText());
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\"> </xml-fragment>", rt.getCTRst().getRArray(2).xgetT().xmlText());
        }

        /**
         * Ensure that strings with the color set as RGB are treated differently when the color is different.
         */
        [Test]
        public void TestRgbColor()
        {
            const string testText = "Apache";
            XSSFRichTextString rt = new XSSFRichTextString(testText);
            XSSFFont font = new XSSFFont { FontName = "Times New Roman", FontHeightInPoints = 11 };
            font.SetColor(new XSSFColor(Color.Red));
            rt.ApplyFont(0, testText.Length, font);
            CT_Rst ct = rt.GetCTRst();

            Assert.AreEqual("<r><rPr><color rgb=\"FF0000\"/><rFont val=\"Times New Roman\"/><sz val=\"11\"/></rPr><t>Apache</t></r>", ct.XmlText);
        }

        /**
         * Test that unicode representation_ xHHHH_ is properly Processed
         */
        [Test]
        public void TestUtfDecode()
        {
            CT_Rst st = new CT_Rst();
            st.t = ("abc_x000D_2ef_x000D_");
            XSSFRichTextString rt = new XSSFRichTextString(st);
            //_x000D_ is Converted into carriage return
            Assert.AreEqual("abc\r2ef\r", rt.String);

        }
        [Ignore("Not Implemented")]
        public void TestApplyFont_lowlevel()
        {
            CT_Rst st = new CT_Rst();
            String text = "Apache Software Foundation";
            XSSFRichTextString str = new XSSFRichTextString(text);
            Assert.AreEqual(26, text.Length);

            st.AddNewR().t = (text);

            //SortedDictionary<int, CT_RPrElt> formats = str.GetFormatMap(st);
            //Assert.AreEqual(1, formats.Count);
            //Assert.AreEqual(26, (int)formats.Keys[0]);
            //Assert.IsNull(formats.Get(formats.firstKey()));

            //CT_RPrElt fmt1 = new CT_RPrElt();
            //str.ApplyFont(formats, 0, 6, fmt1);
            //Assert.AreEqual(2, formats.Count);
            //Assert.AreEqual("[6, 26]", formats.Keys.ToString());
            //Object[] Runs1 = formats.Values.ToArray();
            //Assert.AreSame(fmt1, Runs1[0]);
            //Assert.AreSame(null, Runs1[1]);

            //CT_RPrElt fmt2 = new CT_RPrElt();
            //str.ApplyFont(formats, 7, 15, fmt2);
            //Assert.AreEqual(4, formats.Count);
            //Assert.AreEqual("[6, 7, 15, 26]", formats.Keys.ToString());
            //Object[] Runs2 = formats.Values.ToArray();
            //Assert.AreSame(fmt1, Runs2[0]);
            //Assert.AreSame(null, Runs2[1]);
            //Assert.AreSame(fmt2, Runs2[2]);
            //Assert.AreSame(null, Runs2[3]);

            //CT_RPrElt fmt3 = new CT_RPrElt();
            //str.ApplyFont(formats, 6, 7, fmt3);
            //Assert.AreEqual(4, formats.Count);
            //Assert.AreEqual("[6, 7, 15, 26]", formats.Keys.ToString());
            //Object[] Runs3 = formats.Values.ToArray();
            //Assert.AreSame(fmt1, Runs3[0]);
            //Assert.AreSame(fmt3, Runs3[1]);
            //Assert.AreSame(fmt2, Runs3[2]);
            //Assert.AreSame(null, Runs3[3]);

            //CT_RPrElt fmt4 = new CT_RPrElt();
            //str.ApplyFont(formats, 0, 7, fmt4);
            //Assert.AreEqual(3, formats.Count);
            //Assert.AreEqual("[7, 15, 26]", formats.Keys.ToString());
            //Object[] Runs4 = formats.Values.ToArray();
            //Assert.AreSame(fmt4, Runs4[0]);
            //Assert.AreSame(fmt2, Runs4[1]);
            //Assert.AreSame(null, Runs4[2]);

            //CT_RPrElt fmt5 = new CT_RPrElt();
            //str.ApplyFont(formats, 0, 26, fmt5);
            //Assert.AreEqual(1, formats.Count);
            //Assert.AreEqual("[26]", formats.Keys.ToString());
            //Object[] Runs5 = formats.Values.ToArray();
            //Assert.AreSame(fmt5, Runs5[0]);

            //CT_RPrElt fmt6 = new CT_RPrElt();
            //str.ApplyFont(formats, 15, 26, fmt6);
            //Assert.AreEqual(2, formats.Count);
            //Assert.AreEqual("[15, 26]", formats.Keys.ToString());
            //Object[] Runs6 = formats.Values.ToArray();
            //Assert.AreSame(fmt5, Runs6[0]);
            //Assert.AreSame(fmt6, Runs6[1]);

            //str.ApplyFont(formats, 0, 26, null);
            //Assert.AreEqual(1, formats.Count);
            //Assert.AreEqual("[26]", formats.Keys.ToString());
            //Object[] Runs7 = formats.Values.ToArray();
            //Assert.AreSame(null, Runs7[0]);

            //str.ApplyFont(formats, 15, 26, fmt6);
            //Assert.AreEqual(2, formats.Count);
            //Assert.AreEqual("[15, 26]", formats.Keys.ToString());
            //Object[] Runs8 = formats.Values.ToArray();
            //Assert.AreSame(null, Runs8[0]);
            //Assert.AreSame(fmt6, Runs8[1]);

            //str.ApplyFont(formats, 15, 26, fmt5);
            //Assert.AreEqual(2, formats.Count);
            //Assert.AreEqual("[15, 26]", formats.Keys.ToString());
            //Object[] Runs9 = formats.Values.ToArray();
            //Assert.AreSame(null, Runs9[0]);
            //Assert.AreSame(fmt5, Runs9[1]);

            //str.ApplyFont(formats, 2, 20, fmt6);
            //Assert.AreEqual(3, formats.Count);
            //Assert.AreEqual("[2, 20, 26]", formats.Keys.ToString());
            //Object[] Runs10 = formats.Values.ToArray();
            //Assert.AreSame(null, Runs10[0]);
            //Assert.AreSame(fmt6, Runs10[1]);
            //Assert.AreSame(fmt5, Runs10[2]);

            //str.ApplyFont(formats, 22, 24, fmt4);
            //Assert.AreEqual(5, formats.Count);
            //Assert.AreEqual("[2, 20, 22, 24, 26]", formats.Keys.ToString());
            //Object[] Runs11 = formats.Values.ToArray();
            //Assert.AreSame(null, Runs11[0]);
            //Assert.AreSame(fmt6, Runs11[1]);
            //Assert.AreSame(fmt5, Runs11[2]);
            //Assert.AreSame(fmt4, Runs11[3]);
            //Assert.AreSame(fmt5, Runs11[4]);

            //str.ApplyFont(formats, 0, 10, fmt1);
            //Assert.AreEqual(5, formats.Count);
            //Assert.AreEqual("[10, 20, 22, 24, 26]", formats.Keys.ToString());
            //Object[] Runs12 = formats.Values.ToArray();
            //Assert.AreSame(fmt1, Runs12[0]);
            //Assert.AreSame(fmt6, Runs12[1]);
            //Assert.AreSame(fmt5, Runs12[2]);
            //Assert.AreSame(fmt4, Runs12[3]);
            //Assert.AreSame(fmt5, Runs12[4]);

            Assert.Fail("implement STXString");
        }
        [Test]
        public void TestApplyFont_usermodel()
        {
            String text = "Apache Software Foundation";
            XSSFRichTextString str = new XSSFRichTextString(text);
            XSSFFont font1 = new XSSFFont();
            XSSFFont font2 = new XSSFFont();
            XSSFFont font3 = new XSSFFont();
            str.ApplyFont(font1);
            Assert.AreEqual(1, str.NumFormattingRuns);

            str.ApplyFont(0, 6, font1);
            str.ApplyFont(6, text.Length, font2);
            Assert.AreEqual(2, str.NumFormattingRuns);
            Assert.AreEqual("Apache", str.GetCTRst().GetRArray(0).t);
            Assert.AreEqual(" Software Foundation", str.GetCTRst().GetRArray(1).t);

            str.ApplyFont(15, 26, font3);
            Assert.AreEqual(3, str.NumFormattingRuns);
            Assert.AreEqual("Apache", str.GetCTRst().GetRArray(0).t);
            Assert.AreEqual(" Software", str.GetCTRst().GetRArray(1).t);
            Assert.AreEqual(" Foundation", str.GetCTRst().GetRArray(2).t);

            str.ApplyFont(6, text.Length, font2);
            Assert.AreEqual(2, str.NumFormattingRuns);
            Assert.AreEqual("Apache", str.GetCTRst().GetRArray(0).t);
            Assert.AreEqual(" Software Foundation", str.GetCTRst().GetRArray(1).t);
        }
        [Ignore("test")]
        public void TestLineBreaks_bug48877()
        {

            //XSSFFont font = new XSSFFont();
            //font.Boldweight = (short)FontBoldWeight.Bold;
            //font.FontHeightInPoints = ((short)14);
            //XSSFRichTextString str;
            //STXstring t1, t2, t3;

            //str = new XSSFRichTextString("Incorrect\nLine-Breaking");
            //str.ApplyFont(0, 8, font);
            //t1 = str.GetCTRst().r[0].xgetT();
            //t2 = str.GetCTRst().r[1].xgetT();
            //Assert.AreEqual("<xml-fragment>Incorrec</xml-fragment>", t1.xmlText());
            //Assert.AreEqual("<xml-fragment>t\nLine-Breaking</xml-fragment>", t2.xmlText());

            //str = new XSSFRichTextString("Incorrect\nLine-Breaking");
            //str.ApplyFont(0, 9, font);
            //t1 = str.GetCTRst().r[0].xgetT();
            //t2 = str.GetCTRst().r[1].xgetT();
            //Assert.AreEqual("<xml-fragment>Incorrect</xml-fragment>", t1.xmlText());
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\nLine-Breaking</xml-fragment>", t2.xmlText());

            //str = new XSSFRichTextString("Incorrect\n Line-Breaking");
            //str.ApplyFont(0, 9, font);
            //t1 = str.GetCTRst().r[0].xgetT();
            //t2 = str.GetCTRst().r[1].xgetT();
            //Assert.AreEqual("<xml-fragment>Incorrect</xml-fragment>", t1.xmlText());
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\n Line-Breaking</xml-fragment>", t2.xmlText());

            //str = new XSSFRichTextString("Tab\tSeparated\n");
            //t1 = str.GetCTRst().xgetT();
            //// trailing \n causes must be preserved
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">Tab\tSeparated\n</xml-fragment>", t1.xmlText());

            //str.ApplyFont(0, 3, font);
            //t1 = str.GetCTRst().r[0].xgetT();
            //t2 = str.GetCTRst().r[1].xgetT();
            //Assert.AreEqual("<xml-fragment>Tab</xml-fragment>", t1.xmlText());
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\tSeparated\n</xml-fragment>", t2.xmlText());

            //str = new XSSFRichTextString("Tab\tSeparated\n");
            //str.ApplyFont(0, 4, font);
            //t1 = str.GetCTRst().r[0].xgetT();
            //t2 = str.GetCTRst().r[1].xgetT();
            //// YK: don't know why, but XmlBeans Converts leading tab characters to spaces
            ////Assert.AreEqual("<xml-fragment>Tab\t</xml-fragment>", t1.xmlText());
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">Separated\n</xml-fragment>", t2.xmlText());

            //str = new XSSFRichTextString("\n\n\nNew Line\n\n");
            //str.ApplyFont(0, 3, font);
            //str.ApplyFont(11, 13, font);
            //t1 = str.GetCTRst().r[0].xgetT();
            //t2 = str.GetCTRst().r[1].xgetT();
            //t3 = str.GetCTRst().r[2].xgetT();
            //// YK: don't know why, but XmlBeans Converts leading tab characters to spaces
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\n\n\n</xml-fragment>", t1.xmlText());
            //Assert.AreEqual("<xml-fragment>New Line</xml-fragment>", t2.xmlText());
            //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\n\n</xml-fragment>", t3.xmlText());
            Assert.Fail("implement STXString");
        }

        [Test]
        public void TestBug56511()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56511.xlsx");
            foreach (XSSFSheet sheet in wb)
            {
                int lastRow = sheet.LastRowNum;
                for (int rowIdx = sheet.FirstRowNum; rowIdx <= lastRow; rowIdx++)
                {
                    XSSFRow row = sheet.GetRow(rowIdx) as XSSFRow;
                    if (row != null)
                    {
                        int lastCell = row.LastCellNum;

                        for (int cellIdx = row.FirstCellNum; cellIdx <= lastCell; cellIdx++)
                        {

                            XSSFCell cell = row.GetCell(cellIdx) as XSSFCell;
                            if (cell != null)
                            {
                                //System.out.Println("row " + rowIdx + " column " + cellIdx + ": " + cell.CellType + ": " + cell.ToString());

                                XSSFRichTextString richText = cell.RichStringCellValue as XSSFRichTextString;
                                int anzFormattingRuns = richText.NumFormattingRuns;
                                for (int run = 0; run < anzFormattingRuns; run++)
                                {
                                    /*XSSFFont font =*/
                                    richText.GetFontOfFormattingRun(run);
                                    //System.out.Println("  run " + run
                                    //        + " font " + (font == null ? "<null>" : font.FontName));
                                }
                            }
                        }
                    }
                }
            }
        }

        [Test]
        public void TestBug56511_values()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56511.xlsx");
            ISheet sheet = wb.GetSheetAt(0);
            IRow row = sheet.GetRow(0);

            // verify the values to ensure future Changes keep the returned information equal 
            XSSFRichTextString rt = (XSSFRichTextString)row.GetCell(0).RichStringCellValue;
            Assert.AreEqual(0, rt.NumFormattingRuns);
            Assert.IsNull(rt.GetFontOfFormattingRun(0));
            Assert.AreEqual(-1, rt.GetLengthOfFormattingRun(0));

            rt = (XSSFRichTextString)row.GetCell(1).RichStringCellValue;
            Assert.AreEqual(0, row.GetCell(1).RichStringCellValue.NumFormattingRuns);
            Assert.IsNull(rt.GetFontOfFormattingRun(1));
            Assert.AreEqual(-1, rt.GetLengthOfFormattingRun(1));

            rt = (XSSFRichTextString)row.GetCell(2).RichStringCellValue;
            Assert.AreEqual(2, rt.NumFormattingRuns);
            Assert.IsNotNull(rt.GetFontOfFormattingRun(0));
            Assert.AreEqual(4, rt.GetLengthOfFormattingRun(0));

            Assert.IsNotNull(rt.GetFontOfFormattingRun(1));
            Assert.AreEqual(9, rt.GetLengthOfFormattingRun(1));

            Assert.IsNull(rt.GetFontOfFormattingRun(2));

            rt = (XSSFRichTextString)row.GetCell(3).RichStringCellValue;
            Assert.AreEqual(3, rt.NumFormattingRuns);
            Assert.IsNull(rt.GetFontOfFormattingRun(0));
            Assert.AreEqual(1, rt.GetLengthOfFormattingRun(0));

            Assert.IsNotNull(rt.GetFontOfFormattingRun(1));
            Assert.AreEqual(3, rt.GetLengthOfFormattingRun(1));

            Assert.IsNotNull(rt.GetFontOfFormattingRun(2));
            Assert.AreEqual(9, rt.GetLengthOfFormattingRun(2));

        }

        [Test]
        public void TestToString()
        {
            XSSFRichTextString rt = new XSSFRichTextString("Apache POI");
            Assert.IsNotNull(rt.ToString());

            // TODO: normally ToString() should never return null, should we adjust this?
            rt = new XSSFRichTextString();
            Assert.IsNull(rt.ToString());
        }
        [Test]
        public void Test59008Font()
        {
            XSSFFont font = new XSSFFont(new CT_Font());

            XSSFRichTextString rts = new XSSFRichTextString();
            rts.Append("This is correct ");
            int s1 = rts.Length;
            rts.Append("This is Bold Red", font);
            int s2 = rts.Length;
            rts.Append(" This uses the default font rather than the cell style font");
            int s3 = rts.Length;

            //Assert.AreEqual("<xml-fragment/>", rts.GetFontAtIndex(s1 - 1).ToString());
            Assert.AreEqual("<font></font>", rts.GetFontAtIndex(s1 - 1).ToString());
            Assert.AreEqual(font, rts.GetFontAtIndex(s2 - 1));
            //Assert.AreEqual("<xml-fragment/>", rts.GetFontAtIndex(s3 - 1).ToString());
            Assert.AreEqual("<font></font>", rts.GetFontAtIndex(s3 - 1).ToString());
        }
    }
}