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
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Collections.Generic;
using System;
using NPOI.SS.UserModel;
namespace NPOI.XSSF.UserModel
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

            CT_Rst st = rt.GetCTRst();
            Assert.IsTrue(st.IsSetT());
            Assert.AreEqual("Apache POI", st.t);

            rt.Append(" is cool stuff");
            Assert.AreEqual(2, st.sizeOfRArray());
            Assert.IsFalse(st.IsSetT());

            Assert.AreEqual("Apache POI is cool stuff", rt.String);
        }

        [Test]
        public void TestApplyFont()
        {

            XSSFRichTextString rt = new XSSFRichTextString();
            rt.Append("123");
            rt.Append("4567");
            rt.Append("89");

            Assert.AreEqual("123456789", rt.String);

            XSSFFont font1 = new XSSFFont();
            font1.IsBold = (true);

            rt.ApplyFont(2, 5, font1);

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
        }
        [Test]
        public void TestClearFormatting()
        {

            XSSFRichTextString rt = new XSSFRichTextString("Apache POI");
            Assert.AreEqual("Apache POI", rt.String);

            rt.ClearFormatting();

            CT_Rst st = rt.GetCTRst();
            Assert.IsTrue(st.IsSetT());
            Assert.AreEqual("Apache POI", rt.String);
            Assert.AreEqual(0, rt.NumFormattingRuns);

            XSSFFont font = new XSSFFont();
            font.IsBold = true;

            rt.ApplyFont(7, 10, font);
            Assert.AreEqual(2, rt.NumFormattingRuns);
            rt.ClearFormatting();
            Assert.AreEqual("Apache POI", rt.String);
            Assert.AreEqual(0, rt.NumFormattingRuns);
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

        /**
         * make sure we insert xml:space="preserve" attribute
         * if a string has leading or trailing white spaces
         */
        //   [Test]
        //public void TestPreserveSpaces()
        //{
        //    XSSFRichTextString rt = new XSSFRichTextString("Apache");
        //    CT_Rst ct = rt.GetCTRst();
        //    STXstring xs = ct.xgetT();
        //    Assert.AreEqual("<xml-fragment>Apache</xml-fragment>", xs.xmlText());
        //    rt.String = ("  Apache");
        //    Assert.AreEqual("<xml-fragment xml:space=\"preserve\">  Apache</xml-fragment>", xs.xmlText());
        //rt.Append(" POI");
        //rt.Append(" ");
        //Assert.AreEqual("  Apache POI ", rt.getString());
        //Assert.AreEqual("<xml-fragment xml:space=\"preserve\">  Apache</xml-fragment>", rt.getCTRst().getRArray(0).xgetT().xmlText());
        //Assert.AreEqual("<xml-fragment xml:space=\"preserve\"> POI</xml-fragment>", rt.getCTRst().getRArray(1).xgetT().xmlText());
        //Assert.AreEqual("<xml-fragment xml:space=\"preserve\"> </xml-fragment>", rt.getCTRst().getRArray(2).xgetT().xmlText());
 
        //}

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
        //[Test]
        //public void TestApplyFont_lowlevel()
        //{
        //    CT_Rst st = new CT_Rst();
        //    String text = "Apache Software Foundation";
        //    XSSFRichTextString str = new XSSFRichTextString(text);
        //    Assert.AreEqual(26, text.Length);

        //    st.AddNewR().t = (text);

        //    Dictionary<int, CT_RPrElt> formats = str.GetFormatMap(st);
        //    Assert.AreEqual(1, formats.Count);
        //    Assert.AreEqual(26, (int)formats.firstKey());
        //    Assert.IsNull(formats.Get(formats.firstKey()));

        //    CT_RPrElt fmt1 = new CT_RPrElt();
        //    str.ApplyFont(formats, 0, 6, fmt1);
        //    Assert.AreEqual(2, formats.Count);
        //    Assert.AreEqual("[6, 26]", formats.Keys.ToString());
        //    Object[] Runs1 = formats.Values.ToArray();
        //    Assert.AreSame(fmt1, Runs1[0]);
        //    Assert.AreSame(null, Runs1[1]);

        //    CT_RPrElt fmt2 = new CT_RPrElt();
        //    str.ApplyFont(formats, 7, 15, fmt2);
        //    Assert.AreEqual(4, formats.Count);
        //    Assert.AreEqual("[6, 7, 15, 26]", formats.Keys.ToString());
        //    Object[] Runs2 = formats.Values.ToArray();
        //    Assert.AreSame(fmt1, Runs2[0]);
        //    Assert.AreSame(null, Runs2[1]);
        //    Assert.AreSame(fmt2, Runs2[2]);
        //    Assert.AreSame(null, Runs2[3]);

        //    CT_RPrElt fmt3 = new CT_RPrElt();
        //    str.ApplyFont(formats, 6, 7, fmt3);
        //    Assert.AreEqual(4, formats.Count);
        //    Assert.AreEqual("[6, 7, 15, 26]", formats.Keys.ToString());
        //    Object[] Runs3 = formats.Values.ToArray();
        //    Assert.AreSame(fmt1, Runs3[0]);
        //    Assert.AreSame(fmt3, Runs3[1]);
        //    Assert.AreSame(fmt2, Runs3[2]);
        //    Assert.AreSame(null, Runs3[3]);

        //    CT_RPrElt fmt4 = new CT_RPrElt();
        //    str.ApplyFont(formats, 0, 7, fmt4);
        //    Assert.AreEqual(3, formats.Count);
        //    Assert.AreEqual("[7, 15, 26]", formats.Keys.ToString());
        //    Object[] Runs4 = formats.Values.ToArray();
        //    Assert.AreSame(fmt4, Runs4[0]);
        //    Assert.AreSame(fmt2, Runs4[1]);
        //    Assert.AreSame(null, Runs4[2]);

        //    CT_RPrElt fmt5 = new CT_RPrElt();
        //    str.ApplyFont(formats, 0, 26, fmt5);
        //    Assert.AreEqual(1, formats.Count);
        //    Assert.AreEqual("[26]", formats.Keys.ToString());
        //    Object[] Runs5 = formats.Values.ToArray();
        //    Assert.AreSame(fmt5, Runs5[0]);

        //    CT_RPrElt fmt6 = new CT_RPrElt();
        //    str.ApplyFont(formats, 15, 26, fmt6);
        //    Assert.AreEqual(2, formats.Count);
        //    Assert.AreEqual("[15, 26]", formats.Keys.ToString());
        //    Object[] Runs6 = formats.Values.ToArray();
        //    Assert.AreSame(fmt5, Runs6[0]);
        //    Assert.AreSame(fmt6, Runs6[1]);

        //    str.ApplyFont(formats, 0, 26, null);
        //    Assert.AreEqual(1, formats.Count);
        //    Assert.AreEqual("[26]", formats.Keys.ToString());
        //    Object[] Runs7 = formats.Values.ToArray();
        //    Assert.AreSame(null, Runs7[0]);

        //    str.ApplyFont(formats, 15, 26, fmt6);
        //    Assert.AreEqual(2, formats.Count);
        //    Assert.AreEqual("[15, 26]", formats.Keys.ToString());
        //    Object[] Runs8 = formats.Values.ToArray();
        //    Assert.AreSame(null, Runs8[0]);
        //    Assert.AreSame(fmt6, Runs8[1]);

        //    str.ApplyFont(formats, 15, 26, fmt5);
        //    Assert.AreEqual(2, formats.Count);
        //    Assert.AreEqual("[15, 26]", formats.Keys.ToString());
        //    Object[] Runs9 = formats.Values.ToArray();
        //    Assert.AreSame(null, Runs9[0]);
        //    Assert.AreSame(fmt5, Runs9[1]);

        //    str.ApplyFont(formats, 2, 20, fmt6);
        //    Assert.AreEqual(3, formats.Count);
        //    Assert.AreEqual("[2, 20, 26]", formats.Keys.ToString());
        //    Object[] Runs10 = formats.Values.ToArray();
        //    Assert.AreSame(null, Runs10[0]);
        //    Assert.AreSame(fmt6, Runs10[1]);
        //    Assert.AreSame(fmt5, Runs10[2]);

        //    str.ApplyFont(formats, 22, 24, fmt4);
        //    Assert.AreEqual(5, formats.Count);
        //    Assert.AreEqual("[2, 20, 22, 24, 26]", formats.Keys.ToString());
        //    Object[] Runs11 = formats.Values.ToArray();
        //    Assert.AreSame(null, Runs11[0]);
        //    Assert.AreSame(fmt6, Runs11[1]);
        //    Assert.AreSame(fmt5, Runs11[2]);
        //    Assert.AreSame(fmt4, Runs11[3]);
        //    Assert.AreSame(fmt5, Runs11[4]);

        //    str.ApplyFont(formats, 0, 10, fmt1);
        //    Assert.AreEqual(5, formats.Count);
        //    Assert.AreEqual("[10, 20, 22, 24, 26]", formats.Keys.ToString());
        //    Object[] Runs12 = formats.Values.ToArray();
        //    Assert.AreSame(fmt1, Runs12[0]);
        //    Assert.AreSame(fmt6, Runs12[1]);
        //    Assert.AreSame(fmt5, Runs12[2]);
        //    Assert.AreSame(fmt4, Runs12[3]);
        //    Assert.AreSame(fmt5, Runs12[4]);
        //}
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
        //[Test]
        //public void TestLineBreaks_bug48877()
        //{

        //    XSSFFont font = new XSSFFont();
        //    font.Boldweight = (short)FontBoldWeight.BOLD;
        //    font.FontHeightInPoints = ((short)14);
        //    XSSFRichTextString str;
        //    STXstring t1, t2, t3;

        //    str = new XSSFRichTextString("Incorrect\nLine-Breaking");
        //    str.ApplyFont(0, 8, font);
        //    t1 = str.GetCTRst().r[0].xgetT();
        //    t2 = str.GetCTRst().r[1].xgetT();
        //    Assert.AreEqual("<xml-fragment>Incorrec</xml-fragment>", t1.xmlText());
        //    Assert.AreEqual("<xml-fragment>t\nLine-Breaking</xml-fragment>", t2.xmlText());

        //    str = new XSSFRichTextString("Incorrect\nLine-Breaking");
        //    str.ApplyFont(0, 9, font);
        //    t1 = str.GetCTRst().r[0].xgetT();
        //    t2 = str.GetCTRst().r[1].xgetT();
        //    Assert.AreEqual("<xml-fragment>Incorrect</xml-fragment>", t1.xmlText());
        //    Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\nLine-Breaking</xml-fragment>", t2.xmlText());

        //    str = new XSSFRichTextString("Incorrect\n Line-Breaking");
        //    str.ApplyFont(0, 9, font);
        //    t1 = str.GetCTRst().r[0].xgetT();
        //    t2 = str.GetCTRst().r[1].xgetT();
        //    Assert.AreEqual("<xml-fragment>Incorrect</xml-fragment>", t1.xmlText());
        //    Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\n Line-Breaking</xml-fragment>", t2.xmlText());

        //    str = new XSSFRichTextString("Tab\tSeparated\n");
        //    t1 = str.GetCTRst().xgetT();
        //    // trailing \n causes must be preserved
        //    Assert.AreEqual("<xml-fragment xml:space=\"preserve\">Tab\tSeparated\n</xml-fragment>", t1.xmlText());

        //    str.ApplyFont(0, 3, font);
        //    t1 = str.GetCTRst().r[0].xgetT();
        //    t2 = str.GetCTRst().r[1].xgetT();
        //    Assert.AreEqual("<xml-fragment>Tab</xml-fragment>", t1.xmlText());
        //    Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\tSeparated\n</xml-fragment>", t2.xmlText());

        //    str = new XSSFRichTextString("Tab\tSeparated\n");
        //    str.ApplyFont(0, 4, font);
        //    t1 = str.GetCTRst().r[0].xgetT();
        //    t2 = str.GetCTRst().r[1].xgetT();
        //    // YK: don't know why, but XmlBeans Converts leading tab characters to spaces
        //    //Assert.AreEqual("<xml-fragment>Tab\t</xml-fragment>", t1.xmlText());
        //    Assert.AreEqual("<xml-fragment xml:space=\"preserve\">Separated\n</xml-fragment>", t2.xmlText());

        //    str = new XSSFRichTextString("\n\n\nNew Line\n\n");
        //    str.ApplyFont(0, 3, font);
        //    str.ApplyFont(11, 13, font);
        //    t1 = str.GetCTRst().r[0].xgetT();
        //    t2 = str.GetCTRst().r[1].xgetT();
        //    t3 = str.GetCTRst().r[2].xgetT();
        //    // YK: don't know why, but XmlBeans Converts leading tab characters to spaces
        //    Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\n\n\n</xml-fragment>", t1.xmlText());
        //    Assert.AreEqual("<xml-fragment>New Line</xml-fragment>", t2.xmlText());
        //    Assert.AreEqual("<xml-fragment xml:space=\"preserve\">\n\n</xml-fragment>", t3.xmlText());
        //}
        }
}