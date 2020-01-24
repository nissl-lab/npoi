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

using NPOI;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System.Text;
using TestCases.SS.UserModel;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFFont : BaseTestFont
    {

        public TestXSSFFont()
            : base(XSSFITestDataProvider.instance)
        {

        }
        [Test]
        public void TestDefaultFont()
        {
            BaseTestDefaultFont("Calibri", 220, IndexedColors.Black.Index);
        }
        [Test]
        public void TestConstructor()
        {
            XSSFFont xssfFont = new XSSFFont();
            Assert.IsNotNull(xssfFont.GetCTFont());
        }
        [Test]
        public void TestBold()
        {
            CT_Font ctFont = new CT_Font();
            CT_BooleanProperty bool1 = ctFont.AddNewB();
            bool1.val = (false);
            ctFont.SetBArray(0, bool1);
            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(false, xssfFont.IsBold);

            xssfFont.IsBold = (true);
            Assert.AreEqual(ctFont.b.Count, 1);
            Assert.AreEqual(true, ctFont.GetBArray(0).val);
        }
        [Test]
        public void TestCharSet()
        {
            CT_Font ctFont = new CT_Font();
            CT_IntProperty prop = ctFont.AddNewCharset();
            prop.val = (FontCharset.ANSI.Value);

            ctFont.SetCharsetArray(0, prop);
            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(FontCharset.ANSI.Value, xssfFont.Charset);

            xssfFont.SetCharSet(FontCharset.DEFAULT);
            Assert.AreEqual(FontCharset.DEFAULT.Value, ctFont.GetCharsetArray(0).val);

            // Try with a few less usual ones:
            // Set with the Charset itself
            xssfFont.SetCharSet(FontCharset.RUSSIAN);
            Assert.AreEqual(FontCharset.RUSSIAN.Value, xssfFont.Charset);
            // And Set with the Charset index
            xssfFont.SetCharSet(FontCharset.ARABIC.Value);
            Assert.AreEqual(FontCharset.ARABIC.Value, xssfFont.Charset);
            xssfFont.SetCharSet((byte)(FontCharset.ARABIC.Value));
            Assert.AreEqual(FontCharset.ARABIC.Value, xssfFont.Charset);

            // This one isn't allowed
            Assert.AreEqual(null, FontCharset.ValueOf(9999));
            try
            {
                xssfFont.SetCharSet(9999);
                Assert.Fail("Shouldn't be able to Set an invalid charset");
            }
            catch (POIXMLException) { }


            // Now try with a few sample files

            // Normal charset
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("Formatting.xlsx");
            Assert.AreEqual(0,
                  ((XSSFCellStyle)workbook.GetSheetAt(0).GetRow(0).GetCell(0).CellStyle).GetFont().Charset
            );

            // GB2312 charact Set
            workbook = XSSFTestDataSamples.OpenSampleWorkbook("49273.xlsx");
            Assert.AreEqual(134,
                  ((XSSFCellStyle)workbook.GetSheetAt(0).GetRow(0).GetCell(0).CellStyle).GetFont().Charset
            );
        }
        [Test]
        public void TestFontName()
        {
            CT_Font ctFont = new CT_Font();
            CT_FontName fname = ctFont.AddNewName();
            fname.val = "Arial";
            ctFont.name = fname;

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual("Arial", xssfFont.FontName);

            xssfFont.FontName = "Courier";
            Assert.AreEqual("Courier", ctFont.name.val);
        }
        [Test]
        public void TestItalic()
        {
            CT_Font ctFont = new CT_Font();
            CT_BooleanProperty bool1 = ctFont.AddNewI();
            bool1.val = (false);
            ctFont.SetIArray(0, bool1);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(false, xssfFont.IsItalic);

            xssfFont.IsItalic = (true);
            Assert.AreEqual(1, ctFont.i.Count);
            Assert.AreEqual(true, ctFont.GetIArray(0).val);
            Assert.AreEqual(true, ctFont.GetIArray(0).val);
        }
        [Test]
        public void TestStrikeout()
        {
            CT_Font ctFont = new CT_Font();
            CT_BooleanProperty bool1 = ctFont.AddNewStrike();
            bool1.val = (false);
            ctFont.SetStrikeArray(0, bool1);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(false, xssfFont.IsStrikeout);

            xssfFont.IsStrikeout = (true);
            Assert.AreEqual(1, ctFont.strike.Count);
            Assert.AreEqual(true, ctFont.GetStrikeArray(0).val);
            Assert.AreEqual(true, ctFont.GetStrikeArray(0).val);
        }
        [Test]
        public void TestFontHeight()
        {
            CT_Font ctFont = new CT_Font();
            CT_FontSize size = ctFont.AddNewSz();
            size.val = (11);
            ctFont.SetSzArray(0, size);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(11, xssfFont.FontHeightInPoints);

            //should use FontHeightInPoints here, not FontHeight.
            xssfFont.FontHeightInPoints = 20;
            Assert.AreEqual(20.0, ctFont.GetSzArray(0).val, 0.0);
        }
        [Test]
        public void TestFontHeightInPoint()
        {
            CT_Font ctFont = new CT_Font();
            CT_FontSize size = ctFont.AddNewSz();
            size.val = (14);
            ctFont.SetSzArray(0, size);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(14, xssfFont.FontHeightInPoints);

            xssfFont.FontHeightInPoints = (short)20;
            Assert.AreEqual(20.0, ctFont.GetSzArray(0).val, 0.0);
        }
        [Test]
        public void TestUnderline()
        {
            CT_Font ctFont = new CT_Font();
            CT_UnderlineProperty underlinePropr = ctFont.AddNewU();
            underlinePropr.val = (ST_UnderlineValues.single);
            ctFont.SetUArray(0, underlinePropr);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(FontUnderlineType.Single, xssfFont.Underline);

            xssfFont.SetUnderline(FontUnderlineType.Double);
            Assert.AreEqual(1, ctFont.u.Count);
            Assert.AreEqual(ST_UnderlineValues.@double, ctFont.GetUArray(0).val);

            xssfFont.SetUnderline(FontUnderlineType.DoubleAccounting);
            Assert.AreEqual(1, ctFont.u.Count);
            Assert.AreEqual(ST_UnderlineValues.doubleAccounting, ctFont.GetUArray(0).val);
        }
        [Test]
        public void TestColor()
        {
            CT_Font ctFont = new CT_Font();
            CT_Color color = ctFont.AddNewColor();
            color.indexed = (uint)(XSSFFont.DEFAULT_FONT_COLOR);
            ctFont.SetColorArray(0, color);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(IndexedColors.Black.Index, xssfFont.Color);

            xssfFont.Color = IndexedColors.Red.Index;
            Assert.AreEqual((uint)IndexedColors.Red.Index, ctFont.GetColorArray(0).indexed);
        }
        [Test]
        public void TestRgbColor()
        {
            CT_Font ctFont = new CT_Font();
            CT_Color color = ctFont.AddNewColor();

            //Integer.toHexString(0xFFFFFF).getBytes() = [102, 102, 102, 102, 102, 102]
            color.SetRgb(Encoding.ASCII.GetBytes("ffffff"));
            ctFont.SetColorArray(0, color);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[0], xssfFont.GetXSSFColor().RGB[0]);
            Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[1], xssfFont.GetXSSFColor().RGB[1]);
            Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[2], xssfFont.GetXSSFColor().RGB[2]);
            Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[3], xssfFont.GetXSSFColor().RGB[3]);

            xssfFont.Color = ((short)23);

            byte[] bytes = Encoding.ASCII.GetBytes(HexDump.ToHex(0xF1F1F1));
            color.rgb = (bytes);

            XSSFColor newColor = new XSSFColor(color);
            xssfFont.SetColor(newColor);
            Assert.AreEqual(ctFont.GetColorArray(0).GetRgb()[2], newColor.RGB[2]);

            CollectionAssert.AreEqual(bytes, xssfFont.GetXSSFColor().RGB);
            Assert.AreEqual(0, xssfFont.Color);
        }
        [Test]
        public void TestThemeColor()
        {
            CT_Font ctFont = new CT_Font();
            CT_Color color = ctFont.AddNewColor();
            color.theme = (1);
            color.themeSpecified = true;
            ctFont.SetColorArray(0, color);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual((short)ctFont.GetColorArray(0).theme, xssfFont.GetThemeColor());

            xssfFont.SetThemeColor(IndexedColors.Red.Index);
            Assert.AreEqual((uint)IndexedColors.Red.Index, ctFont.GetColorArray(0).theme);
        }
        [Test]
        public void TestFamily()
        {
            CT_Font ctFont = new CT_Font();
            CT_IntProperty family = ctFont.AddNewFamily();
            family.val = (FontFamily.MODERN.Value);
            ctFont.SetFamilyArray(0, family);

            XSSFFont xssfFont = new XSSFFont(ctFont);
            Assert.AreEqual(FontFamily.MODERN.Value, xssfFont.Family);
        }
        [Test]
        public void TestScheme()
        {
            CT_Font ctFont = new CT_Font();
            CT_FontScheme scheme = ctFont.AddNewScheme();
            scheme.val = (ST_FontScheme.major);
            ctFont.SetSchemeArray(0, scheme);

            XSSFFont font = new XSSFFont(ctFont);
            Assert.AreEqual(FontScheme.MAJOR, font.GetScheme());

            font.SetScheme(FontScheme.NONE);
            Assert.AreEqual(ST_FontScheme.none, ctFont.GetSchemeArray(0).val);
        }
        [Test]
        public void TestTypeOffset()
        {
            CT_Font ctFont = new CT_Font();
            CT_VerticalAlignFontProperty valign = ctFont.AddNewVertAlign();
            valign.val = (ST_VerticalAlignRun.baseline);
            ctFont.SetVertAlignArray(0, valign);

            XSSFFont font = new XSSFFont(ctFont);
            Assert.AreEqual(FontSuperScript.None, font.TypeOffset);

            font.TypeOffset = FontSuperScript.Super;
            Assert.AreEqual(ST_VerticalAlignRun.superscript, ctFont.GetVertAlignArray(0).val);
        }

        // store test from TestSheetUtil here as it uses XSSF
        [Test]
        public void TestCanComputeWidthXSSF()
        {
            IWorkbook wb = new XSSFWorkbook();

            // cannot check on result because on some machines we get back false here!
            SheetUtil.CanComputeColumnWidth(wb.GetFontAt((short)0));

            wb.Close();
        }

        // store test from TestSheetUtil here as it uses XSSF
        [Test]
        public void TestCanComputeWidthInvalidFont()
        {
            IFont font = new XSSFFont(new CT_Font());
            font.FontName = ("some non existing font name");

            // Even with invalid fonts we still get back useful data most of the time... 
            SheetUtil.CanComputeColumnWidth(font);
        }
    }
}