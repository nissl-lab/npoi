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
namespace TestCases.XSSF.UserModel
{
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using SixLabors.ImageSharp;
    using SixLabors.ImageSharp.PixelFormats;
    using System.Collections.Generic;

    [TestFixture]
    public class TestXSSFTextParagraph
    {
        [Test]
        public void XSSFTextParagraph_()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
                XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

                XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4)) as XSSFTextBox;
                XSSFRichTextString rt = new XSSFRichTextString("Test String");

                XSSFFont font = wb.CreateFont() as XSSFFont;
                Rgb24 color =new Rgb24(0, 255, 255);
                font.SetColor(new XSSFColor(color));
                font.FontName = "Arial";
                rt.ApplyFont(font);

                shape.SetText(rt);

                List<XSSFTextParagraph> paras = shape.TextParagraphs;
                ClassicAssert.AreEqual(1, paras.Count);

                XSSFTextParagraph text = paras[(0)];
                ClassicAssert.AreEqual("Test String", text.Text);

                ClassicAssert.IsFalse(text.IsBullet);
                ClassicAssert.IsNotNull(text.GetXmlObject());
                ClassicAssert.AreEqual(shape.GetCTShape(), text.ParentShape);
                ClassicAssert.IsNotNull(text.GetEnumerator());
                ClassicAssert.IsNotNull(text.AddLineBreak());

                ClassicAssert.IsNotNull(text.TextRuns);
                ClassicAssert.AreEqual(2, text.TextRuns.Count);
                text.AddNewTextRun();
                ClassicAssert.AreEqual(3, text.TextRuns.Count);

                ClassicAssert.AreEqual(TextAlign.LEFT, text.TextAlign);
                text.TextAlign = TextAlign.None;
                ClassicAssert.AreEqual(TextAlign.LEFT, text.TextAlign);
                text.TextAlign = (TextAlign.CENTER);
                ClassicAssert.AreEqual(TextAlign.CENTER, text.TextAlign);
                text.TextAlign = (TextAlign.RIGHT);
                ClassicAssert.AreEqual(TextAlign.RIGHT, text.TextAlign);
                text.TextAlign = TextAlign.None;
                ClassicAssert.AreEqual(TextAlign.LEFT, text.TextAlign);

                text.TextFontAlign = (TextFontAlign.BASELINE);
                ClassicAssert.AreEqual(TextFontAlign.BASELINE, text.TextFontAlign);
                text.TextFontAlign = (TextFontAlign.BOTTOM);
                ClassicAssert.AreEqual(TextFontAlign.BOTTOM, text.TextFontAlign);
                text.TextFontAlign = TextFontAlign.None;
                ClassicAssert.AreEqual(TextFontAlign.BASELINE, text.TextFontAlign);
                text.TextFontAlign = TextFontAlign.None;
                ClassicAssert.AreEqual(TextFontAlign.BASELINE, text.TextFontAlign);

                ClassicAssert.IsNull(text.BulletFont);
                text.BulletFont = ("Arial");
                ClassicAssert.AreEqual("Arial", text.BulletFont);

                ClassicAssert.IsNull(text.BulletCharacter);
                text.BulletCharacter = (".");
                ClassicAssert.AreEqual(".", text.BulletCharacter);

                //ClassicAssert.IsNull(text.BulletFontColor);
                ClassicAssert.AreEqual(POIUtils.Color_Empty, text.BulletFontColor);
                text.BulletFontColor = (color);
                ClassicAssert.AreEqual(color, text.BulletFontColor);

                ClassicAssert.AreEqual(100.0, text.BulletFontSize, 0.01);
                text.BulletFontSize = (1.0);
                ClassicAssert.AreEqual(1.0, text.BulletFontSize, 0.01);
                text.BulletFontSize = (1.0);
                ClassicAssert.AreEqual(1.0, text.BulletFontSize, 0.01);
                text.BulletFontSize = (-9.0);
                ClassicAssert.AreEqual(-9.0, text.BulletFontSize, 0.01);
                text.BulletFontSize = (-9.0);
                ClassicAssert.AreEqual(-9.0, text.BulletFontSize, 0.01);
                text.BulletFontSize = (1.0);
                ClassicAssert.AreEqual(1.0, text.BulletFontSize, 0.01);
                text.BulletFontSize = (-9.0);
                ClassicAssert.AreEqual(-9.0, text.BulletFontSize, 0.01);

                ClassicAssert.AreEqual(0.0, text.Indent, 0.01);
                text.Indent = (2.0);
                ClassicAssert.AreEqual(2.0, text.Indent, 0.01);
                text.Indent = (-1.0);
                ClassicAssert.AreEqual(0.0, text.Indent, 0.01);
                text.Indent = (-1.0);
                ClassicAssert.AreEqual(0.0, text.Indent, 0.01);

                ClassicAssert.AreEqual(0.0, text.LeftMargin, 0.01);
                text.LeftMargin = (3.0);
                ClassicAssert.AreEqual(3.0, text.LeftMargin, 0.01);
                text.LeftMargin = (-1.0);
                ClassicAssert.AreEqual(0.0, text.LeftMargin, 0.01);
                text.LeftMargin = (-1.0);
                ClassicAssert.AreEqual(0.0, text.LeftMargin, 0.01);

                ClassicAssert.AreEqual(0.0, text.RightMargin, 0.01);
                text.RightMargin = (4.5);
                ClassicAssert.AreEqual(4.5, text.RightMargin, 0.01);
                text.RightMargin = (-1.0);
                ClassicAssert.AreEqual(0.0, text.RightMargin, 0.01);
                text.RightMargin = (-1.0);
                ClassicAssert.AreEqual(0.0, text.RightMargin, 0.01);

                ClassicAssert.AreEqual(0.0, text.DefaultTabSize, 0.01);

                ClassicAssert.AreEqual(0.0, text.GetTabStop(0), 0.01);
                text.AddTabStop(3.14);
                ClassicAssert.AreEqual(3.14, text.GetTabStop(0), 0.01);

                ClassicAssert.AreEqual(100.0, text.LineSpacing, 0.01);
                text.LineSpacing = (3.15);
                ClassicAssert.AreEqual(3.15, text.LineSpacing, 0.01);
                text.LineSpacing = (-2.13);
                ClassicAssert.AreEqual(-2.13, text.LineSpacing, 0.01);

                ClassicAssert.AreEqual(0.0, text.SpaceBefore, 0.01);
                text.SpaceBefore = (3.17);
                ClassicAssert.AreEqual(3.17, text.SpaceBefore, 0.01);
                text.SpaceBefore = (-4.7);
                ClassicAssert.AreEqual(-4.7, text.SpaceBefore, 0.01);

                ClassicAssert.AreEqual(0.0, text.SpaceAfter, 0.01);
                text.SpaceAfter = (6.17);
                ClassicAssert.AreEqual(6.17, text.SpaceAfter, 0.01);
                text.SpaceAfter = (-8.17);
                ClassicAssert.AreEqual(-8.17, text.SpaceAfter, 0.01);

                ClassicAssert.AreEqual(0, text.Level);
                text.Level = (1);
                ClassicAssert.AreEqual(1, text.Level);
                text.Level = (4);
                ClassicAssert.AreEqual(4, text.Level);

                ClassicAssert.IsTrue(text.IsBullet);
                ClassicAssert.IsFalse(text.IsBulletAutoNumber);
                text.IsBullet = (false);
                text.IsBullet = (false);
                ClassicAssert.IsFalse(text.IsBullet);
                ClassicAssert.IsFalse(text.IsBulletAutoNumber);
                text.IsBullet = (true);
                ClassicAssert.IsTrue(text.IsBullet);
                ClassicAssert.IsFalse(text.IsBulletAutoNumber);
                ClassicAssert.AreEqual(0, text.BulletAutoNumberStart);
                ClassicAssert.AreEqual(ListAutoNumber.ARABIC_PLAIN, text.BulletAutoNumberScheme);

                text.IsBullet = (false);
                ClassicAssert.IsFalse(text.IsBullet);
                text.SetBullet(ListAutoNumber.CIRCLE_NUM_DB_PLAIN);
                ClassicAssert.IsTrue(text.IsBullet);
                ClassicAssert.IsTrue(text.IsBulletAutoNumber);
                
                //ClassicAssert.AreEqual(0, text.BulletAutoNumberStart);
                //This value should be 1, see CT_TextAutonumberBullet.startAt, default value is 1;
                ClassicAssert.AreEqual(1, text.BulletAutoNumberStart);


                ClassicAssert.AreEqual(ListAutoNumber.CIRCLE_NUM_DB_PLAIN, text.BulletAutoNumberScheme);
                text.IsBullet = (false);
                ClassicAssert.IsFalse(text.IsBullet);
                ClassicAssert.IsFalse(text.IsBulletAutoNumber);
                text.SetBullet(ListAutoNumber.CIRCLE_NUM_WD_BLACK_PLAIN, 10);
                ClassicAssert.IsTrue(text.IsBullet);
                ClassicAssert.IsTrue(text.IsBulletAutoNumber);
                ClassicAssert.AreEqual(10, text.BulletAutoNumberStart);
                ClassicAssert.AreEqual(ListAutoNumber.CIRCLE_NUM_WD_BLACK_PLAIN, text.BulletAutoNumberScheme);


                ClassicAssert.IsNotNull(text.ToString());

                new XSSFTextParagraph(text.GetXmlObject(), shape.GetCTShape());
            }
            finally
            {
                wb.Close();
            }
        }
    }

}