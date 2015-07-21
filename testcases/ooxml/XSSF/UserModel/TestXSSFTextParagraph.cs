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
namespace NPOI.XSSF.UserModel
{
    using System;

using static org.junit.Assert;





using org.junit;

    [TestFixture]
    public class TestXSSFTextParagraph
{
    [Test]
        [Test]
    public void TestXSSFTextParagraph(){
        XSSFWorkbook wb = new XSSFWorkbook();
        try {
            XSSFSheet sheet = wb.CreateSheet();
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch();
    
            XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));
            XSSFRichTextString rt = new XSSFRichTextString("Test String");
    
            XSSFFont font = wb.CreateFont();
            Color color = new Color(0, 255, 255);
            font.Color=(/*setter*/new XSSFColor(color));
            font.FontName=(/*setter*/"Arial");
            rt.ApplyFont(font);
    
            shape.Text=(/*setter*/rt);
            
            List<XSSFTextParagraph> paras = shape.TextParagraphs;
            Assert.AreEqual(1, paras.Count);
    
            XSSFTextParagraph text = paras.Get(0);
            Assert.AreEqual("Test String", text.Text);
            
            Assert.IsFalse(text.IsBullet());
            Assert.IsNotNull(text.XmlObject);
            Assert.AreEqual(shape.CTShape, text.ParentShape);
            Assert.IsNotNull(text.Iterator());
            Assert.IsNotNull(text.AddLineBreak());

            Assert.IsNotNull(text.TextRuns);
            Assert.AreEqual(2, text.TextRuns.Count);
            text.AddNewTextRun();
            Assert.AreEqual(3, text.TextRuns.Count);
            
            Assert.AreEqual(TextAlign.LEFT, text.TextAlign);
            text.TextAlign=(/*setter*/null);
            Assert.AreEqual(TextAlign.LEFT, text.TextAlign);
            text.TextAlign=(/*setter*/TextAlign.CENTER);
            Assert.AreEqual(TextAlign.CENTER, text.TextAlign);
            text.TextAlign=(/*setter*/TextAlign.RIGHT);
            Assert.AreEqual(TextAlign.RIGHT, text.TextAlign);
            text.TextAlign=(/*setter*/null);
            Assert.AreEqual(TextAlign.LEFT, text.TextAlign);
            
            text.TextFontAlign=(/*setter*/TextFontAlign.BASELINE);
            Assert.AreEqual(TextFontAlign.BASELINE, text.TextFontAlign);
            text.TextFontAlign=(/*setter*/TextFontAlign.BOTTOM);
            Assert.AreEqual(TextFontAlign.BOTTOM, text.TextFontAlign);
            text.TextFontAlign=(/*setter*/null);
            Assert.AreEqual(TextFontAlign.BASELINE, text.TextFontAlign);
            text.TextFontAlign=(/*setter*/null);
            Assert.AreEqual(TextFontAlign.BASELINE, text.TextFontAlign);
            
            Assert.IsNull(text.BulletFont);
            text.BulletFont=(/*setter*/"Arial");
            Assert.AreEqual("Arial", text.BulletFont);
            
            Assert.IsNull(text.BulletCharacter);
            text.BulletCharacter=(/*setter*/".");
            Assert.AreEqual(".", text.BulletCharacter);
            
            Assert.IsNull(text.BulletFontColor);
            text.BulletFontColor=(/*setter*/color);
            Assert.AreEqual(color, text.BulletFontColor);
            
            Assert.AreEqual(100.0, text.BulletFontSize, 0.01);
            text.BulletFontSize=(/*setter*/1.0);
            Assert.AreEqual(1.0, text.BulletFontSize, 0.01);
            text.BulletFontSize=(/*setter*/1.0);
            Assert.AreEqual(1.0, text.BulletFontSize, 0.01);
            text.BulletFontSize=(/*setter*/-9.0);
            Assert.AreEqual(-9.0, text.BulletFontSize, 0.01);
            text.BulletFontSize=(/*setter*/-9.0);
            Assert.AreEqual(-9.0, text.BulletFontSize, 0.01);
            text.BulletFontSize=(/*setter*/1.0);
            Assert.AreEqual(1.0, text.BulletFontSize, 0.01);
            text.BulletFontSize=(/*setter*/-9.0);
            Assert.AreEqual(-9.0, text.BulletFontSize, 0.01);

            Assert.AreEqual(0.0, text.Indent, 0.01);
            text.Indent=(/*setter*/2.0);
            Assert.AreEqual(2.0, text.Indent, 0.01);
            text.Indent=(/*setter*/-1.0);
            Assert.AreEqual(0.0, text.Indent, 0.01);
            text.Indent=(/*setter*/-1.0);
            Assert.AreEqual(0.0, text.Indent, 0.01);
            
            Assert.AreEqual(0.0, text.LeftMargin, 0.01);
            text.LeftMargin=(/*setter*/3.0);
            Assert.AreEqual(3.0, text.LeftMargin, 0.01);
            text.LeftMargin=(/*setter*/-1.0);
            Assert.AreEqual(0.0, text.LeftMargin, 0.01);
            text.LeftMargin=(/*setter*/-1.0);
            Assert.AreEqual(0.0, text.LeftMargin, 0.01);
            
            Assert.AreEqual(0.0, text.RightMargin, 0.01);
            text.RightMargin=(/*setter*/4.5);
            Assert.AreEqual(4.5, text.RightMargin, 0.01);
            text.RightMargin=(/*setter*/-1.0);
            Assert.AreEqual(0.0, text.RightMargin, 0.01);
            text.RightMargin=(/*setter*/-1.0);
            Assert.AreEqual(0.0, text.RightMargin, 0.01);
            
            Assert.AreEqual(0.0, text.DefaultTabSize, 0.01);
            
            Assert.AreEqual(0.0, text.GetTabStop(0), 0.01);
            text.AddTabStop(3.14);
            Assert.AreEqual(3.14, text.GetTabStop(0), 0.01);
            
            Assert.AreEqual(100.0, text.LineSpacing, 0.01);
            text.LineSpacing=(/*setter*/3.15);
            Assert.AreEqual(3.15, text.LineSpacing, 0.01);
            text.LineSpacing=(/*setter*/-2.13);
            Assert.AreEqual(-2.13, text.LineSpacing, 0.01);

            Assert.AreEqual(0.0, text.SpaceBefore, 0.01);
            text.SpaceBefore=(/*setter*/3.17);
            Assert.AreEqual(3.17, text.SpaceBefore, 0.01);
            text.SpaceBefore=(/*setter*/-4.7);
            Assert.AreEqual(-4.7, text.SpaceBefore, 0.01);

            Assert.AreEqual(0.0, text.SpaceAfter, 0.01);
            text.SpaceAfter=(/*setter*/6.17);
            Assert.AreEqual(6.17, text.SpaceAfter, 0.01);
            text.SpaceAfter=(/*setter*/-8.17);
            Assert.AreEqual(-8.17, text.SpaceAfter, 0.01);
            
            Assert.AreEqual(0, text.Level);
            text.Level=(/*setter*/1);
            Assert.AreEqual(1, text.Level);
            text.Level=(/*setter*/4);
            Assert.AreEqual(4, text.Level);
            
            Assert.IsTrue(text.IsBullet());
            Assert.IsFalse(text.IsBulletAutoNumber());
            text.Bullet=(/*setter*/false);
            text.Bullet=(/*setter*/false);
            Assert.IsFalse(text.IsBullet());
            Assert.IsFalse(text.IsBulletAutoNumber());
            text.Bullet=(/*setter*/true);
            Assert.IsTrue(text.IsBullet());
            Assert.IsFalse(text.IsBulletAutoNumber());
            Assert.AreEqual(0, text.BulletAutoNumberStart);
            Assert.AreEqual(ListAutoNumber.ARABIC_PLAIN, text.BulletAutoNumberScheme);
            
            text.Bullet=(/*setter*/false);
            Assert.IsFalse(text.IsBullet());
            text.Bullet=(/*setter*/ListAutoNumber.CIRCLE_NUM_DB_PLAIN);
            Assert.IsTrue(text.IsBullet());
            Assert.IsTrue(text.IsBulletAutoNumber());
            Assert.AreEqual(0, text.BulletAutoNumberStart);
            Assert.AreEqual(ListAutoNumber.CIRCLE_NUM_DB_PLAIN, text.BulletAutoNumberScheme);
            text.Bullet=(/*setter*/false);
            Assert.IsFalse(text.IsBullet());
            Assert.IsFalse(text.IsBulletAutoNumber());
            text.SetBullet(ListAutoNumber.CIRCLE_NUM_WD_BLACK_PLAIN, 10);
            Assert.IsTrue(text.IsBullet());
            Assert.IsTrue(text.IsBulletAutoNumber());
            Assert.AreEqual(10, text.BulletAutoNumberStart);
            Assert.AreEqual(ListAutoNumber.CIRCLE_NUM_WD_BLACK_PLAIN, text.BulletAutoNumberScheme);

            
            Assert.IsNotNull(text.ToString());
            
            new XSSFTextParagraph(text.XmlObject, shape.CTShape);
        } finally {
            wb.Close();
        }
    }
}

