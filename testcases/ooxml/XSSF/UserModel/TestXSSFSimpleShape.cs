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




using NPOI.SS.UserModel;
using org.junit;

    [TestFixture]
    public class TestXSSFSimpleShape
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

            Assert.IsNotNull(shape.CTShape);
            Assert.IsNotNull(shape.Iterator());
            Assert.IsNotNull(XSSFSimpleShape.Prototype());
            
            foreach(ListAutoNumber nr in ListAutoNumber.Values()) {
                shape.TextParagraphs.Get(0).Bullet=(/*setter*/nr);
                Assert.IsNotNull(shape.Text);
            }
            
            shape.TextParagraphs.Get(0).Bullet=(/*setter*/false);
            Assert.IsNotNull(shape.Text);

            shape.Text=(/*setter*/"testtext");
            Assert.AreEqual("testtext", shape.Text);
            
            shape.Text=(/*setter*/new XSSFRichTextString());
            Assert.AreEqual("null", shape.Text);
            
            shape.AddNewTextParagraph();
            shape.AddNewTextParagraph("test-other-text");
            shape.AddNewTextParagraph(new XSSFRichTextString("rtstring"));
            shape.AddNewTextParagraph(new XSSFRichTextString());
            Assert.AreEqual("null\n\ntest-other-text\nrtstring\nnull", shape.Text);
            
            Assert.AreEqual(TextHorizontalOverflow.OVERFLOW, shape.TextHorizontalOverflow);
            shape.TextHorizontalOverflow=(/*setter*/TextHorizontalOverflow.CLIP);
            Assert.AreEqual(TextHorizontalOverflow.CLIP, shape.TextHorizontalOverflow);
            shape.TextHorizontalOverflow=(/*setter*/TextHorizontalOverflow.OVERFLOW);
            Assert.AreEqual(TextHorizontalOverflow.OVERFLOW, shape.TextHorizontalOverflow);
            shape.TextHorizontalOverflow=(/*setter*/null);
            Assert.AreEqual(TextHorizontalOverflow.OVERFLOW, shape.TextHorizontalOverflow);
            shape.TextHorizontalOverflow=(/*setter*/null);
            Assert.AreEqual(TextHorizontalOverflow.OVERFLOW, shape.TextHorizontalOverflow);
            
            Assert.AreEqual(TextVerticalOverflow.OVERFLOW, shape.TextVerticalOverflow);
            shape.TextVerticalOverflow=(/*setter*/TextVerticalOverflow.CLIP);
            Assert.AreEqual(TextVerticalOverflow.CLIP, shape.TextVerticalOverflow);
            shape.TextVerticalOverflow=(/*setter*/TextVerticalOverflow.OVERFLOW);
            Assert.AreEqual(TextVerticalOverflow.OVERFLOW, shape.TextVerticalOverflow);
            shape.TextVerticalOverflow=(/*setter*/null);
            Assert.AreEqual(TextVerticalOverflow.OVERFLOW, shape.TextVerticalOverflow);
            shape.TextVerticalOverflow=(/*setter*/null);
            Assert.AreEqual(TextVerticalOverflow.OVERFLOW, shape.TextVerticalOverflow);

            Assert.AreEqual(VerticalAlignment.TOP, shape.VerticalAlignment);
            shape.VerticalAlignment=(/*setter*/VerticalAlignment.BOTTOM);
            Assert.AreEqual(VerticalAlignment.BOTTOM, shape.VerticalAlignment);
            shape.VerticalAlignment=(/*setter*/VerticalAlignment.TOP);
            Assert.AreEqual(VerticalAlignment.TOP, shape.VerticalAlignment);
            shape.VerticalAlignment=(/*setter*/null);
            Assert.AreEqual(VerticalAlignment.TOP, shape.VerticalAlignment);
            shape.VerticalAlignment=(/*setter*/null);
            Assert.AreEqual(VerticalAlignment.TOP, shape.VerticalAlignment);

            Assert.AreEqual(TextDirection.HORIZONTAL, shape.TextDirection);
            shape.TextDirection=(/*setter*/TextDirection.STACKED);
            Assert.AreEqual(TextDirection.STACKED, shape.TextDirection);
            shape.TextDirection=(/*setter*/TextDirection.HORIZONTAL);
            Assert.AreEqual(TextDirection.HORIZONTAL, shape.TextDirection);
            shape.TextDirection=(/*setter*/null);
            Assert.AreEqual(TextDirection.HORIZONTAL, shape.TextDirection);
            shape.TextDirection=(/*setter*/null);
            Assert.AreEqual(TextDirection.HORIZONTAL, shape.TextDirection);

            Assert.AreEqual(3.6, shape.BottomInset, 0.01);
            shape.BottomInset=(/*setter*/12.32);
            Assert.AreEqual(12.32, shape.BottomInset, 0.01);
            shape.BottomInset=(/*setter*/-1);
            Assert.AreEqual(3.6, shape.BottomInset, 0.01);
            shape.BottomInset=(/*setter*/-1);
            Assert.AreEqual(3.6, shape.BottomInset, 0.01);
            
            Assert.AreEqual(3.6, shape.LeftInset, 0.01);
            shape.LeftInset=(/*setter*/12.31);
            Assert.AreEqual(12.31, shape.LeftInset, 0.01);
            shape.LeftInset=(/*setter*/-1);
            Assert.AreEqual(3.6, shape.LeftInset, 0.01);
            shape.LeftInset=(/*setter*/-1);
            Assert.AreEqual(3.6, shape.LeftInset, 0.01);

            Assert.AreEqual(3.6, shape.RightInset, 0.01);
            shape.RightInset=(/*setter*/13.31);
            Assert.AreEqual(13.31, shape.RightInset, 0.01);
            shape.RightInset=(/*setter*/-1);
            Assert.AreEqual(3.6, shape.RightInset, 0.01);
            shape.RightInset=(/*setter*/-1);
            Assert.AreEqual(3.6, shape.RightInset, 0.01);

            Assert.AreEqual(3.6, shape.TopInset, 0.01);
            shape.TopInset=(/*setter*/23.31);
            Assert.AreEqual(23.31, shape.TopInset, 0.01);
            shape.TopInset=(/*setter*/-1);
            Assert.AreEqual(3.6, shape.TopInset, 0.01);
            shape.TopInset=(/*setter*/-1);
            Assert.AreEqual(3.6, shape.TopInset, 0.01);
            
            Assert.IsTrue(shape.WordWrap);
            shape.WordWrap=(/*setter*/false);
            Assert.IsFalse(shape.WordWrap);
            shape.WordWrap=(/*setter*/true);
            Assert.IsTrue(shape.WordWrap);

            Assert.AreEqual(TextAutofit.NORMAL, shape.TextAutofit);
            shape.TextAutofit=(/*setter*/TextAutofit.NORMAL);
            Assert.AreEqual(TextAutofit.NORMAL, shape.TextAutofit);
            shape.TextAutofit=(/*setter*/TextAutofit.SHAPE);
            Assert.AreEqual(TextAutofit.SHAPE, shape.TextAutofit);
            shape.TextAutofit=(/*setter*/TextAutofit.NONE);
            Assert.AreEqual(TextAutofit.NONE, shape.TextAutofit);
            
            Assert.AreEqual(5, shape.ShapeType);
            shape.ShapeType=(/*setter*/23);
            Assert.AreEqual(23, shape.ShapeType);

            // TODO: should this be supported?
//            shape.ShapeType=(/*setter*/-1);
//            Assert.AreEqual(-1, shape.ShapeType);
//            shape.ShapeType=(/*setter*/-1);
//            Assert.AreEqual(-1, shape.ShapeType);
            
            Assert.IsNotNull(shape.ShapeProperties);
        } finally {
            wb.Close();
        }
    }
}

