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
    using System;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NPOI.XSSF.UserModel;
    using SixLabors.ImageSharp;

    [TestFixture]
    public class TestXSSFSimpleShape
    {
        [Test]
        public void TestXSSFTextParagraph()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
                XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

                XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4)) as XSSFTextBox;

                XSSFRichTextString rt = new XSSFRichTextString("Test String");

                XSSFFont font = wb.CreateFont() as XSSFFont;
                Color color = Color.FromRgb(0, 255, 255);
                font.SetColor(new XSSFColor(color));
                font.FontName = (/*setter*/"Arial");
                rt.ApplyFont(font);

                shape.SetText(/*setter*/rt);

                Assert.IsNotNull(shape.GetCTShape());
                Assert.IsNotNull(shape.GetEnumerator());
                Assert.IsNotNull(XSSFSimpleShape.Prototype());

                foreach (ListAutoNumber nr in Enum.GetValues(typeof(ListAutoNumber)))
                {
                    shape.TextParagraphs[(0)].SetBullet(nr);
                    Assert.IsNotNull(shape.Text);
                }

                shape.TextParagraphs[(0)].IsBullet = (false);
                Assert.IsNotNull(shape.Text);

                shape.SetText("testtext");
                Assert.AreEqual("testtext", shape.Text);

                shape.SetText(new XSSFRichTextString());
                //Assert.AreEqual("null", shape.Text);
                Assert.AreEqual(String.Empty, shape.Text);

                shape.AddNewTextParagraph();
                shape.AddNewTextParagraph("test-other-text");
                shape.AddNewTextParagraph(new XSSFRichTextString("rtstring"));
                shape.AddNewTextParagraph(new XSSFRichTextString());
                //Assert.AreEqual("null\n\ntest-other-text\nrtstring\nnull", shape.Text);
                Assert.AreEqual("test-other-text\nrtstring\n", shape.Text);

                Assert.AreEqual(TextHorizontalOverflow.OVERFLOW, shape.TextHorizontalOverflow);
                shape.TextHorizontalOverflow = (/*setter*/TextHorizontalOverflow.CLIP);
                Assert.AreEqual(TextHorizontalOverflow.CLIP, shape.TextHorizontalOverflow);
                shape.TextHorizontalOverflow = (/*setter*/TextHorizontalOverflow.OVERFLOW);
                Assert.AreEqual(TextHorizontalOverflow.OVERFLOW, shape.TextHorizontalOverflow);
                shape.TextHorizontalOverflow = TextHorizontalOverflow.None;
                Assert.AreEqual(TextHorizontalOverflow.OVERFLOW, shape.TextHorizontalOverflow);
                shape.TextHorizontalOverflow = TextHorizontalOverflow.None;
                Assert.AreEqual(TextHorizontalOverflow.OVERFLOW, shape.TextHorizontalOverflow);

                Assert.AreEqual(TextVerticalOverflow.OVERFLOW, shape.TextVerticalOverflow);
                shape.TextVerticalOverflow = (/*setter*/TextVerticalOverflow.CLIP);
                Assert.AreEqual(TextVerticalOverflow.CLIP, shape.TextVerticalOverflow);
                shape.TextVerticalOverflow = (/*setter*/TextVerticalOverflow.OVERFLOW);
                Assert.AreEqual(TextVerticalOverflow.OVERFLOW, shape.TextVerticalOverflow);
                shape.TextVerticalOverflow = TextVerticalOverflow.None;
                Assert.AreEqual(TextVerticalOverflow.OVERFLOW, shape.TextVerticalOverflow);
                shape.TextVerticalOverflow = TextVerticalOverflow.None;
                Assert.AreEqual(TextVerticalOverflow.OVERFLOW, shape.TextVerticalOverflow);

                Assert.AreEqual(VerticalAlignment.Top, shape.VerticalAlignment);
                shape.VerticalAlignment = VerticalAlignment.Bottom;
                Assert.AreEqual(VerticalAlignment.Bottom, shape.VerticalAlignment);
                shape.VerticalAlignment = VerticalAlignment.Top;
                Assert.AreEqual(VerticalAlignment.Top, shape.VerticalAlignment);
                shape.VerticalAlignment = VerticalAlignment.None;
                Assert.AreEqual(VerticalAlignment.Top, shape.VerticalAlignment);
                shape.VerticalAlignment = VerticalAlignment.None;
                Assert.AreEqual(VerticalAlignment.Top, shape.VerticalAlignment);

                Assert.AreEqual(TextDirection.HORIZONTAL, shape.TextDirection);
                shape.TextDirection = (/*setter*/TextDirection.STACKED);
                Assert.AreEqual(TextDirection.STACKED, shape.TextDirection);
                shape.TextDirection = (/*setter*/TextDirection.HORIZONTAL);
                Assert.AreEqual(TextDirection.HORIZONTAL, shape.TextDirection);
                shape.TextDirection = (/*setter*/TextDirection.None);
                Assert.AreEqual(TextDirection.HORIZONTAL, shape.TextDirection);
                shape.TextDirection = (/*setter*/TextDirection.None);
                Assert.AreEqual(TextDirection.HORIZONTAL, shape.TextDirection);

                Assert.AreEqual(3.6, shape.BottomInset, 0.01);
                shape.BottomInset = (/*setter*/12.32);
                Assert.AreEqual(12.32, shape.BottomInset, 0.01);
                shape.BottomInset = (/*setter*/-1);
                Assert.AreEqual(3.6, shape.BottomInset, 0.01);
                shape.BottomInset = (/*setter*/-1);
                Assert.AreEqual(3.6, shape.BottomInset, 0.01);

                Assert.AreEqual(3.6, shape.LeftInset, 0.01);
                shape.LeftInset = (/*setter*/12.31);
                Assert.AreEqual(12.31, shape.LeftInset, 0.01);
                shape.LeftInset = (/*setter*/-1);
                Assert.AreEqual(3.6, shape.LeftInset, 0.01);
                shape.LeftInset = (/*setter*/-1);
                Assert.AreEqual(3.6, shape.LeftInset, 0.01);

                Assert.AreEqual(3.6, shape.RightInset, 0.01);
                shape.RightInset = (/*setter*/13.31);
                Assert.AreEqual(13.31, shape.RightInset, 0.01);
                shape.RightInset = (/*setter*/-1);
                Assert.AreEqual(3.6, shape.RightInset, 0.01);
                shape.RightInset = (/*setter*/-1);
                Assert.AreEqual(3.6, shape.RightInset, 0.01);

                Assert.AreEqual(3.6, shape.TopInset, 0.01);
                shape.TopInset = (/*setter*/23.31);
                Assert.AreEqual(23.31, shape.TopInset, 0.01);
                shape.TopInset = (/*setter*/-1);
                Assert.AreEqual(3.6, shape.TopInset, 0.01);
                shape.TopInset = (/*setter*/-1);
                Assert.AreEqual(3.6, shape.TopInset, 0.01);

                Assert.IsTrue(shape.WordWrap);
                shape.WordWrap = (/*setter*/false);
                Assert.IsFalse(shape.WordWrap);
                shape.WordWrap = (/*setter*/true);
                Assert.IsTrue(shape.WordWrap);

                Assert.AreEqual(TextAutofit.NORMAL, shape.TextAutofit);
                shape.TextAutofit = (/*setter*/TextAutofit.NORMAL);
                Assert.AreEqual(TextAutofit.NORMAL, shape.TextAutofit);
                shape.TextAutofit = (/*setter*/TextAutofit.SHAPE);
                Assert.AreEqual(TextAutofit.SHAPE, shape.TextAutofit);
                shape.TextAutofit = (/*setter*/TextAutofit.NONE);
                Assert.AreEqual(TextAutofit.NONE, shape.TextAutofit);

                Assert.AreEqual(5, shape.ShapeType);
                shape.ShapeType = (/*setter*/23);
                Assert.AreEqual(23, shape.ShapeType);

                // TODO: should this be supported?
                //            shape.ShapeType=(/*setter*/-1);
                //            Assert.AreEqual(-1, shape.ShapeType);
                //            shape.ShapeType=(/*setter*/-1);
                //            Assert.AreEqual(-1, shape.ShapeType);

                Assert.IsNotNull(shape.GetShapeProperties());
            }
            finally
            {
                wb.Close();
            }
        }
    }

}