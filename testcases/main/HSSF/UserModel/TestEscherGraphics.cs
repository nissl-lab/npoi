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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;


    using NUnit.Framework;using NUnit.Framework.Legacy;
    using SkiaSharp;
    using SixLabors.ImageSharp;
    using SixLabors.ImageSharp.PixelFormats;


    /**
     * Tests the capabilities of the EscherGraphics class.
     * 
     * All Tests have two escher groups available to them,
     *  one anchored at 0,0,1022,255 and another anchored
     *  at 20,30,500,200 
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestEscherGraphics
    {
        private HSSFWorkbook workbook;
        private HSSFPatriarch patriarch;
        private HSSFShapeGroup escherGroupA;
        private HSSFShapeGroup escherGroupB;
        private EscherGraphics graphics;

        // Default DPI used by EscherGraphics and SheetUtil
        private const float dpi = 96f;

        [SetUp]
        public void SetUp()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            workbook = new HSSFWorkbook();

            ISheet sheet = workbook.CreateSheet("Test");
            patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            escherGroupA = patriarch.CreateGroup(new HSSFClientAnchor(0, 0, 1022, 255, (short)0, 0, (short)0, 0));
            escherGroupB = patriarch.CreateGroup(new HSSFClientAnchor(20, 30, 500, 200, (short)0, 0, (short)0, 0));
            //        escherGroup = new HSSFShapeGroup(null, new HSSFChildAnchor());
            graphics = new EscherGraphics(this.escherGroupA, workbook, Color.Black, 1.0f);

        }

        [Test]
        public void TestGetFont()
        {
            SKFont f = graphics.Font;
            // Font was created at 10pt; stored as pixels at 96 DPI
            float expectedSizePx = 10f * dpi / 72f;
            ClassicAssert.AreEqual(expectedSizePx, f.Size, 0.01f);
            ClassicAssert.IsFalse(f.Typeface.IsBold);
            ClassicAssert.IsFalse(f.Typeface.IsItalic);
        }

        [Test]
        public void TestGetFontMetrics()
        {
            SKFont f = graphics.Font;
            using var paint = new SKPaint { Typeface = f.Typeface, TextSize = f.Size };
            float width = paint.MeasureText("X");
            ClassicAssert.IsTrue(width > 0, "Expected font metrics width > 0, but got: " + width);
            ClassicAssert.IsTrue(f.Size > 0);
            ClassicAssert.IsFalse(f.Typeface.IsBold);
            ClassicAssert.IsFalse(f.Typeface.IsItalic);
        }

        [Test]
        public void TestSetFont()
        {
            var fi = POIDataSamples.GetSpreadSheetInstance().GetFileInfo("Helvetica.ttf");
            var typeface = SKTypeface.FromFile(fi.FullName);
            var f = new SKFont(typeface, 12f * dpi / 72f);
            graphics.SetFont(f);
            ClassicAssert.AreEqual(f, graphics.Font);
        }

        [Test]
        public void TestSetColor()
        {
            graphics.SetColor(Color.Red);
            ClassicAssert.AreEqual(Color.Red, graphics.Color);
        }

        [Test]
        public void TestFillRect()
        {
            graphics.FillRect(10, 10, 20, 20);
            HSSFSimpleShape s = (HSSFSimpleShape)escherGroupA.Children[0];
            ClassicAssert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE, s.ShapeType);
            ClassicAssert.AreEqual(10, s.Anchor.Dx1);
            ClassicAssert.AreEqual(10, s.Anchor.Dy1);
            ClassicAssert.AreEqual(30, s.Anchor.Dy2);
            ClassicAssert.AreEqual(30, s.Anchor.Dx2);
        }

        [Test]
        public void TestDrawString()
        {
            graphics.DrawString("This is a Test", 10, 10);
            HSSFTextbox t = (HSSFTextbox)escherGroupA.Children[0];
            ClassicAssert.AreEqual("This is a Test", t.String.String);
        }

        [Test]
        public void TestGetDataBackAgain()
        {
            HSSFSheet s;
            HSSFShapeGroup s1;
            HSSFShapeGroup s2;

            patriarch.SetCoordinates(10, 20, 30, 40);

            MemoryStream baos = new MemoryStream();
            workbook.Write(baos);
            workbook = new HSSFWorkbook(new MemoryStream(baos.ToArray()));
            s = (HSSFSheet)workbook.GetSheetAt(0);

            patriarch = (HSSFPatriarch)s.DrawingPatriarch;

            ClassicAssert.IsNotNull(patriarch);
            ClassicAssert.AreEqual(10, patriarch.X1);
            ClassicAssert.AreEqual(20, patriarch.Y1);
            ClassicAssert.AreEqual(30, patriarch.X2);
            ClassicAssert.AreEqual(40, patriarch.Y2);

            // Check the two groups too
            ClassicAssert.AreEqual(2, patriarch.CountOfAllChildren);
            ClassicAssert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            ClassicAssert.IsTrue(patriarch.Children[1] is HSSFShapeGroup);

            s1 = (HSSFShapeGroup)patriarch.Children[0];
            s2 = (HSSFShapeGroup)patriarch.Children[1];

            ClassicAssert.AreEqual(0, s1.X1);
            ClassicAssert.AreEqual(0, s1.Y1);
            ClassicAssert.AreEqual(1023, s1.X2);
            ClassicAssert.AreEqual(255, s1.Y2);
            ClassicAssert.AreEqual(0, s2.X1);
            ClassicAssert.AreEqual(0, s2.Y1);
            ClassicAssert.AreEqual(1023, s2.X2);
            ClassicAssert.AreEqual(255, s2.Y2);

            ClassicAssert.AreEqual(0, s1.Anchor.Dx1);
            ClassicAssert.AreEqual(0, s1.Anchor.Dy1);
            ClassicAssert.AreEqual(1022, s1.Anchor.Dx2);
            ClassicAssert.AreEqual(255, s1.Anchor.Dy2);
            ClassicAssert.AreEqual(20, s2.Anchor.Dx1);
            ClassicAssert.AreEqual(30, s2.Anchor.Dy1);
            ClassicAssert.AreEqual(500, s2.Anchor.Dx2);
            ClassicAssert.AreEqual(200, s2.Anchor.Dy2);


            // Write and re-load once more, to Check that's ok
            baos = new MemoryStream();
            workbook.Write(baos);
            workbook = new HSSFWorkbook(new MemoryStream(baos.ToArray()));
            s = (HSSFSheet)workbook.GetSheetAt(0);
            patriarch = (HSSFPatriarch)s.DrawingPatriarch;

            ClassicAssert.IsNotNull(patriarch);
            ClassicAssert.AreEqual(10, patriarch.X1);
            ClassicAssert.AreEqual(20, patriarch.Y1);
            ClassicAssert.AreEqual(30, patriarch.X2);
            ClassicAssert.AreEqual(40, patriarch.Y2);

            // Check the two groups too
            ClassicAssert.AreEqual(2, patriarch.CountOfAllChildren);
            ClassicAssert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            ClassicAssert.IsTrue(patriarch.Children[1] is HSSFShapeGroup);

            s1 = (HSSFShapeGroup)patriarch.Children[0];
            s2 = (HSSFShapeGroup)patriarch.Children[1];

            ClassicAssert.AreEqual(0, s1.X1);
            ClassicAssert.AreEqual(0, s1.Y1);
            ClassicAssert.AreEqual(1023, s1.X2);
            ClassicAssert.AreEqual(255, s1.Y2);
            ClassicAssert.AreEqual(0, s2.X1);
            ClassicAssert.AreEqual(0, s2.Y1);
            ClassicAssert.AreEqual(1023, s2.X2);
            ClassicAssert.AreEqual(255, s2.Y2);

            ClassicAssert.AreEqual(0, s1.Anchor.Dx1);
            ClassicAssert.AreEqual(0, s1.Anchor.Dy1);
            ClassicAssert.AreEqual(1022, s1.Anchor.Dx2);
            ClassicAssert.AreEqual(255, s1.Anchor.Dy2);
            ClassicAssert.AreEqual(20, s2.Anchor.Dx1);
            ClassicAssert.AreEqual(30, s2.Anchor.Dy1);
            ClassicAssert.AreEqual(500, s2.Anchor.Dx2);
            ClassicAssert.AreEqual(200, s2.Anchor.Dy2);

            // Change the positions of the first groups,
            //  but not of their anchors
            s1.SetCoordinates(2, 3, 1021, 242);

            baos = new MemoryStream();
            workbook.Write(baos);
            workbook = new HSSFWorkbook(new MemoryStream(baos.ToArray()));
            s =(HSSFSheet)workbook.GetSheetAt(0);
            patriarch = (HSSFPatriarch)s.DrawingPatriarch;

            ClassicAssert.IsNotNull(patriarch);
            ClassicAssert.AreEqual(10, patriarch.X1);
            ClassicAssert.AreEqual(20, patriarch.Y1);
            ClassicAssert.AreEqual(30, patriarch.X2);
            ClassicAssert.AreEqual(40, patriarch.Y2);

            // Check the two groups too
            ClassicAssert.AreEqual(2, patriarch.CountOfAllChildren);
            ClassicAssert.AreEqual(2, patriarch.Children.Count);
            ClassicAssert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            ClassicAssert.IsTrue(patriarch.Children[1] is HSSFShapeGroup);

            s1 = (HSSFShapeGroup)patriarch.Children[0];
            s2 = (HSSFShapeGroup)patriarch.Children[1];

            ClassicAssert.AreEqual(2, s1.X1);
            ClassicAssert.AreEqual(3, s1.Y1);
            ClassicAssert.AreEqual(1021, s1.X2);
            ClassicAssert.AreEqual(242, s1.Y2);
            ClassicAssert.AreEqual(0, s2.X1);
            ClassicAssert.AreEqual(0, s2.Y1);
            ClassicAssert.AreEqual(1023, s2.X2);
            ClassicAssert.AreEqual(255, s2.Y2);

            ClassicAssert.AreEqual(0, s1.Anchor.Dx1);
            ClassicAssert.AreEqual(0, s1.Anchor.Dy1);
            ClassicAssert.AreEqual(1022, s1.Anchor.Dx2);
            ClassicAssert.AreEqual(255, s1.Anchor.Dy2);
            ClassicAssert.AreEqual(20, s2.Anchor.Dx1);
            ClassicAssert.AreEqual(30, s2.Anchor.Dy1);
            ClassicAssert.AreEqual(500, s2.Anchor.Dx2);
            ClassicAssert.AreEqual(200, s2.Anchor.Dy2);


            // Now Add some text to one group, and some more
            //  to the base, and Check we can get it back again
            HSSFTextbox tbox1 =
                patriarch.CreateTextbox(new HSSFClientAnchor(1, 2, 3, 4, (short)0, 0, (short)0, 0)) as HSSFTextbox;
            tbox1.String = (new HSSFRichTextString("I am text box 1"));
            HSSFTextbox tbox2 =
                s2.CreateTextbox(new HSSFChildAnchor(41, 42, 43, 44)) as HSSFTextbox;
            tbox2.String = (new HSSFRichTextString("This is text box 2"));

            ClassicAssert.AreEqual(3, patriarch.Children.Count);


            baos = new MemoryStream();
            workbook.Write(baos);
            workbook = new HSSFWorkbook(new MemoryStream(baos.ToArray()));
            s = (HSSFSheet)workbook.GetSheetAt(0);

            patriarch = (HSSFPatriarch)s.DrawingPatriarch;

            ClassicAssert.IsNotNull(patriarch);
            ClassicAssert.AreEqual(10, patriarch.X1);
            ClassicAssert.AreEqual(20, patriarch.Y1);
            ClassicAssert.AreEqual(30, patriarch.X2);
            ClassicAssert.AreEqual(40, patriarch.Y2);

            // Check the two groups and the text
            // Result of patriarch.countOfAllChildren() makes no sense: 
            // Returns 4 for 2 empty groups + 1 TextBox.
            //ClassicAssert.AreEqual(3, patriarch.CountOfAllChildren);
            ClassicAssert.AreEqual(3, patriarch.Children.Count);

            // Should be two groups and a text
            ClassicAssert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            ClassicAssert.IsTrue(patriarch.Children[1] is HSSFShapeGroup);
            ClassicAssert.IsTrue(patriarch.Children[2] is HSSFTextbox);
    

            s1 = (HSSFShapeGroup)patriarch.Children[0];
            tbox1 = (HSSFTextbox)patriarch.Children[2];

            s2 = (HSSFShapeGroup)patriarch.Children[1];

            ClassicAssert.AreEqual(2, s1.X1);
            ClassicAssert.AreEqual(3, s1.Y1);
            ClassicAssert.AreEqual(1021, s1.X2);
            ClassicAssert.AreEqual(242, s1.Y2);
            ClassicAssert.AreEqual(0, s2.X1);
            ClassicAssert.AreEqual(0, s2.Y1);
            ClassicAssert.AreEqual(1023, s2.X2);
            ClassicAssert.AreEqual(255, s2.Y2);

            ClassicAssert.AreEqual(0, s1.Anchor.Dx1);
            ClassicAssert.AreEqual(0, s1.Anchor.Dy1);
            ClassicAssert.AreEqual(1022, s1.Anchor.Dx2);
            ClassicAssert.AreEqual(255, s1.Anchor.Dy2);
            ClassicAssert.AreEqual(20, s2.Anchor.Dx1);
            ClassicAssert.AreEqual(30, s2.Anchor.Dy1);
            ClassicAssert.AreEqual(500, s2.Anchor.Dx2);
            ClassicAssert.AreEqual(200, s2.Anchor.Dy2);

            // Not working just yet
            //ClassicAssert.AreEqual("I am text box 1", tbox1.String.String);
        }
    }
}