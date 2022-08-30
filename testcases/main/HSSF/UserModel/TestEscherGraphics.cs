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


    using NUnit.Framework;
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

        [SetUp]
        public void SetUp()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            workbook = new HSSFWorkbook();

            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet("Test");
            patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            escherGroupA = patriarch.CreateGroup(new HSSFClientAnchor(0, 0, 1022, 255, (short)0, 0, (short)0, 0));
            escherGroupB = patriarch.CreateGroup(new HSSFClientAnchor(20, 30, 500, 200, (short)0, 0, (short)0, 0));
            //        escherGroup = new HSSFShapeGroup(null, new HSSFChildAnchor());
            graphics = new EscherGraphics(this.escherGroupA, workbook, Color.Black, 1.0f);

        }

        /* TODO-Fonts:
        [Test]
        public void TestGetFont()
        {
            System.Drawing.Font f = graphics.Font;
            if (f.ToString().IndexOf("dialog") == -1 && f.ToString().IndexOf("Dialog") == -1)
            {
                //Assert.AreEqual("java.awt.Font[family=Arial,name=Arial,style=plain,size=10]", f.ToString());
                Assert.AreEqual("[Font: Name=Arial, Size=10, Units=3, GdiCharSet=1, GdiVerticalFont=False]", f.ToString());
            }
        }
        */

        //[Test]
        //public void TestGetFontMetrics()
        //{
        //    Font f = graphics.Font;
        //    if (f.ToString().IndexOf("dialog") != -1 || f.ToString().IndexOf("Dialog") != -1)
        //        return;

        //    Assert.AreEqual(7, TextRenderer.MeasureText("X", f).Width);
        //    Assert.AreEqual("Arial", f.FontFamily.Name);
        //    Assert.AreEqual(10, f.Size);
        //    //Assert.AreEqual("java.awt.Font[family=Arial,name=Arial,style=plain,size=10]", fontMetrics.GetFont().ToString());
        //}

        /* TODO-Fonts:
        [Test]
        public void TestSetFont()
        {
            System.Drawing.Font f = new System.Drawing.Font("Helvetica", 12,FontStyle.Regular);
            graphics.SetFont(f);
            Assert.AreEqual(f, graphics.Font);
        }
*/
        [Test]
        public void TestSetColor()
        {
            graphics.SetColor(Color.Red);
            Assert.AreEqual(Color.Red, graphics.Color);
        }

        [Test]
        public void TestFillRect()
        {
            graphics.FillRect(10, 10, 20, 20);
            HSSFSimpleShape s = (HSSFSimpleShape)escherGroupA.Children[0];
            Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE, s.ShapeType);
            Assert.AreEqual(10, s.Anchor.Dx1);
            Assert.AreEqual(10, s.Anchor.Dy1);
            Assert.AreEqual(30, s.Anchor.Dy2);
            Assert.AreEqual(30, s.Anchor.Dx2);
        }

        [Test]
        public void TestDrawString()
        {
            graphics.DrawString("This is a Test", 10, 10);
            HSSFTextbox t = (HSSFTextbox)escherGroupA.Children[0];
            Assert.AreEqual("This is a Test", t.String.String);
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

            Assert.IsNotNull(patriarch);
            Assert.AreEqual(10, patriarch.X1);
            Assert.AreEqual(20, patriarch.Y1);
            Assert.AreEqual(30, patriarch.X2);
            Assert.AreEqual(40, patriarch.Y2);

            // Check the two groups too
            Assert.AreEqual(2, patriarch.CountOfAllChildren);
            Assert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            Assert.IsTrue(patriarch.Children[1] is HSSFShapeGroup);

            s1 = (HSSFShapeGroup)patriarch.Children[0];
            s2 = (HSSFShapeGroup)patriarch.Children[1];

            Assert.AreEqual(0, s1.X1);
            Assert.AreEqual(0, s1.Y1);
            Assert.AreEqual(1023, s1.X2);
            Assert.AreEqual(255, s1.Y2);
            Assert.AreEqual(0, s2.X1);
            Assert.AreEqual(0, s2.Y1);
            Assert.AreEqual(1023, s2.X2);
            Assert.AreEqual(255, s2.Y2);

            Assert.AreEqual(0, s1.Anchor.Dx1);
            Assert.AreEqual(0, s1.Anchor.Dy1);
            Assert.AreEqual(1022, s1.Anchor.Dx2);
            Assert.AreEqual(255, s1.Anchor.Dy2);
            Assert.AreEqual(20, s2.Anchor.Dx1);
            Assert.AreEqual(30, s2.Anchor.Dy1);
            Assert.AreEqual(500, s2.Anchor.Dx2);
            Assert.AreEqual(200, s2.Anchor.Dy2);


            // Write and re-load once more, to Check that's ok
            baos = new MemoryStream();
            workbook.Write(baos);
            workbook = new HSSFWorkbook(new MemoryStream(baos.ToArray()));
            s = (HSSFSheet)workbook.GetSheetAt(0);
            patriarch = (HSSFPatriarch)s.DrawingPatriarch;

            Assert.IsNotNull(patriarch);
            Assert.AreEqual(10, patriarch.X1);
            Assert.AreEqual(20, patriarch.Y1);
            Assert.AreEqual(30, patriarch.X2);
            Assert.AreEqual(40, patriarch.Y2);

            // Check the two groups too
            Assert.AreEqual(2, patriarch.CountOfAllChildren);
            Assert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            Assert.IsTrue(patriarch.Children[1] is HSSFShapeGroup);

            s1 = (HSSFShapeGroup)patriarch.Children[0];
            s2 = (HSSFShapeGroup)patriarch.Children[1];

            Assert.AreEqual(0, s1.X1);
            Assert.AreEqual(0, s1.Y1);
            Assert.AreEqual(1023, s1.X2);
            Assert.AreEqual(255, s1.Y2);
            Assert.AreEqual(0, s2.X1);
            Assert.AreEqual(0, s2.Y1);
            Assert.AreEqual(1023, s2.X2);
            Assert.AreEqual(255, s2.Y2);

            Assert.AreEqual(0, s1.Anchor.Dx1);
            Assert.AreEqual(0, s1.Anchor.Dy1);
            Assert.AreEqual(1022, s1.Anchor.Dx2);
            Assert.AreEqual(255, s1.Anchor.Dy2);
            Assert.AreEqual(20, s2.Anchor.Dx1);
            Assert.AreEqual(30, s2.Anchor.Dy1);
            Assert.AreEqual(500, s2.Anchor.Dx2);
            Assert.AreEqual(200, s2.Anchor.Dy2);

            // Change the positions of the first groups,
            //  but not of their anchors
            s1.SetCoordinates(2, 3, 1021, 242);

            baos = new MemoryStream();
            workbook.Write(baos);
            workbook = new HSSFWorkbook(new MemoryStream(baos.ToArray()));
            s =(HSSFSheet)workbook.GetSheetAt(0);
            patriarch = (HSSFPatriarch)s.DrawingPatriarch;

            Assert.IsNotNull(patriarch);
            Assert.AreEqual(10, patriarch.X1);
            Assert.AreEqual(20, patriarch.Y1);
            Assert.AreEqual(30, patriarch.X2);
            Assert.AreEqual(40, patriarch.Y2);

            // Check the two groups too
            Assert.AreEqual(2, patriarch.CountOfAllChildren);
            Assert.AreEqual(2, patriarch.Children.Count);
            Assert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            Assert.IsTrue(patriarch.Children[1] is HSSFShapeGroup);

            s1 = (HSSFShapeGroup)patriarch.Children[0];
            s2 = (HSSFShapeGroup)patriarch.Children[1];

            Assert.AreEqual(2, s1.X1);
            Assert.AreEqual(3, s1.Y1);
            Assert.AreEqual(1021, s1.X2);
            Assert.AreEqual(242, s1.Y2);
            Assert.AreEqual(0, s2.X1);
            Assert.AreEqual(0, s2.Y1);
            Assert.AreEqual(1023, s2.X2);
            Assert.AreEqual(255, s2.Y2);

            Assert.AreEqual(0, s1.Anchor.Dx1);
            Assert.AreEqual(0, s1.Anchor.Dy1);
            Assert.AreEqual(1022, s1.Anchor.Dx2);
            Assert.AreEqual(255, s1.Anchor.Dy2);
            Assert.AreEqual(20, s2.Anchor.Dx1);
            Assert.AreEqual(30, s2.Anchor.Dy1);
            Assert.AreEqual(500, s2.Anchor.Dx2);
            Assert.AreEqual(200, s2.Anchor.Dy2);


            // Now Add some text to one group, and some more
            //  to the base, and Check we can get it back again
            HSSFTextbox tbox1 =
                patriarch.CreateTextbox(new HSSFClientAnchor(1, 2, 3, 4, (short)0, 0, (short)0, 0)) as HSSFTextbox;
            tbox1.String = (new HSSFRichTextString("I am text box 1"));
            HSSFTextbox tbox2 =
                s2.CreateTextbox(new HSSFChildAnchor(41, 42, 43, 44)) as HSSFTextbox;
            tbox2.String = (new HSSFRichTextString("This is text box 2"));

            Assert.AreEqual(3, patriarch.Children.Count);


            baos = new MemoryStream();
            workbook.Write(baos);
            workbook = new HSSFWorkbook(new MemoryStream(baos.ToArray()));
            s = (HSSFSheet)workbook.GetSheetAt(0);

            patriarch = (HSSFPatriarch)s.DrawingPatriarch;

            Assert.IsNotNull(patriarch);
            Assert.AreEqual(10, patriarch.X1);
            Assert.AreEqual(20, patriarch.Y1);
            Assert.AreEqual(30, patriarch.X2);
            Assert.AreEqual(40, patriarch.Y2);

            // Check the two groups and the text
            // Result of patriarch.countOfAllChildren() makes no sense: 
            // Returns 4 for 2 empty groups + 1 TextBox.
            //Assert.AreEqual(3, patriarch.CountOfAllChildren);
            Assert.AreEqual(3, patriarch.Children.Count);

            // Should be two groups and a text
            Assert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            Assert.IsTrue(patriarch.Children[1] is HSSFShapeGroup);
            Assert.IsTrue(patriarch.Children[2] is HSSFTextbox);
    

            s1 = (HSSFShapeGroup)patriarch.Children[0];
            tbox1 = (HSSFTextbox)patriarch.Children[2];

            s2 = (HSSFShapeGroup)patriarch.Children[1];

            Assert.AreEqual(2, s1.X1);
            Assert.AreEqual(3, s1.Y1);
            Assert.AreEqual(1021, s1.X2);
            Assert.AreEqual(242, s1.Y2);
            Assert.AreEqual(0, s2.X1);
            Assert.AreEqual(0, s2.Y1);
            Assert.AreEqual(1023, s2.X2);
            Assert.AreEqual(255, s2.Y2);

            Assert.AreEqual(0, s1.Anchor.Dx1);
            Assert.AreEqual(0, s1.Anchor.Dy1);
            Assert.AreEqual(1022, s1.Anchor.Dx2);
            Assert.AreEqual(255, s1.Anchor.Dy2);
            Assert.AreEqual(20, s2.Anchor.Dx1);
            Assert.AreEqual(30, s2.Anchor.Dy1);
            Assert.AreEqual(500, s2.Anchor.Dx2);
            Assert.AreEqual(200, s2.Anchor.Dy2);

            // Not working just yet
            //Assert.AreEqual("I am text box 1", tbox1.String.String);
        }
    }
}