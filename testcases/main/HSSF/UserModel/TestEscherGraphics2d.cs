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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Drawing;
using System;
namespace TestCases.HSSF.UserModel
{



    /**
     * Tests the Graphics2d drawing capability.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestClass]
    public class TestEscherGraphics2d
    {
        private HSSFShapeGroup escherGroup;
        private EscherGraphics2d graphics;
        [TestInitialize]
        public void SetUp()
        {

            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Test");
            escherGroup = sheet.CreateDrawingPatriarch().CreateGroup(new HSSFClientAnchor(0, 0, 1023, 255, (short)0, 0, (short)0, 0));
            escherGroup = new HSSFShapeGroup(null, new HSSFChildAnchor());
            EscherGraphics g = new EscherGraphics(this.escherGroup, workbook, Color.black, 1.0f);
            graphics = new EscherGraphics2d(g);

        }

        public void TestDrawString()
        {
            graphics.drawString("This is a Test", 10, 10);
            HSSFTextbox t = (HSSFTextbox)escherGroup.GetChildren()[0];
            Assert.AreEqual("This is a Test", t.String.String);

            // Check that with a valid font, it's still ok
            Font font = new Font("Forte", 12, FontStyle.Regular);
            graphics.SetFont(font);
            graphics.drawString("This is another Test", 10, 10);

            // And Test with ones that need the style appending
            font = new Font("dialog", 12, FontStyle.Regular);
            graphics.SetFont(font);
            graphics.drawString("This is another Test", 10, 10);

            font = new Font("dialog", 12, FontStyle.Bold);
            graphics.SetFont(font);
            graphics.drawString("This is another Test", 10, 10);

            // But with an invalid font, we get an exception
            font = new Font("IamAmadeUPfont", 22, FontStyle.Regular);
            graphics.SetFont(font);
            try
            {
                graphics.drawString("This is another Test", 10, 10);
                Assert.Fail();
            }
            catch (ArgumentException)
            {
            }
        }

        public void TestFillRect()
        {
            graphics.fillRect(10, 10, 20, 20);
            HSSFSimpleShape s = (HSSFSimpleShape)escherGroup.GetChildren()[0];
            Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE, s.GetShapeType());
            Assert.AreEqual(10, s.GetAnchor().GetDx1());
            Assert.AreEqual(10, s.GetAnchor().GetDy1());
            Assert.AreEqual(30, s.GetAnchor().GetDy2());
            Assert.AreEqual(30, s.GetAnchor().GetDx2());
        }

        public void TestGetFontMetrics()
        {
            FontMetrics fontMetrics = graphics.GetFontMetrics(graphics.GetFont());
            if (graphics.GetFont().toString().indexOf("dialog") != -1 || graphics.GetFont().toString().indexOf("Dialog") != -1) // if dialog is returned we can't run the Test properly.
                return;
            Assert.AreEqual(7, fontMetrics.charWidth('X'));
            Assert.AreEqual("java.awt.Font[family=Arial,name=Arial,style=plain,size=10]", fontMetrics.GetFont().toString());
        }

        public void TestSetFont()
        {
            Font f = new Font("Helvetica", 12);
            graphics.SetFont(f);
            Assert.AreEqual(f, graphics.GetFont());
        }

        public void TestSetColor()
        {
            graphics.SetColor(Color.red);
            Assert.AreEqual(Color.red, graphics.GetColor());
        }

        public void TestGetFont()
        {
            Font f = graphics.GetFont();
            if (graphics.GetFont().toString().indexOf("dialog") != -1 || graphics.GetFont().toString().indexOf("Dialog") != -1) // if dialog is returned we can't run the Test properly.
                return;

            Assert.AreEqual("java.awt.Font[family=Arial,name=Arial,style=plain,size=10]", f.toString());
        }

        public void TestDraw()
        {
            graphics.draw(new Line2D.Double(10, 10, 20, 20));
            HSSFSimpleShape s = (HSSFSimpleShape)escherGroup.GetChildren()[0];
            Assert.IsTrue(s.GetShapeType() == HSSFSimpleShape.OBJECT_TYPE_LINE);
            Assert.AreEqual(10, s.GetAnchor().GetDx1());
            Assert.AreEqual(10, s.GetAnchor().GetDy1());
            Assert.AreEqual(20, s.GetAnchor().GetDx2());
            Assert.AreEqual(20, s.GetAnchor().GetDy2());
            System.Console.WriteLine("s = " + s);
        }
    }
}