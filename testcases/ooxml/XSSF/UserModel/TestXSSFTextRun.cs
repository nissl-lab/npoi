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
    public class TestXSSFTextRun
{
    [Test]
        [Test]
    public void TestXSSFTextParagraph(){
        XSSFWorkbook wb = new XSSFWorkbook();
        try {
            XSSFSheet sheet = wb.CreateSheet();
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch();
    
            XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));

            XSSFTextParagraph para = shape.AddNewTextParagraph();
            para.AddNewTextRun().Text=(/*setter*/"Line 1");

            List<XSSFTextRun> Runs = para.TextRuns;
            Assert.AreEqual(1, Runs.Count);
            XSSFTextRun run = Runs.Get(0);
            Assert.AreEqual("Line 1", Run.Text);
            
            Assert.IsNotNull(Run.ParentParagraph);
            Assert.IsNotNull(Run.XmlObject);
            Assert.IsNotNull(Run.RPr);
            
            Assert.AreEqual(new Color(0,0,0), Run.FontColor);
            
            Color color = new Color(0, 255, 255);
            Run.FontColor=(/*setter*/color);
            Assert.AreEqual(color, Run.FontColor);
            
            Assert.AreEqual(11.0, Run.FontSize, 0.01);
            Run.FontSize=(/*setter*/12.32);
            Assert.AreEqual(12.32, Run.FontSize, 0.01);
            Run.FontSize=(/*setter*/-1.0);
            Assert.AreEqual(11.0, Run.FontSize, 0.01);
            Run.FontSize=(/*setter*/-1.0);
            Assert.AreEqual(11.0, Run.FontSize, 0.01);
            try {
                Run.FontSize=(/*setter*/0.9);
                Assert.Fail("Should fail");
            } catch (ArgumentException e) {
                Assert.IsTrue(e.Message.Contains("0.9"));
            }
            Assert.AreEqual(11.0, Run.FontSize, 0.01);
            
            Assert.AreEqual(0.0, Run.CharacterSpacing, 0.01);
            Run.CharacterSpacing=(/*setter*/12.31);
            Assert.AreEqual(12.31, Run.CharacterSpacing, 0.01);
            Run.CharacterSpacing=(/*setter*/0.0);
            Assert.AreEqual(0.0, Run.CharacterSpacing, 0.01);
            Run.CharacterSpacing=(/*setter*/0.0);
            Assert.AreEqual(0.0, Run.CharacterSpacing, 0.01);
            
            Assert.AreEqual("Calibri", Run.FontFamily);
            Run.SetFontFamily("Arial", (byte)1, (byte)1, false);
            Assert.AreEqual("Arial", Run.FontFamily);
            Run.SetFontFamily("Arial", (byte)-1, (byte)1, false);
            Assert.AreEqual("Arial", Run.FontFamily);
            Run.SetFontFamily("Arial", (byte)1, (byte)-1, false);
            Assert.AreEqual("Arial", Run.FontFamily);
            Run.SetFontFamily("Arial", (byte)1, (byte)1, true);
            Assert.AreEqual("Arial", Run.FontFamily);
            Run.SetFontFamily(null, (byte)1, (byte)1, false);
            Assert.AreEqual("Calibri", Run.FontFamily);
            Run.SetFontFamily(null, (byte)1, (byte)1, false);
            Assert.AreEqual("Calibri", Run.FontFamily);

            Run.Font=(/*setter*/"Arial");
            Assert.AreEqual("Arial", Run.FontFamily);
            
            Assert.AreEqual((byte)0, Run.PitchAndFamily);
            Run.Font=(/*setter*/null);
            Assert.AreEqual((byte)0, Run.PitchAndFamily);
            
            Assert.IsFalse(Run.IsStrikethrough());
            Run.Strikethrough=(/*setter*/true);
            Assert.IsTrue(Run.IsStrikethrough());
            Run.Strikethrough=(/*setter*/false);
            Assert.IsFalse(Run.IsStrikethrough());

            Assert.IsFalse(Run.IsSuperscript());
            Run.Superscript=(/*setter*/true);
            Assert.IsTrue(Run.IsSuperscript());
            Run.Superscript=(/*setter*/false);
            Assert.IsFalse(Run.IsSuperscript());

            Assert.IsFalse(Run.IsSubscript());
            Run.Subscript=(/*setter*/true);
            Assert.IsTrue(Run.IsSubscript());
            Run.Subscript=(/*setter*/false);
            Assert.IsFalse(Run.IsSubscript());
            
            Assert.AreEqual(TextCap.NONE, Run.TextCap);

            Assert.IsFalse(Run.IsBold());
            Run.Bold=(/*setter*/true);
            Assert.IsTrue(Run.IsBold());
            Run.Bold=(/*setter*/false);
            Assert.IsFalse(Run.IsBold());

            Assert.IsFalse(Run.IsItalic());
            Run.Italic=(/*setter*/true);
            Assert.IsTrue(Run.IsItalic());
            Run.Italic=(/*setter*/false);
            Assert.IsFalse(Run.IsItalic());

            Assert.IsFalse(Run.IsUnderline());
            Run.Underline=(/*setter*/true);
            Assert.IsTrue(Run.IsUnderline());
            Run.Underline=(/*setter*/false);
            Assert.IsFalse(Run.IsUnderline());
            
            Assert.IsNotNull(Run.ToString());
        } finally {
            wb.Close();
        }
    }
}

