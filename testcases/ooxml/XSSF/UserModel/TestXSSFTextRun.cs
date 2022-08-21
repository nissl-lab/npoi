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
    using NUnit.Framework;
    using System.Collections.Generic;
    using NPOI.XSSF.UserModel;
    using SixLabors.ImageSharp;
    using SixLabors.ImageSharp.PixelFormats;

    [TestFixture]
    public class TestXSSFTextRun
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

                XSSFTextParagraph para = shape.AddNewTextParagraph();
                para.AddNewTextRun().Text=("Line 1");

                List<XSSFTextRun> Runs = para.TextRuns;
                Assert.AreEqual(1, Runs.Count);
                XSSFTextRun run = Runs[0];
                Assert.AreEqual("Line 1", run.Text);

                //Assert.IsNotNull(run.ParentParagraph);
                //Assert.IsNotNull(run.XmlObject);
                Assert.IsNotNull(run.GetRPr());

                Assert.AreEqual(new Rgb24(0, 0, 0), run.FontColor);

                var color = new Rgb24(0, 255, 255);
                run.FontColor = color;
                Assert.AreEqual(color, run.FontColor);

                Assert.AreEqual(11.0, run.FontSize, 0.01);
                run.FontSize = 12.32;
                Assert.AreEqual(12.32, run.FontSize, 0.01);
                run.FontSize = -1.0;
                Assert.AreEqual(11.0, run.FontSize, 0.01);
                run.FontSize = -1.0;
                Assert.AreEqual(11.0, run.FontSize, 0.01);
                try
                {
                    run.FontSize = (/*setter*/0.9);
                    Assert.Fail("Should fail");
                }
                catch (ArgumentException e)
                {
                    Assert.IsTrue(e.Message.Contains("0.9"));
                }
                Assert.AreEqual(11.0, run.FontSize, 0.01);

                Assert.AreEqual(0.0, run.CharacterSpacing, 0.01);
                run.CharacterSpacing = (/*setter*/12.31);
                Assert.AreEqual(12.31, run.CharacterSpacing, 0.01);
                run.CharacterSpacing = (/*setter*/0.0);
                Assert.AreEqual(0.0, run.CharacterSpacing, 0.01);
                run.CharacterSpacing = (/*setter*/0.0);
                Assert.AreEqual(0.0, run.CharacterSpacing, 0.01);

                Assert.AreEqual("Calibri", run.FontFamily);
                run.SetFontFamily("Arial", (byte)1, (byte)1, false);
                Assert.AreEqual("Arial", run.FontFamily);
                run.SetFontFamily("Arial", unchecked((byte)-1), (byte)1, false);
                Assert.AreEqual("Arial", run.FontFamily);
                run.SetFontFamily("Arial", (byte)1, unchecked((byte)-1), false);
                Assert.AreEqual("Arial", run.FontFamily);
                run.SetFontFamily("Arial", (byte)1, (byte)1, true);
                Assert.AreEqual("Arial", run.FontFamily);
                run.SetFontFamily(null, (byte)1, (byte)1, false);
                Assert.AreEqual("Calibri", run.FontFamily);
                run.SetFontFamily(null, (byte)1, (byte)1, false);
                Assert.AreEqual("Calibri", run.FontFamily);

                run.SetFont("Arial");
                Assert.AreEqual("Arial", run.FontFamily);

                Assert.AreEqual((byte)0, run.PitchAndFamily);
                run.SetFont(null);
                Assert.AreEqual((byte)0, run.PitchAndFamily);

                Assert.IsFalse(run.IsStrikethrough);
                run.IsStrikethrough = (/*setter*/true);
                Assert.IsTrue(run.IsStrikethrough);
                run.IsStrikethrough = (/*setter*/false);
                Assert.IsFalse(run.IsStrikethrough);

                Assert.IsFalse(run.IsSuperscript);
                run.IsSuperscript = (/*setter*/true);
                Assert.IsTrue(run.IsSuperscript);
                run.IsSuperscript = (/*setter*/false);
                Assert.IsFalse(run.IsSuperscript);

                Assert.IsFalse(run.IsSubscript);
                run.IsSubscript = (/*setter*/true);
                Assert.IsTrue(run.IsSubscript);
                run.IsSubscript = (/*setter*/false);
                Assert.IsFalse(run.IsSubscript);

                Assert.AreEqual(TextCap.NONE, run.TextCap);

                Assert.IsFalse(run.IsBold);
                run.IsBold = (/*setter*/true);
                Assert.IsTrue(run.IsBold);
                run.IsBold = (/*setter*/false);
                Assert.IsFalse(run.IsBold);

                Assert.IsFalse(run.IsItalic);
                run.IsItalic = (/*setter*/true);
                Assert.IsTrue(run.IsItalic);
                run.IsItalic = (/*setter*/false);
                Assert.IsFalse(run.IsItalic);

                Assert.IsFalse(run.IsUnderline);
                run.IsUnderline = (/*setter*/true);
                Assert.IsTrue(run.IsUnderline);
                run.IsUnderline = (/*setter*/false);
                Assert.IsFalse(run.IsUnderline);

                Assert.IsNotNull(run.ToString());
            }
            finally
            {
                wb.Close();
            }
        }
    }

}