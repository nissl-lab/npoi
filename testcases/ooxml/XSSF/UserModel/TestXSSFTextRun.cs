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
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System.Collections.Generic;
    using NPOI.XSSF.UserModel;
    using SixLabors.ImageSharp;
    using SixLabors.ImageSharp.PixelFormats;
    using System.Threading;

    [TestFixture]
    public class TestXSSFTextRun
    {
        [Test]
        public void TestXSSFTextParagraph()
        {
            Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
                XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

                XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4)) as XSSFTextBox;

                XSSFTextParagraph para = shape.AddNewTextParagraph();
                para.AddNewTextRun().Text=("Line 1");

                List<XSSFTextRun> Runs = para.TextRuns;
                ClassicAssert.AreEqual(1, Runs.Count);
                XSSFTextRun run = Runs[0];
                ClassicAssert.AreEqual("Line 1", run.Text);

                //ClassicAssert.IsNotNull(run.ParentParagraph);
                //ClassicAssert.IsNotNull(run.XmlObject);
                ClassicAssert.IsNotNull(run.GetRPr());

                ClassicAssert.AreEqual(new Rgb24(0, 0, 0), run.FontColor);

                var color = new Rgb24(0, 255, 255);
                run.FontColor = color;
                ClassicAssert.AreEqual(color, run.FontColor);

                ClassicAssert.AreEqual(11.0, run.FontSize, 0.01);
                run.FontSize = 12.32;
                ClassicAssert.AreEqual(12.32, run.FontSize, 0.01);
                run.FontSize = -1.0;
                ClassicAssert.AreEqual(11.0, run.FontSize, 0.01);
                run.FontSize = -1.0;
                ClassicAssert.AreEqual(11.0, run.FontSize, 0.01);
                try
                {
                    run.FontSize = (/*setter*/0.9);
                    Assert.Fail("Should fail");
                }
                catch (ArgumentException e)
                {
                    ClassicAssert.IsTrue(e.Message.Contains("0.9"));
                }
                ClassicAssert.AreEqual(11.0, run.FontSize, 0.01);

                ClassicAssert.AreEqual(0.0, run.CharacterSpacing, 0.01);
                run.CharacterSpacing = (/*setter*/12.31);
                ClassicAssert.AreEqual(12.31, run.CharacterSpacing, 0.01);
                run.CharacterSpacing = (/*setter*/0.0);
                ClassicAssert.AreEqual(0.0, run.CharacterSpacing, 0.01);
                run.CharacterSpacing = (/*setter*/0.0);
                ClassicAssert.AreEqual(0.0, run.CharacterSpacing, 0.01);

                ClassicAssert.AreEqual("Calibri", run.FontFamily);
                run.SetFontFamily("Arial", (byte)1, (byte)1, false);
                ClassicAssert.AreEqual("Arial", run.FontFamily);
                run.SetFontFamily("Arial", unchecked((byte)-1), (byte)1, false);
                ClassicAssert.AreEqual("Arial", run.FontFamily);
                run.SetFontFamily("Arial", (byte)1, unchecked((byte)-1), false);
                ClassicAssert.AreEqual("Arial", run.FontFamily);
                run.SetFontFamily("Arial", (byte)1, (byte)1, true);
                ClassicAssert.AreEqual("Arial", run.FontFamily);
                run.SetFontFamily(null, (byte)1, (byte)1, false);
                ClassicAssert.AreEqual("Calibri", run.FontFamily);
                run.SetFontFamily(null, (byte)1, (byte)1, false);
                ClassicAssert.AreEqual("Calibri", run.FontFamily);

                run.SetFont("Arial");
                ClassicAssert.AreEqual("Arial", run.FontFamily);

                ClassicAssert.AreEqual((byte)0, run.PitchAndFamily);
                run.SetFont(null);
                ClassicAssert.AreEqual((byte)0, run.PitchAndFamily);

                ClassicAssert.IsFalse(run.IsStrikethrough);
                run.IsStrikethrough = (/*setter*/true);
                ClassicAssert.IsTrue(run.IsStrikethrough);
                run.IsStrikethrough = (/*setter*/false);
                ClassicAssert.IsFalse(run.IsStrikethrough);

                ClassicAssert.IsFalse(run.IsSuperscript);
                run.IsSuperscript = (/*setter*/true);
                ClassicAssert.IsTrue(run.IsSuperscript);
                run.IsSuperscript = (/*setter*/false);
                ClassicAssert.IsFalse(run.IsSuperscript);

                ClassicAssert.IsFalse(run.IsSubscript);
                run.IsSubscript = (/*setter*/true);
                ClassicAssert.IsTrue(run.IsSubscript);
                run.IsSubscript = (/*setter*/false);
                ClassicAssert.IsFalse(run.IsSubscript);

                ClassicAssert.AreEqual(TextCap.NONE, run.TextCap);

                ClassicAssert.IsFalse(run.IsBold);
                run.IsBold = (/*setter*/true);
                ClassicAssert.IsTrue(run.IsBold);
                run.IsBold = (/*setter*/false);
                ClassicAssert.IsFalse(run.IsBold);

                ClassicAssert.IsFalse(run.IsItalic);
                run.IsItalic = (/*setter*/true);
                ClassicAssert.IsTrue(run.IsItalic);
                run.IsItalic = (/*setter*/false);
                ClassicAssert.IsFalse(run.IsItalic);

                ClassicAssert.IsFalse(run.IsUnderline);
                run.IsUnderline = (/*setter*/true);
                ClassicAssert.IsTrue(run.IsUnderline);
                run.IsUnderline = (/*setter*/false);
                ClassicAssert.IsFalse(run.IsUnderline);

                ClassicAssert.IsNotNull(run.ToString());
            }
            finally
            {
                wb.Close();
            }
        }
    }

}