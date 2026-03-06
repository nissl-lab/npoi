/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XDDF.UserModel.Text
{
    using NPOI.XDDF.UserModel.Text;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    public class TestXDDFTextRun
    {

        [Test]
        [Ignore("Need XSLF support")]
        public void TestTextRunPropertiesInSlide()
        {

            //try (XMLSlideShow ppt = new XMLSlideShow()) {
            //    XSLFSlide slide = ppt.CreateSlide();
            //    XSLFTextShape sh = slide.CreateAutoShape();
            //    sh.addNewTextParagraph();

            //    XDDFTextBody body = sh.TextBody;
            //    XDDFTextParagraph para = body.GetParagraph(0);
            //    XDDFTextRun r = para.AppendRegularRun("text");
            //    ClassicAssert.AreEqual(LocaleUtil.UserLocale.ToLanguageTag(), r.Language.ToLanguageTag());

            //    ClassicAssert.IsNull(r.CharacterSpacing);
            //    r.SetCharacterSpacing(3.0);
            //    ClassicAssert.AreEqual(3., r.CharacterSpacing, 0);
            //    r.SetCharacterSpacing(-3.0);
            //    ClassicAssert.AreEqual(-3., r.CharacterSpacing, 0);
            //    r.SetCharacterSpacing(0.0);
            //    ClassicAssert.AreEqual(0., r.CharacterSpacing, 0);

            //    ClassicAssert.AreEqual(11.0, r.FontSize, 0);
            //    r.SetFontSize(13.0);
            //    ClassicAssert.AreEqual(13.0, r.FontSize, 0);

            //    ClassicAssert.IsFalse(r.IsSuperscript());
            //    r.SetSuperscript(0.8);
            //    ClassicAssert.IsTrue(r.IsSuperscript());
            //    r.SetSuperscript(null);
            //    ClassicAssert.IsFalse(r.IsSuperscript());

            //    ClassicAssert.IsFalse(r.IsSubscript());
            //    r.SetSubscript(0.7);
            //    ClassicAssert.IsTrue(r.IsSubscript());
            //    r.SetSubscript(null);
            //    ClassicAssert.IsFalse(r.IsSubscript());

            //    r.SetBaseline(0.9);
            //    ClassicAssert.IsTrue(r.IsSuperscript());
            //    r.SetBaseline(-0.6);
            //    ClassicAssert.IsTrue(r.IsSubscript());
            //}
        }

        [Test]
        public void TestTextRunPropertiesInSheet()
        {

            using(XSSFWorkbook wb = new XSSFWorkbook())
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
                XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;

                XSSFTextBox shape = Drawing.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));

                shape.AddNewTextParagraph().AddNewTextRun().Text = "Line 1";

                XDDFTextBody body = shape.TextBody;
                XDDFTextParagraph para = body.GetParagraph(1);
                List<XDDFTextRun> runs = para.TextRuns;
                ClassicAssert.AreEqual(1, runs.Count);

                XDDFTextRun run = runs[0];
                ClassicAssert.AreEqual("Line 1", run.Text);

                ClassicAssert.IsFalse(run.IsStrikeThrough);
                run.StrikeThrough = StrikeType.SingleStrike;
                ClassicAssert.IsTrue(run.IsStrikeThrough);
                run.StrikeThrough = StrikeType.NoStrike;
                ClassicAssert.IsFalse(run.IsStrikeThrough);

                ClassicAssert.IsFalse(run.IsCapitals);
                run.Capitals = CapsType.Small;
                ClassicAssert.IsTrue(run.IsCapitals);
                run.Capitals = CapsType.None;
                ClassicAssert.IsFalse(run.IsCapitals);

                ClassicAssert.IsFalse(run.IsBold);
                run.IsBold = true;
                ClassicAssert.IsTrue(run.IsBold);
                run.IsBold = false;
                ClassicAssert.IsFalse(run.IsBold);

                ClassicAssert.IsFalse(run.IsItalic);
                run.IsItalic = true;
                ClassicAssert.IsTrue(run.IsItalic);
                run.IsItalic = false;
                ClassicAssert.IsFalse(run.IsItalic);

                ClassicAssert.IsFalse(run.IsUnderline);
                run.Underline = UnderlineType.WavyDouble;
                ClassicAssert.IsTrue(run.IsUnderline);
                run.Underline = UnderlineType.None;
                ClassicAssert.IsFalse(run.IsUnderline);

                ClassicAssert.IsNotNull(run.Text);
            }
        }
    }
}
