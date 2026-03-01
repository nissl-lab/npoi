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

namespace TestCases.XDDF.UserModel
{

    using NPOI.OpenXmlFormats.Dml;
    using NPOI.XDDF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    [TestFixture]
    public class TestXDDFColor
    {
        private static string XMLNS = "xmlns:a=\"http://schemas.Openxmlformats.org/drawingml/2006/main\"/>";

        [Test]
        [Ignore("Need XSLF support")]
        public void TestSchemeColor()
        {
            
            //try (XMLSlideShow ppt = new XMLSlideShow()) {
            //    XSLFTheme theme = ppt.CreateSlide().Theme;

            //    XDDFColor color = XDDFColor.forColorContainer(getThemeColor(theme, STSchemeColorVal.ACCENT_2));
            //    // accent2 in theme1.xml is <a:srgbClr val="C0504D"/>
            //    Diff d1 = DiffBuilder.compare(Input.fromString("<a:srgbClr val=\"C0504D\" " + XMLNS))
            //            .withTest(color.ColorContainer.ToString()).build();
            //    Assert.IsFalse(d1.ToString(), d1.HasDifferences());

            //    color = XDDFColor.forColorContainer(getThemeColor(theme, STSchemeColorVal.LT_1));
            //    Diff d2 = DiffBuilder.compare(Input.fromString("<a:sysClr lastClr=\"FFFFFF\" val=\"window\" " + XMLNS))
            //            .withTest(color.ColorContainer.ToString()).build();
            //    Assert.IsFalse(d2.ToString(), d2.HasDifferences());

            //    color = XDDFColor.forColorContainer(getThemeColor(theme, STSchemeColorVal.DK_1));
            //    Diff d3 = DiffBuilder.compare(Input.fromString("<a:sysClr lastClr=\"000000\" val=\"windowText\" " + XMLNS))
            //            .withTest(color.ColorContainer.ToString()).build();
            //    Assert.IsFalse(d3.ToString(), d3.HasDifferences());
            //}
        }
        //private CTColor GetThemeColor(XSLFTheme theme, STSchemeColorVal.Enum value) {
        //    // find referenced CTColor in the theme
        //    return theme.GetCTColor(value.ToString());
        //}

        [Test]
        public void TestPreset()
        {
            CT_Color xml = new CT_Color();
            xml.AddNewPrstClr().val = ST_PresetColorVal.aquamarine;
            string expected = XDDFColor.ForColorContainer(xml).GetXmlObject().ToString();
            XDDFColor built = XDDFColor.From(PresetColor.Aquamarine);
            ClassicAssert.AreEqual(expected, built.GetXmlObject().ToString());
        }

        [Test]
        public void TestSystemDefined()
        {
            CT_Color xml = new CT_Color();
            CT_SystemColor sys = xml.AddNewSysClr();
            sys.val = ST_SystemColorVal.captionText;
            string expected = XDDFColor.ForColorContainer(xml).GetXmlObject().ToString();

            XDDFColor built = new XDDFColorSystemDefined(sys, xml);
            ClassicAssert.AreEqual(expected, built.GetXmlObject().ToString());

            built = XDDFColor.From(SystemColor.CaptionText);
            ClassicAssert.AreEqual(expected, built.GetXmlObject().ToString());
        }

        [Test]
        public void TestRgbBinary()
        {
            CT_Color xml = new CT_Color();
            CT_SRgbColor color = xml.AddNewSrgbClr();
            byte[] bs = new byte[]{ unchecked((byte)-1), unchecked((byte)-1), unchecked((byte)-1)};
            color.val = bs;
            string expected = XDDFColor.ForColorContainer(xml).GetXmlObject().ToString();

            XDDFColor built = XDDFColor.From(bs);
            ClassicAssert.AreEqual(expected, built.GetXmlObject().ToString());
            ClassicAssert.AreEqual("FFFFFF", ((XDDFColorRgbBinary) built).ToRGBHex());
        }

        [Test]
        public void TestRgbPercent()
        {
            CT_Color xml = new CT_Color();
            CT_ScRgbColor color = xml.AddNewScRgbClr();
            color.r = 0;
            color.g = 0;
            color.b = 0;
            string expected = XDDFColor.ForColorContainer(xml).GetXmlObject().ToString();

            XDDFColorRgbPercent built = (XDDFColorRgbPercent) XDDFColor.From(-1, -1, -1);
            ClassicAssert.AreEqual(expected, built.GetXmlObject().ToString());
            ClassicAssert.AreEqual("000000", built.ToRGBHex());

            color.r = 100_000;
            color.g = 100_000;
            color.b = 100_000;
            expected = XDDFColor.ForColorContainer(xml).GetXmlObject().ToString();

            built = (XDDFColorRgbPercent) XDDFColor.From(654321, 654321, 654321);
            ClassicAssert.AreEqual(expected, built.GetXmlObject().ToString());
            ClassicAssert.AreEqual("FFFFFF", built.ToRGBHex());
        }
    }
}