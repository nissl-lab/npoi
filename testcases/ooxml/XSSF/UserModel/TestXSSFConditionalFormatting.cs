using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using SixLabors.ImageSharp;
using System;
/*
*  ====================================================================
*    Licensed to the Apache Software Foundation (ASF) under one or more
*    contributor license agreements.  See the NOTICE file distributed with
*    this work for additional information regarding copyright ownership.
*    The ASF licenses this file to You under the Apache License, Version 2.0
*    (the "License"); you may not use this file except in compliance with
*    the License.  You may obtain a copy of the License at
*
*        http://www.apache.org/licenses/LICENSE-2.0
*
*    Unless required by applicable law or agreed to in writing, software
*    distributed under the License is distributed on an "AS IS" BASIS,
*    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
*    See the License for the specific language governing permissions and
*    limitations under the License.
* ====================================================================
*/
using TestCases.SS.UserModel;
namespace TestCases.XSSF.UserModel
{
    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFConditionalFormatting : BaseTestConditionalFormatting
    {
        protected override void AssertColour(String hexExpected, IColor actual)
        {
            ClassicAssert.IsNotNull(actual, "Colour must be given");
            XSSFColor colour = (XSSFColor)actual;
            if (hexExpected.Length == 8)
            {
                ClassicAssert.AreEqual(hexExpected, colour.ARGBHex);
            }
            else
            {
                ClassicAssert.AreEqual(hexExpected, colour.ARGBHex.Substring(2));
            }
        }

        public TestXSSFConditionalFormatting()
            : base(XSSFITestDataProvider.instance)
        {

        }
        [Test]
        public void TestRead()
        {
            this.TestRead("WithConditionalFormatting.xlsx");
        }
        [Test]
        public void TestReadOffice2007()
        {
            // TODO Bring the XSSF support up to the same level
            TestReadOffice2007("NewStyleConditionalFormattings.xlsx");
        }

        private static Color PEAK_ORANGE = Color.FromRgb(255, 239, 221);

        [Test]
        public void TestFontFormattingColor()
        {
            XSSFWorkbook wb = XSSFITestDataProvider.instance.CreateWorkbook() as XSSFWorkbook;
            ISheet sheet = wb.CreateSheet();

            ISheetConditionalFormatting formatting = sheet.SheetConditionalFormatting;

            // the conditional formatting is not automatically added when it is created...
            ClassicAssert.AreEqual(0, formatting.NumConditionalFormattings);
            IConditionalFormattingRule formattingRule = formatting.CreateConditionalFormattingRule("A1");
            ClassicAssert.AreEqual(0, formatting.NumConditionalFormattings);

            // adding the formatting makes it available
            int idx = formatting.AddConditionalFormatting(new CellRangeAddress[] {}, formattingRule);

            // verify that it can be accessed now
            ClassicAssert.AreEqual(0, idx);
            ClassicAssert.AreEqual(1, formatting.NumConditionalFormattings);
            ClassicAssert.AreEqual(1, formatting.GetConditionalFormattingAt(idx).NumberOfRules);

            // this is confusing: the rule is not connected to the sheet, changes are not applied
            // so we need to use SetRule() explicitly!
            IFontFormatting fontFmt = formattingRule.CreateFontFormatting();
            ClassicAssert.IsNotNull(formattingRule.FontFormatting);
            ClassicAssert.AreEqual(1, formatting.GetConditionalFormattingAt(idx).NumberOfRules);
            formatting.GetConditionalFormattingAt(idx).SetRule(0, formattingRule);
            ClassicAssert.IsNotNull(formatting.GetConditionalFormattingAt(idx).GetRule(0).FontFormatting);

            fontFmt.SetFontStyle(true, false);

            ClassicAssert.AreEqual(-1, fontFmt.FontColorIndex);

            //fontFmt.SetFontColorIndex((short)11);
            ExtendedColor extendedColor = new XSSFColor(PEAK_ORANGE, wb.GetStylesSource().IndexedColors);
            fontFmt.FontColor = extendedColor;

            IPatternFormatting patternFmt = formattingRule.CreatePatternFormatting();
            ClassicAssert.IsNotNull(patternFmt);
            patternFmt.FillBackgroundColorColor = extendedColor;

            ClassicAssert.AreEqual(1, formatting.GetConditionalFormattingAt(0).NumberOfRules);
            ClassicAssert.IsNotNull(formatting.GetConditionalFormattingAt(0).GetRule(0).FontFormatting);
            ClassicAssert.IsNotNull(formatting.GetConditionalFormattingAt(0).GetRule(0).FontFormatting.FontColor);
            ClassicAssert.IsNotNull(formatting.GetConditionalFormattingAt(0).GetRule(0).PatternFormatting.FillBackgroundColorColor);

            checkFontFormattingColorWriteOutAndReadBack(wb, extendedColor);
        }

        private void checkFontFormattingColorWriteOutAndReadBack(IWorkbook wb, ExtendedColor extendedColor)
        {
            IWorkbook wbBack = XSSFITestDataProvider.instance.WriteOutAndReadBack(wb);
            ClassicAssert.IsNotNull(wbBack);

            ClassicAssert.AreEqual(1, wbBack.GetSheetAt(0).SheetConditionalFormatting.NumConditionalFormattings);
            IConditionalFormatting formattingBack = wbBack.GetSheetAt(0).SheetConditionalFormatting.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual(1, wbBack.GetSheetAt(0).SheetConditionalFormatting.GetConditionalFormattingAt(0).NumberOfRules);
            IConditionalFormattingRule ruleBack = formattingBack.GetRule(0);
            IFontFormatting fontFormattingBack = ruleBack.FontFormatting;
            ClassicAssert.IsNotNull(formattingBack);
            ClassicAssert.IsNotNull(fontFormattingBack.FontColor);
            ClassicAssert.AreEqual(extendedColor, fontFormattingBack.FontColor);
            ClassicAssert.AreEqual(extendedColor, ruleBack.PatternFormatting.FillBackgroundColorColor);
        }
    }
}