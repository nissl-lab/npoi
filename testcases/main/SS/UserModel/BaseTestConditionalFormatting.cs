/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
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

namespace TestCases.SS.UserModel
{
    using System;

    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using TestCases.SS;
    using NPOI.HSSF.Util;
    using NPOI.HSSF.UserModel;

    /**
     * @author Dmitriy Kumshayev
     * @author Yegor Kozlov
     */
    [TestFixture]
    public abstract class BaseTestConditionalFormatting
    {
        private ITestDataProvider _testDataProvider;
        public BaseTestConditionalFormatting()
        {
            _testDataProvider = TestCases.HSSF.HSSFITestDataProvider.Instance;
        }
        public BaseTestConditionalFormatting(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }
        protected abstract void AssertColour(String hexExpected, IColor actual);
        [Test]
        public void TestBasic()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            ISheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;

            ClassicAssert.AreEqual(0, sheetCF.NumConditionalFormattings);
            try
            {
                ClassicAssert.IsNull(sheetCF.GetConditionalFormattingAt(0));
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.IsTrue(e.Message.StartsWith("Specified CF index 0 is outside the allowable range"));
            }

            try
            {
                sheetCF.RemoveConditionalFormatting(0);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.IsTrue(e.Message.StartsWith("Specified CF index 0 is outside the allowable range"));
            }

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("1");
            IConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule("2");
            IConditionalFormattingRule rule3 = sheetCF.CreateConditionalFormattingRule("3");
            IConditionalFormattingRule rule4 = sheetCF.CreateConditionalFormattingRule("4");
            try
            {
                sheetCF.AddConditionalFormatting(null, rule1);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.IsTrue(e.Message.StartsWith("regions must not be null"));
            }
            try
            {
                sheetCF.AddConditionalFormatting(
                        new CellRangeAddress[] { CellRangeAddress.ValueOf("A1:A3") },
                        (IConditionalFormattingRule)null);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.IsTrue(e.Message.StartsWith("cfRules must not be null"));
            }

            try
            {
                sheetCF.AddConditionalFormatting(
                        new CellRangeAddress[] { CellRangeAddress.ValueOf("A1:A3") },
                        new IConditionalFormattingRule[0]);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.IsTrue(e.Message.StartsWith("cfRules must not be empty"));
            }

            wb.Close();
        }

        /**
         * Test format conditions based on a bool formula
         */
        [Test]
        public void TestBooleanFormulaConditions()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            ISheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("SUM(A1:A5)>10");
            ClassicAssert.AreEqual(ConditionType.Formula, rule1.ConditionType);
            ClassicAssert.AreEqual("SUM(A1:A5)>10", rule1.Formula1);
            int formatIndex1 = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("B1"),
                        CellRangeAddress.ValueOf("C3"),
                }, rule1);
            ClassicAssert.AreEqual(0, formatIndex1);
            ClassicAssert.AreEqual(1, sheetCF.NumConditionalFormattings);
            CellRangeAddress[] ranges1 = sheetCF.GetConditionalFormattingAt(formatIndex1).GetFormattingRanges();
            ClassicAssert.AreEqual(2, ranges1.Length);
            ClassicAssert.AreEqual("B1", ranges1[0].FormatAsString());
            ClassicAssert.AreEqual("C3", ranges1[1].FormatAsString());

            // adjacent Address are merged
            int formatIndex2 = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("B1"),
                        CellRangeAddress.ValueOf("B2"),
                        CellRangeAddress.ValueOf("B3"),
                }, rule1);
            ClassicAssert.AreEqual(1, formatIndex2);
            ClassicAssert.AreEqual(2, sheetCF.NumConditionalFormattings);
            CellRangeAddress[] ranges2 = sheetCF.GetConditionalFormattingAt(formatIndex2).GetFormattingRanges();
            ClassicAssert.AreEqual(1, ranges2.Length);
            ClassicAssert.AreEqual("B1:B3", ranges2[0].FormatAsString());

            wb.Close();
        }
        [Test]
        public void TestSingleFormulaConditions()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            ISheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.Equal, "SUM(A1:A5)+10");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule1.ConditionType);
            ClassicAssert.AreEqual("SUM(A1:A5)+10", rule1.Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.Equal, rule1.ComparisonOperation);

            IConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NotEqual, "15");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule2.ConditionType);
            ClassicAssert.AreEqual("15", rule2.Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.NotEqual, rule2.ComparisonOperation);

            IConditionalFormattingRule rule3 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NotEqual, "15");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule3.ConditionType);
            ClassicAssert.AreEqual("15", rule3.Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.NotEqual, rule3.ComparisonOperation);

            IConditionalFormattingRule rule4 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.GreaterThan, "0");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule4.ConditionType);
            ClassicAssert.AreEqual("0", rule4.Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.GreaterThan, rule4.ComparisonOperation);

            IConditionalFormattingRule rule5 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.LessThan, "0");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule5.ConditionType);
            ClassicAssert.AreEqual("0", rule5.Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.LessThan, rule5.ComparisonOperation);

            IConditionalFormattingRule rule6 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.GreaterThanOrEqual, "0");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule6.ConditionType);
            ClassicAssert.AreEqual("0", rule6.Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.GreaterThanOrEqual, rule6.ComparisonOperation);

            IConditionalFormattingRule rule7 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.LessThanOrEqual, "0");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule7.ConditionType);
            ClassicAssert.AreEqual("0", rule7.Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.LessThanOrEqual, rule7.ComparisonOperation);

            IConditionalFormattingRule rule8 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.Between, "0", "5");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule8.ConditionType);
            ClassicAssert.AreEqual("0", rule8.Formula1);
            ClassicAssert.AreEqual("5", rule8.Formula2);
            ClassicAssert.AreEqual(ComparisonOperator.Between, rule8.ComparisonOperation);

            IConditionalFormattingRule rule9 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NotBetween, "0", "5");
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule9.ConditionType);
            ClassicAssert.AreEqual("0", rule9.Formula1);
            ClassicAssert.AreEqual("5", rule9.Formula2);
            ClassicAssert.AreEqual(ComparisonOperator.NotBetween, rule9.ComparisonOperation);

            wb.Close();
        }
        [Test]
        public void TestCopy()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            ISheet sheet2 = wb.CreateSheet();
            ISheetConditionalFormatting sheet1CF = sheet1.SheetConditionalFormatting;
            ISheetConditionalFormatting sheet2CF = sheet2.SheetConditionalFormatting;
            ClassicAssert.AreEqual(0, sheet1CF.NumConditionalFormattings);
            ClassicAssert.AreEqual(0, sheet2CF.NumConditionalFormattings);

            IConditionalFormattingRule rule1 = sheet1CF.CreateConditionalFormattingRule(
                    ComparisonOperator.Equal, "SUM(A1:A5)+10");

            IConditionalFormattingRule rule2 = sheet1CF.CreateConditionalFormattingRule(
                    ComparisonOperator.NotEqual, "15");

            // adjacent Address are merged
            int formatIndex = sheet1CF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("A1:A5"),
                        CellRangeAddress.ValueOf("C1:C5")
                }, rule1, rule2);
            ClassicAssert.AreEqual(0, formatIndex);
            ClassicAssert.AreEqual(1, sheet1CF.NumConditionalFormattings);

            ClassicAssert.AreEqual(0, sheet2CF.NumConditionalFormattings);
            sheet2CF.AddConditionalFormatting(sheet1CF.GetConditionalFormattingAt(formatIndex));
            ClassicAssert.AreEqual(1, sheet2CF.NumConditionalFormattings);

            IConditionalFormatting sheet2cf = sheet2CF.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual(2, sheet2cf.NumberOfRules);
            ClassicAssert.AreEqual("SUM(A1:A5)+10", sheet2cf.GetRule(0).Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.Equal, sheet2cf.GetRule(0).ComparisonOperation);
            ClassicAssert.AreEqual(ConditionType.CellValueIs, sheet2cf.GetRule(0).ConditionType);
            ClassicAssert.AreEqual("15", sheet2cf.GetRule(1).Formula1);
            ClassicAssert.AreEqual(ComparisonOperator.NotEqual, sheet2cf.GetRule(1).ComparisonOperation);
            ClassicAssert.AreEqual(ConditionType.CellValueIs, sheet2cf.GetRule(1).ConditionType);

            wb.Close();
        }
        [Test]
        public void TestRemove()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            ISheetConditionalFormatting sheetCF = sheet1.SheetConditionalFormatting;
            ClassicAssert.AreEqual(0, sheetCF.NumConditionalFormattings);

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.Equal, "SUM(A1:A5)");

            // adjacent Address are merged
            int formatIndex = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("A1:A5")
                }, rule1);
            ClassicAssert.AreEqual(0, formatIndex);
            ClassicAssert.AreEqual(1, sheetCF.NumConditionalFormattings);
            sheetCF.RemoveConditionalFormatting(0);
            ClassicAssert.AreEqual(0, sheetCF.NumConditionalFormattings);
            try
            {
                ClassicAssert.IsNull(sheetCF.GetConditionalFormattingAt(0));
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.IsTrue(e.Message.StartsWith("Specified CF index 0 is outside the allowable range"));
            }

            formatIndex = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("A1:A5")
                }, rule1);
            ClassicAssert.AreEqual(0, formatIndex);
            ClassicAssert.AreEqual(1, sheetCF.NumConditionalFormattings);
            sheetCF.RemoveConditionalFormatting(0);
            ClassicAssert.AreEqual(0, sheetCF.NumConditionalFormattings);
            try
            {
                ClassicAssert.IsNull(sheetCF.GetConditionalFormattingAt(0));
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                ClassicAssert.IsTrue(e.Message.StartsWith("Specified CF index 0 is outside the allowable range"));
            }

            wb.Close();
        }
        [Test]
        public void TestCreateCF()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            String formula = "7";

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(formula);
            IFontFormatting fontFmt = rule1.CreateFontFormatting();
            fontFmt.SetFontStyle(true, false);

            IBorderFormatting bordFmt = rule1.CreateBorderFormatting();
            bordFmt.BorderBottom = (/*setter*/BorderStyle.Thin);
            bordFmt.BorderTop = (/*setter*/BorderStyle.Thick);
            bordFmt.BorderLeft = (/*setter*/BorderStyle.Dashed);
            bordFmt.BorderRight = (/*setter*/BorderStyle.Dotted);

            IPatternFormatting patternFmt = rule1.CreatePatternFormatting();
            patternFmt.FillBackgroundColor = (/*setter*/HSSFColor.Yellow.Index);


            IConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.Between, "1", "2");
            IConditionalFormattingRule[] cfRules =
        {
            rule1, rule2
        };

            short col = 1;
            CellRangeAddress[] regions = {
            new CellRangeAddress(0, 65535, col, col)
        };

            sheetCF.AddConditionalFormatting(regions, cfRules);
            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            ClassicAssert.AreEqual(2, sheetCF.NumConditionalFormattings);
            sheetCF.RemoveConditionalFormatting(1);
            ClassicAssert.AreEqual(1, sheetCF.NumConditionalFormattings);
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.IsNotNull(cf);

            regions = cf.GetFormattingRanges();
            ClassicAssert.IsNotNull(regions);
            ClassicAssert.AreEqual(1, regions.Length);
            CellRangeAddress r = regions[0];
            ClassicAssert.AreEqual(1, r.FirstColumn);
            ClassicAssert.AreEqual(1, r.LastColumn);
            ClassicAssert.AreEqual(0, r.FirstRow);
            ClassicAssert.AreEqual(65535, r.LastRow);

            ClassicAssert.AreEqual(2, cf.NumberOfRules);

            rule1 = cf.GetRule(0);
            ClassicAssert.AreEqual("7", rule1.Formula1);
            ClassicAssert.IsNull(rule1.Formula2);

            IFontFormatting r1fp = rule1.FontFormatting;
            ClassicAssert.IsNotNull(r1fp);

            ClassicAssert.IsTrue(r1fp.IsItalic);
            ClassicAssert.IsFalse(r1fp.IsBold);

            IBorderFormatting r1bf = rule1.BorderFormatting;
            ClassicAssert.IsNotNull(r1bf);
            ClassicAssert.AreEqual(BorderStyle.Thin, r1bf.BorderBottom);
            ClassicAssert.AreEqual(BorderStyle.Thick, r1bf.BorderTop);
            ClassicAssert.AreEqual(BorderStyle.Dashed, r1bf.BorderLeft);
            ClassicAssert.AreEqual(BorderStyle.Dotted, r1bf.BorderRight);

            IPatternFormatting r1pf = rule1.PatternFormatting;
            ClassicAssert.IsNotNull(r1pf);
            //        ClassicAssert.AreEqual(HSSFColor.Yellow.index,r1pf.FillBackgroundColor);

            rule2 = cf.GetRule(1);
            ClassicAssert.AreEqual("2", rule2.Formula2);
            ClassicAssert.AreEqual("1", rule2.Formula1);

            workbook.Close();
        }
        [Test]
        public void TestClone()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            String formula = "7";

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(formula);
            IFontFormatting fontFmt = rule1.CreateFontFormatting();
            fontFmt.SetFontStyle(true, false);

            IPatternFormatting patternFmt = rule1.CreatePatternFormatting();
            patternFmt.FillBackgroundColor = (/*setter*/HSSFColor.Yellow.Index);


            IConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.Between, "1", "2");
            IConditionalFormattingRule[] cfRules =
        {
            rule1, rule2
        };

            short col = 1;
            CellRangeAddress[] regions = {
            new CellRangeAddress(0, 65535, col, col)
        };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            try
            {
                wb.CloneSheet(0);
                ClassicAssert.AreEqual(2, wb.NumberOfSheets);
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf("needs to define a clone method") > 0)
                {
                    Assert.Fail("Identified bug 45682");
                }
                throw;
            }
            finally
            {
                wb.Close();
            }
        }
        [Test]
        public void TestShiftRows()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.Between, "SUM(A10:A15)", "1+SUM(B16:B30)");
            IFontFormatting fontFmt = rule1.CreateFontFormatting();
            fontFmt.SetFontStyle(true, false);

            IPatternFormatting patternFmt = rule1.CreatePatternFormatting();
            patternFmt.FillBackgroundColor = (/*setter*/HSSFColor.Yellow.Index);

            IConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(
                ComparisonOperator.Between, "SUM(A10:A15)", "1+SUM(B16:B30)");
            IBorderFormatting borderFmt = rule2.CreateBorderFormatting();
            borderFmt.BorderDiagonal = BorderStyle.Medium;


            CellRangeAddress[] regions = {
            new CellRangeAddress(2, 4, 0, 0), // A3:A5
        };
            sheetCF.AddConditionalFormatting(regions, rule1);
            sheetCF.AddConditionalFormatting(regions, rule2);

            // This row-shift should destroy the CF region
            sheet.ShiftRows(10, 20, -9);
            ClassicAssert.AreEqual(0, sheetCF.NumConditionalFormattings);

            // re-add the CF
            sheetCF.AddConditionalFormatting(regions, rule1);
            sheetCF.AddConditionalFormatting(regions, rule2);

            // This row shift should only affect the formulas
            sheet.ShiftRows(14, 17, 8);
            IConditionalFormatting cf1 = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual("SUM(A10:A23)", cf1.GetRule(0).Formula1);
            ClassicAssert.AreEqual("1+SUM(B24:B30)", cf1.GetRule(0).Formula2);
            IConditionalFormatting cf2 = sheetCF.GetConditionalFormattingAt(1);
            ClassicAssert.AreEqual("SUM(A10:A23)", cf2.GetRule(0).Formula1);
            ClassicAssert.AreEqual("1+SUM(B24:B30)", cf2.GetRule(0).Formula2);

            sheet.ShiftRows(0, 8, 21);
            cf1 = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual("SUM(A10:A21)", cf1.GetRule(0).Formula1);
            ClassicAssert.AreEqual("1+SUM(#REF!)", cf1.GetRule(0).Formula2);
            cf2 = sheetCF.GetConditionalFormattingAt(1);
            ClassicAssert.AreEqual("SUM(A10:A21)", cf2.GetRule(0).Formula1);
            ClassicAssert.AreEqual("1+SUM(#REF!)", cf2.GetRule(0).Formula2);

            wb.Close();
        }
        //

        protected void TestRead(string sampleFile)
        {

            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(sampleFile);
            ISheet sh = wb.GetSheet("CF");
            ISheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;
            ClassicAssert.AreEqual(3, sheetCF.NumConditionalFormattings);

            IConditionalFormatting cf1 = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual(2, cf1.NumberOfRules);

            CellRangeAddress[] regions1 = cf1.GetFormattingRanges();
            ClassicAssert.AreEqual(1, regions1.Length);
            ClassicAssert.AreEqual("A1:A8", regions1[0].FormatAsString());

            // CF1 has two rules: values less than -3 are bold-italic red, values greater than 3 are green
            IConditionalFormattingRule rule1 = cf1.GetRule(0);
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule1.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.GreaterThan, rule1.ComparisonOperation);
            ClassicAssert.AreEqual("3", rule1.Formula1);
            ClassicAssert.IsNull(rule1.Formula2);
            // Fills and borders are not Set
            ClassicAssert.IsNull(rule1.PatternFormatting);
            ClassicAssert.IsNull(rule1.BorderFormatting);

            IFontFormatting fmt1 = rule1.FontFormatting;
            //        ClassicAssert.AreEqual(HSSFColor.GREEN.index, fmt1.FontColorIndex);
            ClassicAssert.IsTrue(fmt1.IsBold);
            ClassicAssert.IsFalse(fmt1.IsItalic);

            IConditionalFormattingRule rule2 = cf1.GetRule(1);
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule2.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.LessThan, rule2.ComparisonOperation);
            ClassicAssert.AreEqual("-3", rule2.Formula1);
            ClassicAssert.IsNull(rule2.Formula2);
            ClassicAssert.IsNull(rule2.PatternFormatting);
            ClassicAssert.IsNull(rule2.BorderFormatting);

            IFontFormatting fmt2 = rule2.FontFormatting;
            //        ClassicAssert.AreEqual(HSSFColor.Red.index, fmt2.FontColorIndex);
            ClassicAssert.IsTrue(fmt2.IsBold);
            ClassicAssert.IsTrue(fmt2.IsItalic);


            IConditionalFormatting cf2 = sheetCF.GetConditionalFormattingAt(1);
            ClassicAssert.AreEqual(1, cf2.NumberOfRules);
            CellRangeAddress[] regions2 = cf2.GetFormattingRanges();
            ClassicAssert.AreEqual(1, regions2.Length);
            ClassicAssert.AreEqual("B9", regions2[0].FormatAsString());

            IConditionalFormattingRule rule3 = cf2.GetRule(0);
            ClassicAssert.AreEqual(ConditionType.Formula, rule3.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.NoComparison, rule3.ComparisonOperation);
            ClassicAssert.AreEqual("$A$8>5", rule3.Formula1);
            ClassicAssert.IsNull(rule3.Formula2);

            IFontFormatting fmt3 = rule3.FontFormatting;
            //        ClassicAssert.AreEqual(HSSFColor.Red.index, fmt3.FontColorIndex);
            ClassicAssert.IsTrue(fmt3.IsBold);
            ClassicAssert.IsTrue(fmt3.IsItalic);

            IPatternFormatting fmt4 = rule3.PatternFormatting;
            //        ClassicAssert.AreEqual(HSSFColor.LIGHT_CORNFLOWER_BLUE.index, fmt4.FillBackgroundColor);
            //        ClassicAssert.AreEqual(HSSFColor.Automatic.index, fmt4.FillForegroundColor);
            ClassicAssert.AreEqual(FillPattern.NoFill, fmt4.FillPattern);
            // borders are not Set
            ClassicAssert.IsNull(rule3.BorderFormatting);

            IConditionalFormatting cf3 = sheetCF.GetConditionalFormattingAt(2);
            CellRangeAddress[] regions3 = cf3.GetFormattingRanges();
            ClassicAssert.AreEqual(1, regions3.Length);
            ClassicAssert.AreEqual("B1:B7", regions3[0].FormatAsString());
            ClassicAssert.AreEqual(2, cf3.NumberOfRules);

            IConditionalFormattingRule rule4 = cf3.GetRule(0);
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule4.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.LessThanOrEqual, rule4.ComparisonOperation);
            ClassicAssert.AreEqual("\"AAA\"", rule4.Formula1);
            ClassicAssert.IsNull(rule4.Formula2);

            IConditionalFormattingRule rule5 = cf3.GetRule(1);
            ClassicAssert.AreEqual(ConditionType.CellValueIs, rule5.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.Between, rule5.ComparisonOperation);
            ClassicAssert.AreEqual("\"A\"", rule5.Formula1);
            ClassicAssert.AreEqual("\"AAA\"", rule5.Formula2);

            wb.Close();
        }

        public void TestReadOffice2007(String filename)
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(filename);
            ISheet s = wb.GetSheet("CF");

            // Sanity check data
            ClassicAssert.AreEqual("Values", s.GetRow(0).GetCell(0).ToString());
            ClassicAssert.AreEqual("10", s.GetRow(2).GetCell(0).ToString());
            //ClassicAssert.AreEqual("10.0", s.GetRow(2).GetCell(0).ToString());

            // Check we found all the conditional formattings rules we should have
            ISheetConditionalFormatting sheetCF = s.SheetConditionalFormatting;
            int numCF = 3;
            int numCF12 = 15;
            int numCFEX = 0; // TODO This should be 1, but we don't support CFEX formattings yet
            ClassicAssert.AreEqual(numCF + numCF12 + numCFEX, sheetCF.NumConditionalFormattings);

            int fCF = 0, fCF12 = 0, fCFEX = 0;
            for (int i = 0; i < sheetCF.NumConditionalFormattings; i++)
            {
                IConditionalFormatting cf0 = sheetCF.GetConditionalFormattingAt(i);
                if (cf0 is HSSFConditionalFormatting)
                {
                    String str = cf0.ToString();
                    if (str.Contains("[CF]"))
                        fCF++;
                    if (str.Contains("[CF12]"))
                        fCF12++;
                    if (str.Contains("[CFEX]"))
                        fCFEX++;
                }
                else
                {
                    ConditionType type = cf0.GetRule(cf0.NumberOfRules - 1).ConditionType;
                    if (type == ConditionType.CellValueIs ||
                        type == ConditionType.Formula)
                    {
                        fCF++;
                    }
                    else
                    {
                        // TODO Properly detect Ext ones from the xml
                        fCF12++;
                    }
                }
            }
            ClassicAssert.AreEqual(numCF, fCF);
            ClassicAssert.AreEqual(numCF12, fCF12);
            ClassicAssert.AreEqual(numCFEX, fCFEX);


            // Check the rules / values in detail

            // Highlight Positive values - Column C
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("C2:C17", cf.GetFormattingRanges()[0].FormatAsString());

            ClassicAssert.AreEqual(1, cf.NumberOfRules);
            IConditionalFormattingRule cr = cf.GetRule(0);
            ClassicAssert.AreEqual(ConditionType.CellValueIs, cr.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.GreaterThan, cr.ComparisonOperation);
            ClassicAssert.AreEqual("0", cr.Formula1);
            ClassicAssert.AreEqual(null, cr.Formula2);
            // When it matches:
            //   Sets the font colour to dark green
            //   Sets the background colour to lighter green
            // TODO Should the colours be slightly different between formats?
            if (cr is HSSFConditionalFormattingRule)
            {
                AssertColour("0:8080:0", cr.FontFormatting.FontColor);
                AssertColour("CCCC:FFFF:CCCC", cr.PatternFormatting.FillBackgroundColorColor);
            }
            else
            {
                AssertColour("006100", cr.FontFormatting.FontColor);
                AssertColour("C6EFCE", cr.PatternFormatting.FillBackgroundColorColor);
            }

            // Highlight 10-30 - Column D
            cf = sheetCF.GetConditionalFormattingAt(1);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("D2:D17", cf.GetFormattingRanges()[0].FormatAsString());

            ClassicAssert.AreEqual(1, cf.NumberOfRules);
            cr = cf.GetRule(0);
            ClassicAssert.AreEqual(ConditionType.CellValueIs, cr.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.Between, cr.ComparisonOperation);
            ClassicAssert.AreEqual("10", cr.Formula1);
            ClassicAssert.AreEqual("30", cr.Formula2);
            // When it matches:
            //   Sets the font colour to dark red
            //   Sets the background colour to lighter red
            // TODO Should the colours be slightly different between formats?
            if (cr is HSSFConditionalFormattingRule)
            {
                AssertColour("8080:0:8080", cr.FontFormatting.FontColor);
                AssertColour("FFFF:9999:CCCC", cr.PatternFormatting.FillBackgroundColorColor);
            }
            else
            {
                AssertColour("9C0006", cr.FontFormatting.FontColor);
                AssertColour("FFC7CE", cr.PatternFormatting.FillBackgroundColorColor);
            }

            // Data Bars - Column E
            cf = sheetCF.GetConditionalFormattingAt(2);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("E2:E17", cf.GetFormattingRanges()[0].FormatAsString());
            assertDataBar(cf, "FF63C384");


            // Colours Red->Yellow->Green - Column F
            cf = sheetCF.GetConditionalFormattingAt(3);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("F2:F17", cf.GetFormattingRanges()[0].FormatAsString());
            assertColorScale(cf, "F8696B", "FFEB84", "63BE7B");


            // Colours Blue->White->Red - Column G
            cf = sheetCF.GetConditionalFormattingAt(4);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("G2:G17", cf.GetFormattingRanges()[0].FormatAsString());
            assertColorScale(cf, "5A8AC6", "FCFCFF", "F8696B");


            // TODO Simplify asserts

            // Icons : Default - Column H, percentage thresholds

            cf = sheetCF.GetConditionalFormattingAt(5);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("H2:H17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_TRAFFIC_LIGHTS, 0d, 33d, 67d);

            // Icons : 3 signs - Column I
            cf = sheetCF.GetConditionalFormattingAt(6);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("I2:I17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_SHAPES, 0d, 33d, 67d);


            // Icons : 3 traffic lights 2 - Column J
            cf = sheetCF.GetConditionalFormattingAt(7);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("J2:J17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_TRAFFIC_LIGHTS_BOX, 0d, 33d, 67d);

            // Icons : 4 traffic lights - Column K
            cf = sheetCF.GetConditionalFormattingAt(8);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("K2:K17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYRB_4_TRAFFIC_LIGHTS, 0d, 25d, 50d, 75d);

            // Icons : 3 symbols with backgrounds - Column L
            cf = sheetCF.GetConditionalFormattingAt(9);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("L2:L17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_SYMBOLS_CIRCLE, 0d, 33d, 67d);

            // Icons : 3 flags - Column M2 Only
            cf = sheetCF.GetConditionalFormattingAt(10);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("M2", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_FLAGS, 0d, 33d, 67d);
            // Icons : 3 flags - Column M (all)
            cf = sheetCF.GetConditionalFormattingAt(11);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("M2:M17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_FLAGS, 0d, 33d, 67d);

            // Icons : 3 symbols 2 (no background) - Column N
            cf = sheetCF.GetConditionalFormattingAt(12);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("N2:N17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_SYMBOLS, 0d, 33d, 67d);

            // Icons : 3 arrows - Column O
            cf = sheetCF.GetConditionalFormattingAt(13);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("O2:O17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_ARROW, 0d, 33d, 67d);

            // Icons : 5 arrows grey - Column P    
            cf = sheetCF.GetConditionalFormattingAt(14);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("P2:P17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GREY_5_ARROWS, 0d, 20d, 40d, 60d, 80d);

            // Icons : 3 stars (ext) - Column Q
            // TODO Support EXT formattings

            // Icons : 4 ratings - Column R
            cf = sheetCF.GetConditionalFormattingAt(15);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("R2:R17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.RATINGS_4, 0d, 25d, 50d, 75d);

            // Icons : 5 ratings - Column S
            cf = sheetCF.GetConditionalFormattingAt(16);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("S2:S17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.RATINGS_5, 0d, 20d, 40d, 60d, 80d);

            // Custom Icon+Format - Column T
            cf = sheetCF.GetConditionalFormattingAt(17);
            ClassicAssert.AreEqual(1, cf.GetFormattingRanges().Length);
            ClassicAssert.AreEqual("T2:T17", cf.GetFormattingRanges()[0].FormatAsString());

            // TODO Support IconSet + Other CFs with 2 rules
            //        ClassicAssert.AreEqual(2, cf.NumberOfRules);
            //        cr = cf.getRule(0);
            //        assertIconSetPercentages(cr, IconSet.GYR_3_TRAFFIC_LIGHTS_BOX, 0d, 33d, 67d);
            //        cr = cf.getRule(1);
            //        ClassicAssert.AreEqual(ConditionType.FORMULA, cr.ConditionType);
            //        ClassicAssert.AreEqual(ComparisonOperator.NO_COMPARISON, cr.ComparisonOperation);
            //        // TODO Why aren't these two the same between formats?
            //        if (cr instanceof HSSFConditionalFormattingRule) {
            //            ClassicAssert.AreEqual("MOD(ROW($T1),2)=1", cr.Formula1);
            //        } else {
            //            ClassicAssert.AreEqual("MOD(ROW($T2),2)=1", cr.Formula1);
            //        }
            //        ClassicAssert.AreEqual(null, cr.Formula2);


            // Mixed icons - Column U
            // TODO Support EXT formattings

            wb.Close();
        }

        private void assertDataBar(IConditionalFormatting cf, String color)
        {
            ClassicAssert.AreEqual(1, cf.NumberOfRules);
            IConditionalFormattingRule cr = cf.GetRule(0);
            assertDataBar(cr, color);
        }
        private void assertDataBar(IConditionalFormattingRule cr, String color)
        {
            ClassicAssert.AreEqual(ConditionType.DataBar, cr.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.NoComparison, cr.ComparisonOperation);
            ClassicAssert.AreEqual(null, cr.Formula1);
            ClassicAssert.AreEqual(null, cr.Formula2);

            IDataBarFormatting databar = cr.DataBarFormatting;
            ClassicAssert.IsNotNull(databar);
            ClassicAssert.AreEqual(false, databar.IsIconOnly);
            ClassicAssert.AreEqual(true, databar.IsLeftToRight);
            ClassicAssert.AreEqual(0, databar.WidthMin);
            ClassicAssert.AreEqual(100, databar.WidthMax);

            AssertColour(color, databar.Color);

            IConditionalFormattingThreshold th;
            th = databar.MinThreshold;
            ClassicAssert.AreEqual(RangeType.MIN, th.RangeType);
            ClassicAssert.AreEqual(null, th.Value);
            ClassicAssert.AreEqual(null, th.Formula);
            th = databar.MaxThreshold;
            ClassicAssert.AreEqual(RangeType.MAX, th.RangeType);
            ClassicAssert.AreEqual(null, th.Value);
            ClassicAssert.AreEqual(null, th.Formula);
        }


        private void assertIconSetPercentages(IConditionalFormatting cf, IconSet iconset, params double[] vals)
        {
            ClassicAssert.AreEqual(1, cf.NumberOfRules);
            IConditionalFormattingRule cr = cf.GetRule(0);
            assertIconSetPercentages(cr, iconset, vals);
        }
        private void assertIconSetPercentages(IConditionalFormattingRule cr, IconSet iconset, params double[] vals)
        {
            ClassicAssert.AreEqual(ConditionType.IconSet, cr.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.NoComparison, cr.ComparisonOperation);
            ClassicAssert.AreEqual(null, cr.Formula1);
            ClassicAssert.AreEqual(null, cr.Formula2);

            IIconMultiStateFormatting icon = cr.MultiStateFormatting;
            ClassicAssert.IsNotNull(icon);
            ClassicAssert.AreEqual(iconset, icon.IconSet);
            ClassicAssert.AreEqual(false, icon.IsIconOnly);
            ClassicAssert.AreEqual(false, icon.IsReversed);

            ClassicAssert.IsNotNull(icon.Thresholds);
            ClassicAssert.AreEqual(vals.Length, icon.Thresholds.Length);
            for (int i = 0; i < vals.Length; i++)
            {
                Double v = vals[i];
                IConditionalFormattingThreshold th = icon.Thresholds[i] as IConditionalFormattingThreshold;
                ClassicAssert.AreEqual(RangeType.PERCENT, th.RangeType);
                ClassicAssert.AreEqual(v, th.Value);
                ClassicAssert.AreEqual(null, th.Formula);
            }
        }

        private void assertColorScale(IConditionalFormatting cf, params string[] colors)
        {
            ClassicAssert.AreEqual(1, cf.NumberOfRules);
            IConditionalFormattingRule cr = cf.GetRule(0);
            assertColorScale(cr, colors);
        }
        private void assertColorScale(IConditionalFormattingRule cr, params string[] colors)
        {
            ClassicAssert.AreEqual(ConditionType.ColorScale, cr.ConditionType);
            ClassicAssert.AreEqual(ComparisonOperator.NoComparison, cr.ComparisonOperation);
            ClassicAssert.AreEqual(null, cr.Formula1);
            ClassicAssert.AreEqual(null, cr.Formula2);

            // TODO Implement
            if (cr is HSSFConditionalFormattingRule)
                return;
            IColorScaleFormatting color = cr.ColorScaleFormatting;
            ClassicAssert.IsNotNull(color);
            ClassicAssert.IsNotNull(color.Colors);
            ClassicAssert.IsNotNull(color.Thresholds);
            ClassicAssert.AreEqual(colors.Length, color.NumControlPoints);
            ClassicAssert.AreEqual(colors.Length, color.Colors.Length);
            ClassicAssert.AreEqual(colors.Length, color.Thresholds.Length);

            // Thresholds should be Min / (evenly spaced) / Max
            int steps = 100 / (colors.Length - 1);
            for (int i = 0; i < colors.Length; i++)
            {
                IConditionalFormattingThreshold th = color.Thresholds[i];
                if (i == 0)
                {
                    ClassicAssert.AreEqual(RangeType.MIN, th.RangeType);
                }
                else if (i == colors.Length - 1)
                {
                    ClassicAssert.AreEqual(RangeType.MAX, th.RangeType);
                }
                else
                {
                    ClassicAssert.AreEqual(RangeType.PERCENTILE, th.RangeType);
                    ClassicAssert.AreEqual(steps * i, (int)th.Value.Value);
                }
                ClassicAssert.AreEqual(null, th.Formula);
            }

            // Colors should match
            for (int i = 0; i < colors.Length; i++)
            {
                AssertColour(colors[i], color.Colors[i]);
            }

        }
        [Test]
        public void TestCreateFontFormatting()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.Equal, "7");
            IFontFormatting fontFmt = rule1.CreateFontFormatting();
            ClassicAssert.IsFalse(fontFmt.IsItalic);
            ClassicAssert.IsFalse(fontFmt.IsBold);
            fontFmt.SetFontStyle(true, true);
            ClassicAssert.IsTrue(fontFmt.IsItalic);
            ClassicAssert.IsTrue(fontFmt.IsBold);

            ClassicAssert.AreEqual(-1, fontFmt.FontHeight); // not modified
            fontFmt.FontHeight = (/*setter*/200);
            ClassicAssert.AreEqual(200, fontFmt.FontHeight);
            fontFmt.FontHeight = (/*setter*/100);
            ClassicAssert.AreEqual(100, fontFmt.FontHeight);

            ClassicAssert.AreEqual(FontSuperScript.None, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontSuperScript.Sub);
            ClassicAssert.AreEqual(FontSuperScript.Sub, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontSuperScript.None);
            ClassicAssert.AreEqual(FontSuperScript.None, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontSuperScript.Super);
            ClassicAssert.AreEqual(FontSuperScript.Super, fontFmt.EscapementType);

            ClassicAssert.AreEqual(FontUnderlineType.None, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontUnderlineType.Single);
            ClassicAssert.AreEqual(FontUnderlineType.Single, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontUnderlineType.None);
            ClassicAssert.AreEqual(FontUnderlineType.None, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontUnderlineType.Double);
            ClassicAssert.AreEqual(FontUnderlineType.Double, fontFmt.UnderlineType);

            ClassicAssert.AreEqual(-1, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/HSSFColor.Red.Index);
            ClassicAssert.AreEqual(HSSFColor.Red.Index, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/HSSFColor.Automatic.Index);
            ClassicAssert.AreEqual(HSSFColor.Automatic.Index, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/HSSFColor.Blue.Index);
            ClassicAssert.AreEqual(HSSFColor.Blue.Index, fontFmt.FontColorIndex);

            IConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.IsNotNull(cf);

            ClassicAssert.AreEqual(1, cf.NumberOfRules);

            IFontFormatting r1fp = cf.GetRule(0).FontFormatting;
            ClassicAssert.IsNotNull(r1fp);

            ClassicAssert.IsTrue(r1fp.IsItalic);
            ClassicAssert.IsTrue(r1fp.IsBold);
            ClassicAssert.AreEqual(FontSuperScript.Super, r1fp.EscapementType);
            ClassicAssert.AreEqual(FontUnderlineType.Double, r1fp.UnderlineType);
            ClassicAssert.AreEqual(HSSFColor.Blue.Index, r1fp.FontColorIndex);

            workbook.Close();
        }
        [Test]
        public void TestCreatePatternFormatting()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.Equal, "7");
            IPatternFormatting patternFmt = rule1.CreatePatternFormatting();

            ClassicAssert.AreEqual(0, patternFmt.FillBackgroundColor);
            patternFmt.FillBackgroundColor = (/*setter*/HSSFColor.Red.Index);
            ClassicAssert.AreEqual(HSSFColor.Red.Index, patternFmt.FillBackgroundColor);

            ClassicAssert.AreEqual(0, patternFmt.FillForegroundColor);
            patternFmt.FillForegroundColor = (/*setter*/HSSFColor.Blue.Index);
            ClassicAssert.AreEqual(HSSFColor.Blue.Index, patternFmt.FillForegroundColor);

            ClassicAssert.AreEqual(FillPattern.NoFill, patternFmt.FillPattern);
            patternFmt.FillPattern = FillPattern.SolidForeground;
            ClassicAssert.AreEqual(FillPattern.SolidForeground, patternFmt.FillPattern);
            patternFmt.FillPattern = FillPattern.NoFill;
            ClassicAssert.AreEqual(FillPattern.NoFill, patternFmt.FillPattern);
            if (this._testDataProvider.GetSpreadsheetVersion() == SpreadsheetVersion.EXCEL97)
            {
                patternFmt.FillPattern = FillPattern.Bricks;
                ClassicAssert.AreEqual(FillPattern.Bricks, patternFmt.FillPattern);
            }

            IConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.IsNotNull(cf);

            ClassicAssert.AreEqual(1, cf.NumberOfRules);

            IPatternFormatting r1fp = cf.GetRule(0).PatternFormatting;
            ClassicAssert.IsNotNull(r1fp);

            ClassicAssert.AreEqual(HSSFColor.Red.Index, r1fp.FillBackgroundColor);
            ClassicAssert.AreEqual(HSSFColor.Blue.Index, r1fp.FillForegroundColor);
            if (this._testDataProvider.GetSpreadsheetVersion() == SpreadsheetVersion.EXCEL97)
            {
                ClassicAssert.AreEqual(FillPattern.Bricks, r1fp.FillPattern);
            }

            workbook.Close();
        }

        [Test]
        public void TestAllCreateBorderFormatting()
        {
            // Make sure it is possible to create a conditional formatting rule
            // with every type of Border Style
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.Equal, "7");
            IBorderFormatting borderFmt = rule1.CreateBorderFormatting();
            foreach (BorderStyle border in BorderStyleEnum.Values())
            {
                borderFmt.BorderTop = border;
                ClassicAssert.AreEqual(border, borderFmt.BorderTop);
                borderFmt.BorderBottom = border;
                ClassicAssert.AreEqual(border, borderFmt.BorderBottom);
                borderFmt.BorderLeft = border;
                ClassicAssert.AreEqual(border, borderFmt.BorderLeft);
                borderFmt.BorderRight = border;
                ClassicAssert.AreEqual(border, borderFmt.BorderRight);
                borderFmt.BorderDiagonal = border;
                ClassicAssert.AreEqual(border, borderFmt.BorderDiagonal);
            }
            workbook.Close();
        }

        [Test]
        public void TestCreateBorderFormatting()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.Equal, "7");
            IBorderFormatting borderFmt = rule1.CreateBorderFormatting();

            ClassicAssert.AreEqual(BorderStyle.None, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderStyle.Dotted);
            ClassicAssert.AreEqual(BorderStyle.Dotted, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderStyle.None);
            ClassicAssert.AreEqual(BorderStyle.None, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderStyle.Thick);
            ClassicAssert.AreEqual(BorderStyle.Thick, borderFmt.BorderBottom);

            ClassicAssert.AreEqual(BorderStyle.None, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderStyle.Dotted);
            ClassicAssert.AreEqual(BorderStyle.Dotted, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderStyle.None);
            ClassicAssert.AreEqual(BorderStyle.None, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderStyle.Thick);
            ClassicAssert.AreEqual(BorderStyle.Thick, borderFmt.BorderTop);

            ClassicAssert.AreEqual(BorderStyle.None, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderStyle.Dotted);
            ClassicAssert.AreEqual(BorderStyle.Dotted, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderStyle.None);
            ClassicAssert.AreEqual(BorderStyle.None, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderStyle.Thin);
            ClassicAssert.AreEqual(BorderStyle.Thin, borderFmt.BorderLeft);

            ClassicAssert.AreEqual(BorderStyle.None, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderStyle.Dotted);
            ClassicAssert.AreEqual(BorderStyle.Dotted, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderStyle.None);
            ClassicAssert.AreEqual(BorderStyle.None, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderStyle.Hair);
            ClassicAssert.AreEqual(BorderStyle.Hair, borderFmt.BorderRight);

            IConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.IsNotNull(cf);

            ClassicAssert.AreEqual(1, cf.NumberOfRules);

            IBorderFormatting r1fp = cf.GetRule(0).BorderFormatting;
            ClassicAssert.IsNotNull(r1fp);
            ClassicAssert.AreEqual(BorderStyle.Thick, r1fp.BorderBottom);
            ClassicAssert.AreEqual(BorderStyle.Thick, r1fp.BorderTop);
            ClassicAssert.AreEqual(BorderStyle.Thin, r1fp.BorderLeft);
            ClassicAssert.AreEqual(BorderStyle.Hair, r1fp.BorderRight);

            workbook.Close();
        }

        [Test]
        public void TestCreateIconFormatting()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;
            IConditionalFormattingRule rule1 =
                    sheetCF.CreateConditionalFormattingRule(IconSet.GYRB_4_TRAFFIC_LIGHTS);
            IIconMultiStateFormatting iconFmt = rule1.MultiStateFormatting;

            ClassicAssert.AreEqual(IconSet.GYRB_4_TRAFFIC_LIGHTS, iconFmt.IconSet);
            ClassicAssert.AreEqual(4, iconFmt.Thresholds.Length);
            ClassicAssert.AreEqual(false, iconFmt.IsIconOnly);
            ClassicAssert.AreEqual(false, iconFmt.IsReversed);

            iconFmt.IsIconOnly = (true);
            iconFmt.Thresholds[0].RangeType = RangeType.MIN;
            iconFmt.Thresholds[1].RangeType = RangeType.NUMBER;
            iconFmt.Thresholds[1].Value = (10d);
            iconFmt.Thresholds[2].RangeType = RangeType.PERCENT;
            iconFmt.Thresholds[2].Value = (75d);
            iconFmt.Thresholds[3].RangeType = RangeType.MAX;

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };
            sheetCF.AddConditionalFormatting(regions, rule1);

            // Save, re-load and re-check
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);
            sheetCF = sheet.SheetConditionalFormatting;
            ClassicAssert.AreEqual(1, sheetCF.NumConditionalFormattings);

            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual(1, cf.NumberOfRules);
            rule1 = cf.GetRule(0);
            ClassicAssert.AreEqual(ConditionType.IconSet, rule1.ConditionType);
            iconFmt = rule1.MultiStateFormatting;

            ClassicAssert.AreEqual(IconSet.GYRB_4_TRAFFIC_LIGHTS, iconFmt.IconSet);
            ClassicAssert.AreEqual(4, iconFmt.Thresholds.Length);
            ClassicAssert.AreEqual(true, iconFmt.IsIconOnly);
            ClassicAssert.AreEqual(false, iconFmt.IsReversed);
            ClassicAssert.AreEqual(RangeType.MIN, iconFmt.Thresholds[0].RangeType);
            ClassicAssert.AreEqual(RangeType.NUMBER, iconFmt.Thresholds[1].RangeType);
            ClassicAssert.AreEqual(RangeType.PERCENT, iconFmt.Thresholds[2].RangeType);
            ClassicAssert.AreEqual(RangeType.MAX, iconFmt.Thresholds[3].RangeType);
            ClassicAssert.AreEqual(null, iconFmt.Thresholds[0].Value);
            ClassicAssert.AreEqual(10d, iconFmt.Thresholds[1].Value);
            ClassicAssert.AreEqual(75d, iconFmt.Thresholds[2].Value);
            ClassicAssert.AreEqual(null, iconFmt.Thresholds[3].Value);

            wb2.Close();
        }

        [Test]
        public void TestCreateColorScaleFormatting()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;
            IConditionalFormattingRule rule1 =
                    sheetCF.CreateConditionalFormattingColorScaleRule();
            IColorScaleFormatting clrFmt = rule1.ColorScaleFormatting;

            ClassicAssert.AreEqual(3, clrFmt.NumControlPoints);
            ClassicAssert.AreEqual(3, clrFmt.Colors.Length);
            ClassicAssert.AreEqual(3, clrFmt.Thresholds.Length);

            clrFmt.Thresholds[0].RangeType = (RangeType.MIN);
            clrFmt.Thresholds[1].RangeType = (RangeType.NUMBER);
            clrFmt.Thresholds[1].Value = (10d);
            clrFmt.Thresholds[2].RangeType = (RangeType.MAX);

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };
            sheetCF.AddConditionalFormatting(regions, rule1);

            // Save, re-load and re-check
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);
            sheetCF = sheet.SheetConditionalFormatting;
            ClassicAssert.AreEqual(1, sheetCF.NumConditionalFormattings);

            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual(1, cf.NumberOfRules);
            rule1 = cf.GetRule(0);
            clrFmt = rule1.ColorScaleFormatting;
            ClassicAssert.AreEqual(ConditionType.ColorScale, rule1.ConditionType);

            ClassicAssert.AreEqual(3, clrFmt.NumControlPoints);
            ClassicAssert.AreEqual(3, clrFmt.Colors.Length);
            ClassicAssert.AreEqual(3, clrFmt.Thresholds.Length);
            ClassicAssert.AreEqual(RangeType.MIN, clrFmt.Thresholds[0].RangeType);
            ClassicAssert.AreEqual(RangeType.NUMBER, clrFmt.Thresholds[1].RangeType);
            ClassicAssert.AreEqual(RangeType.MAX, clrFmt.Thresholds[2].RangeType);
            ClassicAssert.AreEqual(null, clrFmt.Thresholds[0].Value);
            ClassicAssert.AreEqual(10d, clrFmt.Thresholds[1].Value);
            ClassicAssert.AreEqual(null, clrFmt.Thresholds[2].Value);

            wb2.Close();
        }
        [Test]
        public void TestCreateDataBarFormatting()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb1.CreateSheet();
            String colorHex = "FFFFEB84";
            ExtendedColor color = wb1.GetCreationHelper().CreateExtendedColor();
            color.ARGBHex = (colorHex);
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;
            IConditionalFormattingRule rule1 =
                    sheetCF.CreateConditionalFormattingRule(color);
            IDataBarFormatting dbFmt = rule1.DataBarFormatting;

            ClassicAssert.AreEqual(false, dbFmt.IsIconOnly);
            ClassicAssert.AreEqual(true, dbFmt.IsLeftToRight);
            ClassicAssert.AreEqual(0, dbFmt.WidthMin);
            ClassicAssert.AreEqual(100, dbFmt.WidthMax);
            AssertColour(colorHex, dbFmt.Color);

            dbFmt.MinThreshold.RangeType = (RangeType.MIN);
            dbFmt.MaxThreshold.RangeType = (RangeType.MAX);

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };
            sheetCF.AddConditionalFormatting(regions, rule1);

            // Save, re-load and re-check
            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0);
            sheetCF = sheet.SheetConditionalFormatting;
            ClassicAssert.AreEqual(1, sheetCF.NumConditionalFormattings);

            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            ClassicAssert.AreEqual(1, cf.NumberOfRules);
            rule1 = cf.GetRule(0);
            dbFmt = rule1.DataBarFormatting;
            ClassicAssert.AreEqual(ConditionType.DataBar, rule1.ConditionType);

            ClassicAssert.AreEqual(false, dbFmt.IsIconOnly);
            ClassicAssert.AreEqual(true, dbFmt.IsLeftToRight);
            ClassicAssert.AreEqual(0, dbFmt.WidthMin);
            ClassicAssert.AreEqual(100, dbFmt.WidthMax);
            AssertColour(colorHex, dbFmt.Color);
            ClassicAssert.AreEqual(RangeType.MIN, dbFmt.MinThreshold.RangeType);
            ClassicAssert.AreEqual(RangeType.MAX, dbFmt.MaxThreshold.RangeType);
            ClassicAssert.AreEqual(null, dbFmt.MinThreshold.Value);
            ClassicAssert.AreEqual(null, dbFmt.MaxThreshold.Value);

            wb2.Close();
        }

        [Test]
        public void TestBug55380()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            CellRangeAddress[] ranges = new CellRangeAddress[] {
                CellRangeAddress.ValueOf("C9:D30"), CellRangeAddress.ValueOf("C7:C31")
            };
            IConditionalFormattingRule rule = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule("$A$1>0");
            sheet.SheetConditionalFormatting.AddConditionalFormatting(ranges, rule);

            wb.Close();
        }

        [Test]
        public void TestSetCellRangeAddresswithSingleRange()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("S1");
            ISheetConditionalFormatting cf = sheet.SheetConditionalFormatting;
            ClassicAssert.AreEqual(0, cf.NumConditionalFormattings);
            IConditionalFormattingRule rule1 = cf.CreateConditionalFormattingRule("$A$1>0");
            cf.AddConditionalFormatting(new CellRangeAddress[] {
                CellRangeAddress.ValueOf("A1:A5")
            }, rule1);

            ClassicAssert.AreEqual(1, cf.NumConditionalFormattings);
            IConditionalFormatting ReadCf = cf.GetConditionalFormattingAt(0);
            CellRangeAddress[] formattingRanges = ReadCf.GetFormattingRanges();
            ClassicAssert.AreEqual(1, formattingRanges.Length);
            CellRangeAddress formattingRange = formattingRanges[0];
            ClassicAssert.AreEqual("A1:A5", formattingRange.FormatAsString());

            ReadCf.SetFormattingRanges(new CellRangeAddress[] {
                CellRangeAddress.ValueOf("A1:A6")
            });

            ReadCf = cf.GetConditionalFormattingAt(0);
            formattingRanges = ReadCf.GetFormattingRanges();
            ClassicAssert.AreEqual(1, formattingRanges.Length);
            formattingRange = formattingRanges[0];
            ClassicAssert.AreEqual("A1:A6", formattingRange.FormatAsString());
        }

        [Test]
        public void TestSetCellRangeAddressWithMultipleRanges()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("S1");
            ISheetConditionalFormatting cf = sheet.SheetConditionalFormatting;
            ClassicAssert.AreEqual(0, cf.NumConditionalFormattings);
            IConditionalFormattingRule rule1 = cf.CreateConditionalFormattingRule("$A$1>0");
            cf.AddConditionalFormatting(new CellRangeAddress[] {
                CellRangeAddress.ValueOf("A1:A5")
            }, rule1);

            ClassicAssert.AreEqual(1, cf.NumConditionalFormattings);
            IConditionalFormatting ReadCf = cf.GetConditionalFormattingAt(0);
            CellRangeAddress[] formattingRanges = ReadCf.GetFormattingRanges();
            ClassicAssert.AreEqual(1, formattingRanges.Length);
            CellRangeAddress formattingRange = formattingRanges[0];
            ClassicAssert.AreEqual("A1:A5", formattingRange.FormatAsString());

            ReadCf.SetFormattingRanges(new CellRangeAddress[] {
                CellRangeAddress.ValueOf("A1:A6"),
                CellRangeAddress.ValueOf("B1:B6")
            });

            ReadCf = cf.GetConditionalFormattingAt(0);
            formattingRanges = ReadCf.GetFormattingRanges();
            ClassicAssert.AreEqual(2, formattingRanges.Length);
            formattingRange = formattingRanges[0];
            ClassicAssert.AreEqual("A1:A6", formattingRange.FormatAsString());
            formattingRange = formattingRanges[1];
            ClassicAssert.AreEqual("B1:B6", formattingRange.FormatAsString());
        }

        [Test]
        public void TestSetCellRangeAddressWithNullRanges()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("S1");
            ISheetConditionalFormatting cf = sheet.SheetConditionalFormatting;
            ClassicAssert.AreEqual(0, cf.NumConditionalFormattings);
            IConditionalFormattingRule rule1 = cf.CreateConditionalFormattingRule("$A$1>0");
            cf.AddConditionalFormatting(new CellRangeAddress[] {
                CellRangeAddress.ValueOf("A1:A5")
            }, rule1);

            ClassicAssert.AreEqual(1, cf.NumConditionalFormattings);
            IConditionalFormatting ReadCf = cf.GetConditionalFormattingAt(0);
            Assert.Throws<ArgumentNullException>(() => ReadCf.SetFormattingRanges(null));
        }

        [Test]
        public void Test52122()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("Conditional Formatting Test");
            sheet.SetColumnWidth(0, 256 * 10);
            sheet.SetColumnWidth(1, 256 * 10);
            sheet.SetColumnWidth(2, 256 * 10);
            // Create some content.
            // row 0
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            cell0.SetCellType(CellType.Numeric);
            cell0.SetCellValue(100);
            ICell cell1 = row.CreateCell(1);
            cell1.SetCellType(CellType.Numeric);
            cell1.SetCellValue(120);
            ICell cell2 = row.CreateCell(2);
            cell2.SetCellType(CellType.Numeric);
            cell2.SetCellValue(130);
            // row 1
            row = sheet.CreateRow(1);
            cell0 = row.CreateCell(0);
            cell0.SetCellType(CellType.Numeric);
            cell0.SetCellValue(200);
            cell1 = row.CreateCell(1);
            cell1.SetCellType(CellType.Numeric);
            cell1.SetCellValue(220);
            cell2 = row.CreateCell(2);
            cell2.SetCellType(CellType.Numeric);
            cell2.SetCellValue(230);
            // row 2
            row = sheet.CreateRow(2);
            cell0 = row.CreateCell(0);
            cell0.SetCellType(CellType.Numeric);
            cell0.SetCellValue(300);
            cell1 = row.CreateCell(1);
            cell1.SetCellType(CellType.Numeric);
            cell1.SetCellValue(320);
            cell2 = row.CreateCell(2);
            cell2.SetCellType(CellType.Numeric);
            cell2.SetCellValue(330);
            // Create conditional formatting, CELL1 should be yellow if CELL0 is not blank.
            ISheetConditionalFormatting formatting = sheet.SheetConditionalFormatting;
            IConditionalFormattingRule rule = formatting.CreateConditionalFormattingRule("$A$1>75");
            IPatternFormatting pattern = rule.CreatePatternFormatting();
            pattern.FillBackgroundColor = IndexedColors.Blue.Index;
            pattern.FillPattern = FillPattern.SolidForeground;
            CellRangeAddress[] range = { CellRangeAddress.ValueOf("B2:C2") };
            CellRangeAddress[] range2 = { CellRangeAddress.ValueOf("B1:C1") };
            formatting.AddConditionalFormatting(range, rule);
            formatting.AddConditionalFormatting(range2, rule);
            // Write file.
            /*FileOutputStream fos = new FileOutputStream("c:\\temp\\52122_conditional-sheet.xls");
            try {
                workbook.write(fos);
            } finally {
                fos.Close();
            }*/
            IWorkbook wbBack = _testDataProvider.WriteOutAndReadBack(workbook);
            ISheet sheetBack = wbBack.GetSheetAt(0);
            ISheetConditionalFormatting sheetConditionalFormattingBack = sheetBack.SheetConditionalFormatting;
            ClassicAssert.IsNotNull(sheetConditionalFormattingBack);
            IConditionalFormatting formattingBack = sheetConditionalFormattingBack.GetConditionalFormattingAt(0);
            ClassicAssert.IsNotNull(formattingBack);
            IConditionalFormattingRule ruleBack = formattingBack.GetRule(0);
            ClassicAssert.IsNotNull(ruleBack);
            IPatternFormatting patternFormattingBack1 = ruleBack.PatternFormatting;
            ClassicAssert.IsNotNull(patternFormattingBack1);
            ClassicAssert.AreEqual(IndexedColors.Blue.Index, patternFormattingBack1.FillBackgroundColor);
            ClassicAssert.AreEqual(FillPattern.SolidForeground, patternFormattingBack1.FillPattern);
        }
    }

}