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

    using NUnit.Framework;
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

            Assert.AreEqual(0, sheetCF.NumConditionalFormattings);
            try
            {
                Assert.IsNull(sheetCF.GetConditionalFormattingAt(0));
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.StartsWith("Specified CF index 0 is outside the allowable range"));
            }

            try
            {
                sheetCF.RemoveConditionalFormatting(0);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.StartsWith("Specified CF index 0 is outside the allowable range"));
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
                Assert.IsTrue(e.Message.StartsWith("regions must not be null"));
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
                Assert.IsTrue(e.Message.StartsWith("cfRules must not be null"));
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
                Assert.IsTrue(e.Message.StartsWith("cfRules must not be empty"));
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
            Assert.AreEqual(ConditionType.Formula, rule1.ConditionType);
            Assert.AreEqual("SUM(A1:A5)>10", rule1.Formula1);
            int formatIndex1 = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("B1"),
                        CellRangeAddress.ValueOf("C3"),
                }, rule1);
            Assert.AreEqual(0, formatIndex1);
            Assert.AreEqual(1, sheetCF.NumConditionalFormattings);
            CellRangeAddress[] ranges1 = sheetCF.GetConditionalFormattingAt(formatIndex1).GetFormattingRanges();
            Assert.AreEqual(2, ranges1.Length);
            Assert.AreEqual("B1", ranges1[0].FormatAsString());
            Assert.AreEqual("C3", ranges1[1].FormatAsString());

            // adjacent Address are merged
            int formatIndex2 = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("B1"),
                        CellRangeAddress.ValueOf("B2"),
                        CellRangeAddress.ValueOf("B3"),
                }, rule1);
            Assert.AreEqual(1, formatIndex2);
            Assert.AreEqual(2, sheetCF.NumConditionalFormattings);
            CellRangeAddress[] ranges2 = sheetCF.GetConditionalFormattingAt(formatIndex2).GetFormattingRanges();
            Assert.AreEqual(1, ranges2.Length);
            Assert.AreEqual("B1:B3", ranges2[0].FormatAsString());

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
            Assert.AreEqual(ConditionType.CellValueIs, rule1.ConditionType);
            Assert.AreEqual("SUM(A1:A5)+10", rule1.Formula1);
            Assert.AreEqual(ComparisonOperator.Equal, rule1.ComparisonOperation);

            IConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NotEqual, "15");
            Assert.AreEqual(ConditionType.CellValueIs, rule2.ConditionType);
            Assert.AreEqual("15", rule2.Formula1);
            Assert.AreEqual(ComparisonOperator.NotEqual, rule2.ComparisonOperation);

            IConditionalFormattingRule rule3 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NotEqual, "15");
            Assert.AreEqual(ConditionType.CellValueIs, rule3.ConditionType);
            Assert.AreEqual("15", rule3.Formula1);
            Assert.AreEqual(ComparisonOperator.NotEqual, rule3.ComparisonOperation);

            IConditionalFormattingRule rule4 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.GreaterThan, "0");
            Assert.AreEqual(ConditionType.CellValueIs, rule4.ConditionType);
            Assert.AreEqual("0", rule4.Formula1);
            Assert.AreEqual(ComparisonOperator.GreaterThan, rule4.ComparisonOperation);

            IConditionalFormattingRule rule5 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.LessThan, "0");
            Assert.AreEqual(ConditionType.CellValueIs, rule5.ConditionType);
            Assert.AreEqual("0", rule5.Formula1);
            Assert.AreEqual(ComparisonOperator.LessThan, rule5.ComparisonOperation);

            IConditionalFormattingRule rule6 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.GreaterThanOrEqual, "0");
            Assert.AreEqual(ConditionType.CellValueIs, rule6.ConditionType);
            Assert.AreEqual("0", rule6.Formula1);
            Assert.AreEqual(ComparisonOperator.GreaterThanOrEqual, rule6.ComparisonOperation);

            IConditionalFormattingRule rule7 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.LessThanOrEqual, "0");
            Assert.AreEqual(ConditionType.CellValueIs, rule7.ConditionType);
            Assert.AreEqual("0", rule7.Formula1);
            Assert.AreEqual(ComparisonOperator.LessThanOrEqual, rule7.ComparisonOperation);

            IConditionalFormattingRule rule8 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.Between, "0", "5");
            Assert.AreEqual(ConditionType.CellValueIs, rule8.ConditionType);
            Assert.AreEqual("0", rule8.Formula1);
            Assert.AreEqual("5", rule8.Formula2);
            Assert.AreEqual(ComparisonOperator.Between, rule8.ComparisonOperation);

            IConditionalFormattingRule rule9 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NotBetween, "0", "5");
            Assert.AreEqual(ConditionType.CellValueIs, rule9.ConditionType);
            Assert.AreEqual("0", rule9.Formula1);
            Assert.AreEqual("5", rule9.Formula2);
            Assert.AreEqual(ComparisonOperator.NotBetween, rule9.ComparisonOperation);

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
            Assert.AreEqual(0, sheet1CF.NumConditionalFormattings);
            Assert.AreEqual(0, sheet2CF.NumConditionalFormattings);

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
            Assert.AreEqual(0, formatIndex);
            Assert.AreEqual(1, sheet1CF.NumConditionalFormattings);

            Assert.AreEqual(0, sheet2CF.NumConditionalFormattings);
            sheet2CF.AddConditionalFormatting(sheet1CF.GetConditionalFormattingAt(formatIndex));
            Assert.AreEqual(1, sheet2CF.NumConditionalFormattings);

            IConditionalFormatting sheet2cf = sheet2CF.GetConditionalFormattingAt(0);
            Assert.AreEqual(2, sheet2cf.NumberOfRules);
            Assert.AreEqual("SUM(A1:A5)+10", sheet2cf.GetRule(0).Formula1);
            Assert.AreEqual(ComparisonOperator.Equal, sheet2cf.GetRule(0).ComparisonOperation);
            Assert.AreEqual(ConditionType.CellValueIs, sheet2cf.GetRule(0).ConditionType);
            Assert.AreEqual("15", sheet2cf.GetRule(1).Formula1);
            Assert.AreEqual(ComparisonOperator.NotEqual, sheet2cf.GetRule(1).ComparisonOperation);
            Assert.AreEqual(ConditionType.CellValueIs, sheet2cf.GetRule(1).ConditionType);

            wb.Close();
        }
        [Test]
        public void TestRemove()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            ISheetConditionalFormatting sheetCF = sheet1.SheetConditionalFormatting;
            Assert.AreEqual(0, sheetCF.NumConditionalFormattings);

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.Equal, "SUM(A1:A5)");

            // adjacent Address are merged
            int formatIndex = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("A1:A5")
                }, rule1);
            Assert.AreEqual(0, formatIndex);
            Assert.AreEqual(1, sheetCF.NumConditionalFormattings);
            sheetCF.RemoveConditionalFormatting(0);
            Assert.AreEqual(0, sheetCF.NumConditionalFormattings);
            try
            {
                Assert.IsNull(sheetCF.GetConditionalFormattingAt(0));
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.StartsWith("Specified CF index 0 is outside the allowable range"));
            }

            formatIndex = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("A1:A5")
                }, rule1);
            Assert.AreEqual(0, formatIndex);
            Assert.AreEqual(1, sheetCF.NumConditionalFormattings);
            sheetCF.RemoveConditionalFormatting(0);
            Assert.AreEqual(0, sheetCF.NumConditionalFormattings);
            try
            {
                Assert.IsNull(sheetCF.GetConditionalFormattingAt(0));
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.StartsWith("Specified CF index 0 is outside the allowable range"));
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
            Assert.AreEqual(2, sheetCF.NumConditionalFormattings);
            sheetCF.RemoveConditionalFormatting(1);
            Assert.AreEqual(1, sheetCF.NumConditionalFormattings);
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.IsNotNull(cf);

            regions = cf.GetFormattingRanges();
            Assert.IsNotNull(regions);
            Assert.AreEqual(1, regions.Length);
            CellRangeAddress r = regions[0];
            Assert.AreEqual(1, r.FirstColumn);
            Assert.AreEqual(1, r.LastColumn);
            Assert.AreEqual(0, r.FirstRow);
            Assert.AreEqual(65535, r.LastRow);

            Assert.AreEqual(2, cf.NumberOfRules);

            rule1 = cf.GetRule(0);
            Assert.AreEqual("7", rule1.Formula1);
            Assert.IsNull(rule1.Formula2);

            IFontFormatting r1fp = rule1.FontFormatting;
            Assert.IsNotNull(r1fp);

            Assert.IsTrue(r1fp.IsItalic);
            Assert.IsFalse(r1fp.IsBold);

            IBorderFormatting r1bf = rule1.BorderFormatting;
            Assert.IsNotNull(r1bf);
            Assert.AreEqual(BorderStyle.Thin, r1bf.BorderBottom);
            Assert.AreEqual(BorderStyle.Thick, r1bf.BorderTop);
            Assert.AreEqual(BorderStyle.Dashed, r1bf.BorderLeft);
            Assert.AreEqual(BorderStyle.Dotted, r1bf.BorderRight);

            IPatternFormatting r1pf = rule1.PatternFormatting;
            Assert.IsNotNull(r1pf);
            //        Assert.AreEqual(HSSFColor.Yellow.index,r1pf.FillBackgroundColor);

            rule2 = cf.GetRule(1);
            Assert.AreEqual("2", rule2.Formula2);
            Assert.AreEqual("1", rule2.Formula1);

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
                Assert.AreEqual(2, wb.NumberOfSheets);
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf("needs to define a clone method") > 0)
                {
                    Assert.Fail("Indentified bug 45682");
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
            Assert.AreEqual(0, sheetCF.NumConditionalFormattings);

            // re-add the CF
            sheetCF.AddConditionalFormatting(regions, rule1);
            sheetCF.AddConditionalFormatting(regions, rule2);

            // This row shift should only affect the formulas
            sheet.ShiftRows(14, 17, 8);
            IConditionalFormatting cf1 = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual("SUM(A10:A23)", cf1.GetRule(0).Formula1);
            Assert.AreEqual("1+SUM(B24:B30)", cf1.GetRule(0).Formula2);
            IConditionalFormatting cf2 = sheetCF.GetConditionalFormattingAt(1);
            Assert.AreEqual("SUM(A10:A23)", cf2.GetRule(0).Formula1);
            Assert.AreEqual("1+SUM(B24:B30)", cf2.GetRule(0).Formula2);

            sheet.ShiftRows(0, 8, 21);
            cf1 = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual("SUM(A10:A21)", cf1.GetRule(0).Formula1);
            Assert.AreEqual("1+SUM(#REF!)", cf1.GetRule(0).Formula2);
            cf2 = sheetCF.GetConditionalFormattingAt(1);
            Assert.AreEqual("SUM(A10:A21)", cf2.GetRule(0).Formula1);
            Assert.AreEqual("1+SUM(#REF!)", cf2.GetRule(0).Formula2);

            wb.Close();
        }
        //

        protected void TestRead(string sampleFile)
        {

            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(sampleFile);
            ISheet sh = wb.GetSheet("CF");
            ISheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;
            Assert.AreEqual(3, sheetCF.NumConditionalFormattings);

            IConditionalFormatting cf1 = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual(2, cf1.NumberOfRules);

            CellRangeAddress[] regions1 = cf1.GetFormattingRanges();
            Assert.AreEqual(1, regions1.Length);
            Assert.AreEqual("A1:A8", regions1[0].FormatAsString());

            // CF1 has two rules: values less than -3 are bold-italic red, values greater than 3 are green
            IConditionalFormattingRule rule1 = cf1.GetRule(0);
            Assert.AreEqual(ConditionType.CellValueIs, rule1.ConditionType);
            Assert.AreEqual(ComparisonOperator.GreaterThan, rule1.ComparisonOperation);
            Assert.AreEqual("3", rule1.Formula1);
            Assert.IsNull(rule1.Formula2);
            // Fills and borders are not Set
            Assert.IsNull(rule1.PatternFormatting);
            Assert.IsNull(rule1.BorderFormatting);

            IFontFormatting fmt1 = rule1.FontFormatting;
            //        Assert.AreEqual(HSSFColor.GREEN.index, fmt1.FontColorIndex);
            Assert.IsTrue(fmt1.IsBold);
            Assert.IsFalse(fmt1.IsItalic);

            IConditionalFormattingRule rule2 = cf1.GetRule(1);
            Assert.AreEqual(ConditionType.CellValueIs, rule2.ConditionType);
            Assert.AreEqual(ComparisonOperator.LessThan, rule2.ComparisonOperation);
            Assert.AreEqual("-3", rule2.Formula1);
            Assert.IsNull(rule2.Formula2);
            Assert.IsNull(rule2.PatternFormatting);
            Assert.IsNull(rule2.BorderFormatting);

            IFontFormatting fmt2 = rule2.FontFormatting;
            //        Assert.AreEqual(HSSFColor.Red.index, fmt2.FontColorIndex);
            Assert.IsTrue(fmt2.IsBold);
            Assert.IsTrue(fmt2.IsItalic);


            IConditionalFormatting cf2 = sheetCF.GetConditionalFormattingAt(1);
            Assert.AreEqual(1, cf2.NumberOfRules);
            CellRangeAddress[] regions2 = cf2.GetFormattingRanges();
            Assert.AreEqual(1, regions2.Length);
            Assert.AreEqual("B9", regions2[0].FormatAsString());

            IConditionalFormattingRule rule3 = cf2.GetRule(0);
            Assert.AreEqual(ConditionType.Formula, rule3.ConditionType);
            Assert.AreEqual(ComparisonOperator.NoComparison, rule3.ComparisonOperation);
            Assert.AreEqual("$A$8>5", rule3.Formula1);
            Assert.IsNull(rule3.Formula2);

            IFontFormatting fmt3 = rule3.FontFormatting;
            //        Assert.AreEqual(HSSFColor.Red.index, fmt3.FontColorIndex);
            Assert.IsTrue(fmt3.IsBold);
            Assert.IsTrue(fmt3.IsItalic);

            IPatternFormatting fmt4 = rule3.PatternFormatting;
            //        Assert.AreEqual(HSSFColor.LIGHT_CORNFLOWER_BLUE.index, fmt4.FillBackgroundColor);
            //        Assert.AreEqual(HSSFColor.Automatic.index, fmt4.FillForegroundColor);
            Assert.AreEqual(FillPattern.NoFill, fmt4.FillPattern);
            // borders are not Set
            Assert.IsNull(rule3.BorderFormatting);

            IConditionalFormatting cf3 = sheetCF.GetConditionalFormattingAt(2);
            CellRangeAddress[] regions3 = cf3.GetFormattingRanges();
            Assert.AreEqual(1, regions3.Length);
            Assert.AreEqual("B1:B7", regions3[0].FormatAsString());
            Assert.AreEqual(2, cf3.NumberOfRules);

            IConditionalFormattingRule rule4 = cf3.GetRule(0);
            Assert.AreEqual(ConditionType.CellValueIs, rule4.ConditionType);
            Assert.AreEqual(ComparisonOperator.LessThanOrEqual, rule4.ComparisonOperation);
            Assert.AreEqual("\"AAA\"", rule4.Formula1);
            Assert.IsNull(rule4.Formula2);

            IConditionalFormattingRule rule5 = cf3.GetRule(1);
            Assert.AreEqual(ConditionType.CellValueIs, rule5.ConditionType);
            Assert.AreEqual(ComparisonOperator.Between, rule5.ComparisonOperation);
            Assert.AreEqual("\"A\"", rule5.Formula1);
            Assert.AreEqual("\"AAA\"", rule5.Formula2);

            wb.Close();
        }

        public void testReadOffice2007(String filename)
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(filename);
            ISheet s = wb.GetSheet("CF");
            IConditionalFormatting cf = null;
            IConditionalFormattingRule cr = null;

            // Sanity check data
            Assert.AreEqual("Values", s.GetRow(0).GetCell(0).ToString());
            Assert.AreEqual("10", s.GetRow(2).GetCell(0).ToString());
            //Assert.AreEqual("10.0", s.GetRow(2).GetCell(0).ToString());

            // Check we found all the conditional formattings rules we should have
            ISheetConditionalFormatting sheetCF = s.SheetConditionalFormatting;
            int numCF = 3;
            int numCF12 = 15;
            int numCFEX = 0; // TODO This should be 1, but we don't support CFEX formattings yet
            Assert.AreEqual(numCF + numCF12 + numCFEX, sheetCF.NumConditionalFormattings);

            int fCF = 0, fCF12 = 0, fCFEX = 0;
            for (int i = 0; i < sheetCF.NumConditionalFormattings; i++)
            {
                cf = sheetCF.GetConditionalFormattingAt(i);
                if (cf is HSSFConditionalFormatting)
                {
                    String str = cf.ToString();
                    if (str.Contains("[CF]"))
                        fCF++;
                    if (str.Contains("[CF12]"))
                        fCF12++;
                    if (str.Contains("[CFEX]"))
                        fCFEX++;
                }
                else
                {
                    ConditionType type = cf.GetRule(cf.NumberOfRules - 1).ConditionType;
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
            Assert.AreEqual(numCF, fCF);
            Assert.AreEqual(numCF12, fCF12);
            Assert.AreEqual(numCFEX, fCFEX);


            // Check the rules / values in detail

            // Highlight Positive values - Column C
            cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("C2:C17", cf.GetFormattingRanges()[0].FormatAsString());

            Assert.AreEqual(1, cf.NumberOfRules);
            cr = cf.GetRule(0);
            Assert.AreEqual(ConditionType.CellValueIs, cr.ConditionType);
            Assert.AreEqual(ComparisonOperator.GreaterThan, cr.ComparisonOperation);
            Assert.AreEqual("0", cr.Formula1);
            Assert.AreEqual(null, cr.Formula2);
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
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("D2:D17", cf.GetFormattingRanges()[0].FormatAsString());

            Assert.AreEqual(1, cf.NumberOfRules);
            cr = cf.GetRule(0);
            Assert.AreEqual(ConditionType.CellValueIs, cr.ConditionType);
            Assert.AreEqual(ComparisonOperator.Between, cr.ComparisonOperation);
            Assert.AreEqual("10", cr.Formula1);
            Assert.AreEqual("30", cr.Formula2);
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
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("E2:E17", cf.GetFormattingRanges()[0].FormatAsString());
            assertDataBar(cf, "FF63C384");


            // Colours Red->Yellow->Green - Column F
            cf = sheetCF.GetConditionalFormattingAt(3);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("F2:F17", cf.GetFormattingRanges()[0].FormatAsString());
            assertColorScale(cf, "F8696B", "FFEB84", "63BE7B");


            // Colours Blue->White->Red - Column G
            cf = sheetCF.GetConditionalFormattingAt(4);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("G2:G17", cf.GetFormattingRanges()[0].FormatAsString());
            assertColorScale(cf, "5A8AC6", "FCFCFF", "F8696B");


            // TODO Simplify asserts

            // Icons : Default - Column H, percentage thresholds

            cf = sheetCF.GetConditionalFormattingAt(5);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("H2:H17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_TRAFFIC_LIGHTS, 0d, 33d, 67d);

            // Icons : 3 signs - Column I
            cf = sheetCF.GetConditionalFormattingAt(6);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("I2:I17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_SHAPES, 0d, 33d, 67d);


            // Icons : 3 traffic lights 2 - Column J
            cf = sheetCF.GetConditionalFormattingAt(7);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("J2:J17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_TRAFFIC_LIGHTS_BOX, 0d, 33d, 67d);

            // Icons : 4 traffic lights - Column K
            cf = sheetCF.GetConditionalFormattingAt(8);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("K2:K17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYRB_4_TRAFFIC_LIGHTS, 0d, 25d, 50d, 75d);

            // Icons : 3 symbols with backgrounds - Column L
            cf = sheetCF.GetConditionalFormattingAt(9);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("L2:L17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_SYMBOLS_CIRCLE, 0d, 33d, 67d);

            // Icons : 3 flags - Column M2 Only
            cf = sheetCF.GetConditionalFormattingAt(10);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("M2", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_FLAGS, 0d, 33d, 67d);
            // Icons : 3 flags - Column M (all)
            cf = sheetCF.GetConditionalFormattingAt(11);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("M2:M17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_FLAGS, 0d, 33d, 67d);

            // Icons : 3 symbols 2 (no background) - Column N
            cf = sheetCF.GetConditionalFormattingAt(12);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("N2:N17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_SYMBOLS, 0d, 33d, 67d);

            // Icons : 3 arrows - Column O
            cf = sheetCF.GetConditionalFormattingAt(13);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("O2:O17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GYR_3_ARROW, 0d, 33d, 67d);

            // Icons : 5 arrows grey - Column P    
            cf = sheetCF.GetConditionalFormattingAt(14);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("P2:P17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.GREY_5_ARROWS, 0d, 20d, 40d, 60d, 80d);

            // Icons : 3 stars (ext) - Column Q
            // TODO Support EXT formattings

            // Icons : 4 ratings - Column R
            cf = sheetCF.GetConditionalFormattingAt(15);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("R2:R17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.RATINGS_4, 0d, 25d, 50d, 75d);

            // Icons : 5 ratings - Column S
            cf = sheetCF.GetConditionalFormattingAt(16);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("S2:S17", cf.GetFormattingRanges()[0].FormatAsString());
            assertIconSetPercentages(cf, IconSet.RATINGS_5, 0d, 20d, 40d, 60d, 80d);

            // Custom Icon+Format - Column T
            cf = sheetCF.GetConditionalFormattingAt(17);
            Assert.AreEqual(1, cf.GetFormattingRanges().Length);
            Assert.AreEqual("T2:T17", cf.GetFormattingRanges()[0].FormatAsString());

            // TODO Support IconSet + Other CFs with 2 rules
            //        Assert.AreEqual(2, cf.NumberOfRules);
            //        cr = cf.getRule(0);
            //        assertIconSetPercentages(cr, IconSet.GYR_3_TRAFFIC_LIGHTS_BOX, 0d, 33d, 67d);
            //        cr = cf.getRule(1);
            //        Assert.AreEqual(ConditionType.FORMULA, cr.ConditionType);
            //        Assert.AreEqual(ComparisonOperator.NO_COMPARISON, cr.ComparisonOperation);
            //        // TODO Why aren't these two the same between formats?
            //        if (cr instanceof HSSFConditionalFormattingRule) {
            //            Assert.AreEqual("MOD(ROW($T1),2)=1", cr.Formula1);
            //        } else {
            //            Assert.AreEqual("MOD(ROW($T2),2)=1", cr.Formula1);
            //        }
            //        Assert.AreEqual(null, cr.Formula2);


            // Mixed icons - Column U
            // TODO Support EXT formattings

            wb.Close();
        }

        private void assertDataBar(IConditionalFormatting cf, String color)
        {
            Assert.AreEqual(1, cf.NumberOfRules);
            IConditionalFormattingRule cr = cf.GetRule(0);
            assertDataBar(cr, color);
        }
        private void assertDataBar(IConditionalFormattingRule cr, String color)
        {
            Assert.AreEqual(ConditionType.DataBar, cr.ConditionType);
            Assert.AreEqual(ComparisonOperator.NoComparison, cr.ComparisonOperation);
            Assert.AreEqual(null, cr.Formula1);
            Assert.AreEqual(null, cr.Formula2);

            IDataBarFormatting databar = cr.DataBarFormatting;
            Assert.IsNotNull(databar);
            Assert.AreEqual(false, databar.IsIconOnly);
            Assert.AreEqual(true, databar.IsLeftToRight);
            Assert.AreEqual(0, databar.WidthMin);
            Assert.AreEqual(100, databar.WidthMax);

            AssertColour(color, databar.Color);

            IConditionalFormattingThreshold th;
            th = databar.MinThreshold;
            Assert.AreEqual(RangeType.MIN, th.RangeType);
            Assert.AreEqual(null, th.Value);
            Assert.AreEqual(null, th.Formula);
            th = databar.MaxThreshold;
            Assert.AreEqual(RangeType.MAX, th.RangeType);
            Assert.AreEqual(null, th.Value);
            Assert.AreEqual(null, th.Formula);
        }


        private void assertIconSetPercentages(IConditionalFormatting cf, IconSet iconset, params double[] vals)
        {
            Assert.AreEqual(1, cf.NumberOfRules);
            IConditionalFormattingRule cr = cf.GetRule(0);
            assertIconSetPercentages(cr, iconset, vals);
        }
        private void assertIconSetPercentages(IConditionalFormattingRule cr, IconSet iconset, params double[] vals)
        {
            Assert.AreEqual(ConditionType.IconSet, cr.ConditionType);
            Assert.AreEqual(ComparisonOperator.NoComparison, cr.ComparisonOperation);
            Assert.AreEqual(null, cr.Formula1);
            Assert.AreEqual(null, cr.Formula2);

            IIconMultiStateFormatting icon = cr.MultiStateFormatting;
            Assert.IsNotNull(icon);
            Assert.AreEqual(iconset, icon.IconSet);
            Assert.AreEqual(false, icon.IsIconOnly);
            Assert.AreEqual(false, icon.IsReversed);

            Assert.IsNotNull(icon.Thresholds);
            Assert.AreEqual(vals.Length, icon.Thresholds.Length);
            for (int i = 0; i < vals.Length; i++)
            {
                Double v = vals[i];
                IConditionalFormattingThreshold th = icon.Thresholds[i] as IConditionalFormattingThreshold;
                Assert.AreEqual(RangeType.PERCENT, th.RangeType);
                Assert.AreEqual(v, th.Value);
                Assert.AreEqual(null, th.Formula);
            }
        }

        private void assertColorScale(IConditionalFormatting cf, params string[] colors)
        {
            Assert.AreEqual(1, cf.NumberOfRules);
            IConditionalFormattingRule cr = cf.GetRule(0);
            assertColorScale(cr, colors);
        }
        private void assertColorScale(IConditionalFormattingRule cr, params string[] colors)
        {
            Assert.AreEqual(ConditionType.ColorScale, cr.ConditionType);
            Assert.AreEqual(ComparisonOperator.NoComparison, cr.ComparisonOperation);
            Assert.AreEqual(null, cr.Formula1);
            Assert.AreEqual(null, cr.Formula2);

            // TODO Implement
            if (cr is HSSFConditionalFormattingRule)
                return;
            IColorScaleFormatting color = cr.ColorScaleFormatting;
            Assert.IsNotNull(color);
            Assert.IsNotNull(color.Colors);
            Assert.IsNotNull(color.Thresholds);
            Assert.AreEqual(colors.Length, color.NumControlPoints);
            Assert.AreEqual(colors.Length, color.Colors.Length);
            Assert.AreEqual(colors.Length, color.Thresholds.Length);

            // Thresholds should be Min / (evenly spaced) / Max
            int steps = 100 / (colors.Length - 1);
            for (int i = 0; i < colors.Length; i++)
            {
                IConditionalFormattingThreshold th = color.Thresholds[i];
                if (i == 0)
                {
                    Assert.AreEqual(RangeType.MIN, th.RangeType);
                }
                else if (i == colors.Length - 1)
                {
                    Assert.AreEqual(RangeType.MAX, th.RangeType);
                }
                else
                {
                    Assert.AreEqual(RangeType.PERCENTILE, th.RangeType);
                    Assert.AreEqual(steps * i, (int)th.Value.Value);
                }
                Assert.AreEqual(null, th.Formula);
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
            Assert.IsFalse(fontFmt.IsItalic);
            Assert.IsFalse(fontFmt.IsBold);
            fontFmt.SetFontStyle(true, true);
            Assert.IsTrue(fontFmt.IsItalic);
            Assert.IsTrue(fontFmt.IsBold);

            Assert.AreEqual(-1, fontFmt.FontHeight); // not modified
            fontFmt.FontHeight = (/*setter*/200);
            Assert.AreEqual(200, fontFmt.FontHeight);
            fontFmt.FontHeight = (/*setter*/100);
            Assert.AreEqual(100, fontFmt.FontHeight);

            Assert.AreEqual(FontSuperScript.None, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontSuperScript.Sub);
            Assert.AreEqual(FontSuperScript.Sub, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontSuperScript.None);
            Assert.AreEqual(FontSuperScript.None, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontSuperScript.Super);
            Assert.AreEqual(FontSuperScript.Super, fontFmt.EscapementType);

            Assert.AreEqual(FontUnderlineType.None, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontUnderlineType.Single);
            Assert.AreEqual(FontUnderlineType.Single, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontUnderlineType.None);
            Assert.AreEqual(FontUnderlineType.None, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontUnderlineType.Double);
            Assert.AreEqual(FontUnderlineType.Double, fontFmt.UnderlineType);

            Assert.AreEqual(-1, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/HSSFColor.Red.Index);
            Assert.AreEqual(HSSFColor.Red.Index, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/HSSFColor.Automatic.Index);
            Assert.AreEqual(HSSFColor.Automatic.Index, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/HSSFColor.Blue.Index);
            Assert.AreEqual(HSSFColor.Blue.Index, fontFmt.FontColorIndex);

            IConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.IsNotNull(cf);

            Assert.AreEqual(1, cf.NumberOfRules);

            IFontFormatting r1fp = cf.GetRule(0).FontFormatting;
            Assert.IsNotNull(r1fp);

            Assert.IsTrue(r1fp.IsItalic);
            Assert.IsTrue(r1fp.IsBold);
            Assert.AreEqual(FontSuperScript.Super, r1fp.EscapementType);
            Assert.AreEqual(FontUnderlineType.Double, r1fp.UnderlineType);
            Assert.AreEqual(HSSFColor.Blue.Index, r1fp.FontColorIndex);

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

            Assert.AreEqual(0, patternFmt.FillBackgroundColor);
            patternFmt.FillBackgroundColor = (/*setter*/HSSFColor.Red.Index);
            Assert.AreEqual(HSSFColor.Red.Index, patternFmt.FillBackgroundColor);

            Assert.AreEqual(0, patternFmt.FillForegroundColor);
            patternFmt.FillForegroundColor = (/*setter*/HSSFColor.Blue.Index);
            Assert.AreEqual(HSSFColor.Blue.Index, patternFmt.FillForegroundColor);

            Assert.AreEqual(FillPattern.NoFill, patternFmt.FillPattern);
            patternFmt.FillPattern = FillPattern.SolidForeground;
            Assert.AreEqual(FillPattern.SolidForeground, patternFmt.FillPattern);
            patternFmt.FillPattern = FillPattern.NoFill;
            Assert.AreEqual(FillPattern.NoFill, patternFmt.FillPattern);
            if (this._testDataProvider.GetSpreadsheetVersion() == SpreadsheetVersion.EXCEL97)
            {
                patternFmt.FillPattern = FillPattern.Bricks;
                Assert.AreEqual(FillPattern.Bricks, patternFmt.FillPattern);
            }

            IConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.IsNotNull(cf);

            Assert.AreEqual(1, cf.NumberOfRules);

            IPatternFormatting r1fp = cf.GetRule(0).PatternFormatting;
            Assert.IsNotNull(r1fp);

            Assert.AreEqual(HSSFColor.Red.Index, r1fp.FillBackgroundColor);
            Assert.AreEqual(HSSFColor.Blue.Index, r1fp.FillForegroundColor);
            if (this._testDataProvider.GetSpreadsheetVersion() == SpreadsheetVersion.EXCEL97)
            {
                Assert.AreEqual(FillPattern.Bricks, r1fp.FillPattern);
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
                Assert.AreEqual(border, borderFmt.BorderTop);
                borderFmt.BorderBottom = border;
                Assert.AreEqual(border, borderFmt.BorderBottom);
                borderFmt.BorderLeft = border;
                Assert.AreEqual(border, borderFmt.BorderLeft);
                borderFmt.BorderRight = border;
                Assert.AreEqual(border, borderFmt.BorderRight);
                borderFmt.BorderDiagonal = border;
                Assert.AreEqual(border, borderFmt.BorderDiagonal);
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

            Assert.AreEqual(BorderStyle.None, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderStyle.Dotted);
            Assert.AreEqual(BorderStyle.Dotted, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderStyle.None);
            Assert.AreEqual(BorderStyle.None, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderStyle.Thick);
            Assert.AreEqual(BorderStyle.Thick, borderFmt.BorderBottom);

            Assert.AreEqual(BorderStyle.None, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderStyle.Dotted);
            Assert.AreEqual(BorderStyle.Dotted, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderStyle.None);
            Assert.AreEqual(BorderStyle.None, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderStyle.Thick);
            Assert.AreEqual(BorderStyle.Thick, borderFmt.BorderTop);

            Assert.AreEqual(BorderStyle.None, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderStyle.Dotted);
            Assert.AreEqual(BorderStyle.Dotted, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderStyle.None);
            Assert.AreEqual(BorderStyle.None, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderStyle.Thin);
            Assert.AreEqual(BorderStyle.Thin, borderFmt.BorderLeft);

            Assert.AreEqual(BorderStyle.None, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderStyle.Dotted);
            Assert.AreEqual(BorderStyle.Dotted, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderStyle.None);
            Assert.AreEqual(BorderStyle.None, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderStyle.Hair);
            Assert.AreEqual(BorderStyle.Hair, borderFmt.BorderRight);

            IConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.IsNotNull(cf);

            Assert.AreEqual(1, cf.NumberOfRules);

            IBorderFormatting r1fp = cf.GetRule(0).BorderFormatting;
            Assert.IsNotNull(r1fp);
            Assert.AreEqual(BorderStyle.Thick, r1fp.BorderBottom);
            Assert.AreEqual(BorderStyle.Thick, r1fp.BorderTop);
            Assert.AreEqual(BorderStyle.Thin, r1fp.BorderLeft);
            Assert.AreEqual(BorderStyle.Hair, r1fp.BorderRight);

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

            Assert.AreEqual(IconSet.GYRB_4_TRAFFIC_LIGHTS, iconFmt.IconSet);
            Assert.AreEqual(4, iconFmt.Thresholds.Length);
            Assert.AreEqual(false, iconFmt.IsIconOnly);
            Assert.AreEqual(false, iconFmt.IsReversed);

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
            Assert.AreEqual(1, sheetCF.NumConditionalFormattings);

            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual(1, cf.NumberOfRules);
            rule1 = cf.GetRule(0);
            Assert.AreEqual(ConditionType.IconSet, rule1.ConditionType);
            iconFmt = rule1.MultiStateFormatting;

            Assert.AreEqual(IconSet.GYRB_4_TRAFFIC_LIGHTS, iconFmt.IconSet);
            Assert.AreEqual(4, iconFmt.Thresholds.Length);
            Assert.AreEqual(true, iconFmt.IsIconOnly);
            Assert.AreEqual(false, iconFmt.IsReversed);
            Assert.AreEqual(RangeType.MIN, iconFmt.Thresholds[0].RangeType);
            Assert.AreEqual(RangeType.NUMBER, iconFmt.Thresholds[1].RangeType);
            Assert.AreEqual(RangeType.PERCENT, iconFmt.Thresholds[2].RangeType);
            Assert.AreEqual(RangeType.MAX, iconFmt.Thresholds[3].RangeType);
            Assert.AreEqual(null, iconFmt.Thresholds[0].Value);
            Assert.AreEqual(10d, iconFmt.Thresholds[1].Value);
            Assert.AreEqual(75d, iconFmt.Thresholds[2].Value);
            Assert.AreEqual(null, iconFmt.Thresholds[3].Value);

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

            Assert.AreEqual(3, clrFmt.NumControlPoints);
            Assert.AreEqual(3, clrFmt.Colors.Length);
            Assert.AreEqual(3, clrFmt.Thresholds.Length);

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
            Assert.AreEqual(1, sheetCF.NumConditionalFormattings);

            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual(1, cf.NumberOfRules);
            rule1 = cf.GetRule(0);
            clrFmt = rule1.ColorScaleFormatting;
            Assert.AreEqual(ConditionType.ColorScale, rule1.ConditionType);

            Assert.AreEqual(3, clrFmt.NumControlPoints);
            Assert.AreEqual(3, clrFmt.Colors.Length);
            Assert.AreEqual(3, clrFmt.Thresholds.Length);
            Assert.AreEqual(RangeType.MIN, clrFmt.Thresholds[0].RangeType);
            Assert.AreEqual(RangeType.NUMBER, clrFmt.Thresholds[1].RangeType);
            Assert.AreEqual(RangeType.MAX, clrFmt.Thresholds[2].RangeType);
            Assert.AreEqual(null, clrFmt.Thresholds[0].Value);
            Assert.AreEqual(10d, clrFmt.Thresholds[1].Value);
            Assert.AreEqual(null, clrFmt.Thresholds[2].Value);

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

            Assert.AreEqual(false, dbFmt.IsIconOnly);
            Assert.AreEqual(true, dbFmt.IsLeftToRight);
            Assert.AreEqual(0, dbFmt.WidthMin);
            Assert.AreEqual(100, dbFmt.WidthMax);
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
            Assert.AreEqual(1, sheetCF.NumConditionalFormattings);

            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual(1, cf.NumberOfRules);
            rule1 = cf.GetRule(0);
            dbFmt = rule1.DataBarFormatting;
            Assert.AreEqual(ConditionType.DataBar, rule1.ConditionType);

            Assert.AreEqual(false, dbFmt.IsIconOnly);
            Assert.AreEqual(true, dbFmt.IsLeftToRight);
            Assert.AreEqual(0, dbFmt.WidthMin);
            Assert.AreEqual(100, dbFmt.WidthMax);
            AssertColour(colorHex, dbFmt.Color);
            Assert.AreEqual(RangeType.MIN, dbFmt.MinThreshold.RangeType);
            Assert.AreEqual(RangeType.MAX, dbFmt.MaxThreshold.RangeType);
            Assert.AreEqual(null, dbFmt.MinThreshold.Value);
            Assert.AreEqual(null, dbFmt.MaxThreshold.Value);

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
    }

}