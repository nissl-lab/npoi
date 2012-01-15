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

namespace NPOI.SS.UserModel
{
    using System;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * @author Dmitriy Kumshayev
     * @author Yegor Kozlov
     */
    [TestClass]
    public abstract class BaseTestConditionalFormatting
    {
        private ITestDataProvider _testDataProvider;

        public BaseTestConditionalFormatting(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }
        [TestMethod]
        public void TestBasic()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            SheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;

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

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("1");
            ConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule("2");
            ConditionalFormattingRule rule3 = sheetCF.CreateConditionalFormattingRule("3");
            ConditionalFormattingRule rule4 = sheetCF.CreateConditionalFormattingRule("4");
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
                        (ConditionalFormattingRule)null);
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
                        new ConditionalFormattingRule[0]);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.StartsWith("cfRules must not be empty"));
            }

            try
            {
                sheetCF.AddConditionalFormatting(
                        new CellRangeAddress[] { CellRangeAddress.ValueOf("A1:A3") },
                        new ConditionalFormattingRule[] { rule1, rule2, rule3, rule4 });
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.StartsWith("Number of rules must not exceed 3"));
            }
        }

        /**
         * Test format conditions based on a bool formula
         */
        [TestMethod]
        public void TestBooleanFormulaConditions()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            SheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("SUM(A1:A5)>10");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_FORMULA, rule1.ConditionType);
            Assert.AreEqual("SUM(A1:A5)>10", rule1.Formula1);
            int formatIndex1 = sheetCF.AddConditionalFormatting(
                    new CellRangeAddress[]{
                        CellRangeAddress.ValueOf("B1"),
                        CellRangeAddress.ValueOf("C3"),
                }, rule1);
            Assert.AreEqual(0, formatIndex1);
            Assert.AreEqual(1, sheetCF.NumConditionalFormattings);
            CellRangeAddress[] ranges1 = sheetCF.GetConditionalFormattingAt(formatIndex1).FormattingRanges;
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
            CellRangeAddress[] ranges2 = sheetCF.GetConditionalFormattingAt(formatIndex2).FormattingRanges;
            Assert.AreEqual(1, ranges2.Length);
            Assert.AreEqual("B1:B3", ranges2[0].FormatAsString());
        }
        [TestMethod]
        public void TestSingleFormulaConditions()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            SheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.EQUAL, "SUM(A1:A5)+10");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule1.ConditionType);
            Assert.AreEqual("SUM(A1:A5)+10", rule1.Formula1);
            Assert.AreEqual(ComparisonOperator.EQUAL, rule1.ComparisonOperation);

            ConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NOT_EQUAL, "15");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule2.ConditionType);
            Assert.AreEqual("15", rule2.Formula1);
            Assert.AreEqual(ComparisonOperator.NOT_EQUAL, rule2.ComparisonOperation);

            ConditionalFormattingRule rule3 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NOT_EQUAL, "15");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule3.ConditionType);
            Assert.AreEqual("15", rule3.Formula1);
            Assert.AreEqual(ComparisonOperator.NOT_EQUAL, rule3.ComparisonOperation);

            ConditionalFormattingRule rule4 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.GT, "0");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule4.ConditionType);
            Assert.AreEqual("0", rule4.Formula1);
            Assert.AreEqual(ComparisonOperator.GT, rule4.ComparisonOperation);

            ConditionalFormattingRule rule5 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.LT, "0");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule5.ConditionType);
            Assert.AreEqual("0", rule5.Formula1);
            Assert.AreEqual(ComparisonOperator.LT, rule5.ComparisonOperation);

            ConditionalFormattingRule rule6 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.GE, "0");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule6.ConditionType);
            Assert.AreEqual("0", rule6.Formula1);
            Assert.AreEqual(ComparisonOperator.GE, rule6.ComparisonOperation);

            ConditionalFormattingRule rule7 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.LE, "0");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule7.ConditionType);
            Assert.AreEqual("0", rule7.Formula1);
            Assert.AreEqual(ComparisonOperator.LE, rule7.ComparisonOperation);

            ConditionalFormattingRule rule8 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.BETWEEN, "0", "5");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule8.ConditionType);
            Assert.AreEqual("0", rule8.Formula1);
            Assert.AreEqual("5", rule8.Formula2);
            Assert.AreEqual(ComparisonOperator.BETWEEN, rule8.ComparisonOperation);

            ConditionalFormattingRule rule9 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.NOT_BETWEEN, "0", "5");
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule9.ConditionType);
            Assert.AreEqual("0", rule9.Formula1);
            Assert.AreEqual("5", rule9.Formula2);
            Assert.AreEqual(ComparisonOperator.NOT_BETWEEN, rule9.ComparisonOperation);
        }
        [TestMethod]
        public void TestCopy()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            ISheet sheet2 = wb.CreateSheet();
            SheetConditionalFormatting sheet1CF = sheet1.SheetConditionalFormatting;
            SheetConditionalFormatting sheet2CF = sheet2.SheetConditionalFormatting;
            Assert.AreEqual(0, sheet1CF.NumConditionalFormattings);
            Assert.AreEqual(0, sheet2CF.NumConditionalFormattings);

            ConditionalFormattingRule rule1 = sheet1CF.CreateConditionalFormattingRule(
                    ComparisonOperator.EQUAL, "SUM(A1:A5)+10");

            ConditionalFormattingRule rule2 = sheet1CF.CreateConditionalFormattingRule(
                    ComparisonOperator.NOT_EQUAL, "15");

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

            ConditionalFormatting sheet2cf = sheet2CF.GetConditionalFormattingAt(0);
            Assert.AreEqual(2, sheet2cf.NumberOfRules);
            Assert.AreEqual("SUM(A1:A5)+10", sheet2cf.GetRule(0).Formula1);
            Assert.AreEqual(ComparisonOperator.EQUAL, sheet2cf.GetRule(0).ComparisonOperation);
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, sheet2cf.GetRule(0).ConditionType);
            Assert.AreEqual("15", sheet2cf.GetRule(1).Formula1);
            Assert.AreEqual(ComparisonOperator.NOT_EQUAL, sheet2cf.GetRule(1).ComparisonOperation);
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, sheet2cf.GetRule(1).ConditionType);
        }
        [TestMethod]
        public void TestRemove()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet();
            SheetConditionalFormatting sheetCF = sheet1.SheetConditionalFormatting;
            Assert.AreEqual(0, sheetCF.NumConditionalFormattings);

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.EQUAL, "SUM(A1:A5)");

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
        }
        [TestMethod]
        public void TestCreateCF() {
        IWorkbook workbook = _testDataProvider.CreateWorkbook();
        ISheet sheet = workbook.CreateSheet();
        String formula = "7";

        SheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

        ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(formula);
        FontFormatting fontFmt = rule1.CreateFontFormatting();
        fontFmt.SetFontStyle=(true, false);

        BorderFormatting bordFmt = rule1.CreateBorderFormatting();
        bordFmt.BorderBottom=(/*setter*/BorderFormatting.BORDER_THIN);
        bordFmt.BorderTop=(/*setter*/BorderFormatting.BORDER_THICK);
        bordFmt.BorderLeft=(/*setter*/BorderFormatting.BORDER_DASHED);
        bordFmt.BorderRight=(/*setter*/BorderFormatting.BORDER_DOTTED);

        PatternFormatting patternFmt = rule1.CreatePatternFormatting();
        patternFmt.FillBackgroundColor=(/*setter*/IndexedColors.YELLOW.index);


        ConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.BETWEEN, "1", "2");
        ConditionalFormattingRule [] cfRules =
        {
            rule1, rule2
        };

        short col = 1;
        CellRangeAddress [] regions = {
            new CellRangeAddress(0, 65535, col, col)
        };

        sheetCF.AddConditionalFormatting(regions, cfRules);
        sheetCF.AddConditionalFormatting(regions, cfRules);

        // Verification
        Assert.AreEqual(2, sheetCF.NumConditionalFormattings);
        sheetCF.RemoveConditionalFormatting(1);
        Assert.AreEqual(1, sheetCF.NumConditionalFormattings);
        ConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
        Assert.IsNotNull(cf);

        regions = cf.FormattingRanges;
        Assert.IsNotNull(regions);
        Assert.AreEqual(1, regions.Length);
        CellRangeAddress r = regions[0];
        Assert.AreEqual(1, r.FirstColumn);
        Assert.AreEqual(1, r.LastColumn);
        Assert.AreEqual(0, r.FirstRow);
        Assert.AreEqual(65535, r.LastRow);

        Assert.AreEqual(2, cf.NumberOfRules);

        rule1 = cf.GetRule(0);
        Assert.AreEqual("7",rule1.Formula1);
        Assert.IsNull(rule1.Formula2);

        FontFormatting    r1fp = rule1.FontFormatting;
        Assert.IsNotNull(r1fp);

        Assert.IsTrue(r1fp.IsItalic());
        Assert.IsFalse(r1fp.IsBold());

        BorderFormatting  r1bf = rule1.BorderFormatting;
        Assert.IsNotNull(r1bf);
        Assert.AreEqual(BorderFormatting.BORDER_THIN, r1bf.BorderBottom);
        Assert.AreEqual(BorderFormatting.BORDER_THICK,r1bf.BorderTop);
        Assert.AreEqual(BorderFormatting.BORDER_DASHED,r1bf.BorderLeft);
        Assert.AreEqual(BorderFormatting.BORDER_DOTTED,r1bf.BorderRight);

        PatternFormatting r1pf = rule1.PatternFormatting;
        Assert.IsNotNull(r1pf);
//        Assert.AreEqual(IndexedColors.YELLOW.index,r1pf.FillBackgroundColor);

        rule2 = cf.GetRule(1);
        Assert.AreEqual("2",rule2.Formula2);
        Assert.AreEqual("1",rule2.Formula1);
    }
        [TestMethod]
        public void TestClone()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();
            String formula = "7";

            SheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(formula);
            FontFormatting fontFmt = rule1.CreateFontFormatting();
            fontFmt.SetFontStyle(true, false);

            PatternFormatting patternFmt = rule1.CreatePatternFormatting();
            patternFmt.FillBackgroundColor = (/*setter*/IndexedColors.YELLOW.index);


            ConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.BETWEEN, "1", "2");
            ConditionalFormattingRule[] cfRules =
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
            }
            catch (RuntimeException e)
            {
                if (e.Message.IndexOf("needs to define a clone method") > 0)
                {
                    Assert.Fail("Indentified bug 45682");
                }
                throw e;
            }
            Assert.AreEqual(2, wb.NumberOfSheets);
        }
        [TestMethod]
        public void TestShiftRows()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet();

            SheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.BETWEEN, "SUM(A10:A15)", "1+SUM(B16:B30)");
            FontFormatting fontFmt = rule1.CreateFontFormatting();
            fontFmt.SetFontStyle(true, false);

            PatternFormatting patternFmt = rule1.CreatePatternFormatting();
            patternFmt.FillBackgroundColor = (/*setter*/IndexedColors.YELLOW.index);
            ConditionalFormattingRule[] cfRules = { rule1, };

            CellRangeAddress[] regions = {
            new CellRangeAddress(2, 4, 0, 0), // A3:A5
        };
            sheetCF.AddConditionalFormatting(regions, cfRules);

            // This row-shift should destroy the CF region
            sheet.ShiftRows(10, 20, -9);
            Assert.AreEqual(0, sheetCF.NumConditionalFormattings);

            // re-add the CF
            sheetCF.AddConditionalFormatting(regions, cfRules);

            // This row shift should only affect the formulas
            sheet.ShiftRows(14, 17, 8);
            ConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual("SUM(A10:A23)", cf.GetRule(0).Formula1);
            Assert.AreEqual("1+SUM(B24:B30)", cf.GetRule(0).Formula2);

            sheet.ShiftRows(0, 8, 21);
            cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual("SUM(A10:A21)", cf.GetRule(0).Formula1);
            Assert.AreEqual("1+SUM(#REF!)", cf.GetRule(0).Formula2);
        }
        [TestMethod]
        protected void TestRead(String filename)
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(filename);
            ISheet sh = wb.GetSheet("CF");
            SheetConditionalFormatting sheetCF = sh.SheetConditionalFormatting;
            Assert.AreEqual(3, sheetCF.NumConditionalFormattings);

            ConditionalFormatting cf1 = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual(2, cf1.NumberOfRules);

            CellRangeAddress[] regions1 = cf1.FormattingRanges;
            Assert.AreEqual(1, regions1.Length);
            Assert.AreEqual("A1:A8", regions1[0].FormatAsString());

            // CF1 has two rules: values less than -3 are bold-italic red, values greater than 3 are green
            ConditionalFormattingRule rule1 = cf1.GetRule(0);
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule1.ConditionType);
            Assert.AreEqual(ComparisonOperator.GT, rule1.ComparisonOperation);
            Assert.AreEqual("3", rule1.Formula1);
            Assert.IsNull(rule1.Formula2);
            // Fills and borders are not Set
            Assert.IsNull(rule1.PatternFormatting);
            Assert.IsNull(rule1.BorderFormatting);

            FontFormatting fmt1 = rule1.FontFormatting;
            //        Assert.AreEqual(IndexedColors.GREEN.index, fmt1.FontColorIndex);
            Assert.IsTrue(fmt1.IsBold());
            Assert.IsFalse(fmt1.IsItalic());

            ConditionalFormattingRule rule2 = cf1.GetRule(1);
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule2.ConditionType);
            Assert.AreEqual(ComparisonOperator.LT, rule2.ComparisonOperation);
            Assert.AreEqual("-3", rule2.Formula1);
            Assert.IsNull(rule2.Formula2);
            Assert.IsNull(rule2.PatternFormatting);
            Assert.IsNull(rule2.BorderFormatting);

            FontFormatting fmt2 = rule2.FontFormatting;
            //        Assert.AreEqual(IndexedColors.RED.index, fmt2.FontColorIndex);
            Assert.IsTrue(fmt2.IsBold());
            Assert.IsTrue(fmt2.IsItalic());


            ConditionalFormatting cf2 = sheetCF.GetConditionalFormattingAt(1);
            Assert.AreEqual(1, cf2.NumberOfRules);
            CellRangeAddress[] regions2 = cf2.FormattingRanges;
            Assert.AreEqual(1, regions2.Length);
            Assert.AreEqual("B9", regions2[0].FormatAsString());

            ConditionalFormattingRule rule3 = cf2.GetRule(0);
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_FORMULA, rule3.ConditionType);
            Assert.AreEqual(ComparisonOperator.NO_COMPARISON, rule3.ComparisonOperation);
            Assert.AreEqual("$A$8>5", rule3.Formula1);
            Assert.IsNull(rule3.Formula2);

            FontFormatting fmt3 = rule3.FontFormatting;
            //        Assert.AreEqual(IndexedColors.RED.index, fmt3.FontColorIndex);
            Assert.IsTrue(fmt3.IsBold());
            Assert.IsTrue(fmt3.IsItalic());

            PatternFormatting fmt4 = rule3.PatternFormatting;
            //        Assert.AreEqual(IndexedColors.LIGHT_CORNFLOWER_BLUE.index, fmt4.FillBackgroundColor);
            //        Assert.AreEqual(IndexedColors.AUTOMATIC.index, fmt4.FillForegroundColor);
            Assert.AreEqual(PatternFormatting.NO_FILL, fmt4.FillPattern);
            // borders are not Set
            Assert.IsNull(rule3.BorderFormatting);

            ConditionalFormatting cf3 = sheetCF.GetConditionalFormattingAt(2);
            CellRangeAddress[] regions3 = cf3.FormattingRanges;
            Assert.AreEqual(1, regions3.Length);
            Assert.AreEqual("B1:B7", regions3[0].FormatAsString());
            Assert.AreEqual(2, cf3.NumberOfRules);

            ConditionalFormattingRule rule4 = cf3.GetRule(0);
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule4.ConditionType);
            Assert.AreEqual(ComparisonOperator.LE, rule4.ComparisonOperation);
            Assert.AreEqual("\"AAA\"", rule4.Formula1);
            Assert.IsNull(rule4.Formula2);

            ConditionalFormattingRule rule5 = cf3.GetRule(1);
            Assert.AreEqual(ConditionalFormattingRule.CONDITION_TYPE_CELL_VALUE_IS, rule5.ConditionType);
            Assert.AreEqual(ComparisonOperator.BETWEEN, rule5.ComparisonOperation);
            Assert.AreEqual("\"A\"", rule5.Formula1);
            Assert.AreEqual("\"AAA\"", rule5.Formula2);
        }

        [TestMethod]
        public void TestCreateFontFormatting()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            SheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.EQUAL, "7");
            FontFormatting fontFmt = rule1.CreateFontFormatting();
            Assert.IsFalse(fontFmt.IsItalic());
            Assert.IsFalse(fontFmt.IsBold());
            fontFmt.SetFontStyle(true, true);
            Assert.IsTrue(fontFmt.IsItalic());
            Assert.IsTrue(fontFmt.IsBold());

            Assert.AreEqual(-1, fontFmt.FontHeight); // not modified
            fontFmt.FontHeight = (/*setter*/200);
            Assert.AreEqual(200, fontFmt.FontHeight);
            fontFmt.FontHeight = (/*setter*/100);
            Assert.AreEqual(100, fontFmt.FontHeight);

            Assert.AreEqual(FontFormatting.SS_NONE, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontFormatting.SS_SUB);
            Assert.AreEqual(FontFormatting.SS_SUB, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontFormatting.SS_NONE);
            Assert.AreEqual(FontFormatting.SS_NONE, fontFmt.EscapementType);
            fontFmt.EscapementType = (/*setter*/FontFormatting.SS_SUPER);
            Assert.AreEqual(FontFormatting.SS_SUPER, fontFmt.EscapementType);

            Assert.AreEqual(FontFormatting.U_NONE, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontFormatting.U_SINGLE);
            Assert.AreEqual(FontFormatting.U_SINGLE, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontFormatting.U_NONE);
            Assert.AreEqual(FontFormatting.U_NONE, fontFmt.UnderlineType);
            fontFmt.UnderlineType = (/*setter*/FontFormatting.U_DOUBLE);
            Assert.AreEqual(FontFormatting.U_DOUBLE, fontFmt.UnderlineType);

            Assert.AreEqual(-1, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/IndexedColors.RED.index);
            Assert.AreEqual(IndexedColors.RED.index, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/IndexedColors.AUTOMATIC.index);
            Assert.AreEqual(IndexedColors.AUTOMATIC.index, fontFmt.FontColorIndex);
            fontFmt.FontColorIndex = (/*setter*/IndexedColors.BLUE.index);
            Assert.AreEqual(IndexedColors.BLUE.index, fontFmt.FontColorIndex);

            ConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            ConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.IsNotNull(cf);

            Assert.AreEqual(1, cf.NumberOfRules);

            FontFormatting r1fp = cf.GetRule(0).FontFormatting;
            Assert.IsNotNull(r1fp);

            Assert.IsTrue(r1fp.IsItalic());
            Assert.IsTrue(r1fp.IsBold());
            Assert.AreEqual(FontFormatting.SS_SUPER, r1fp.EscapementType);
            Assert.AreEqual(FontFormatting.U_DOUBLE, r1fp.UnderlineType);
            Assert.AreEqual(IndexedColors.BLUE.index, r1fp.FontColorIndex);

        }
        [TestMethod]
        public void TestCreatePatternFormatting()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            SheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.EQUAL, "7");
            PatternFormatting patternFmt = rule1.CreatePatternFormatting();

            Assert.AreEqual(0, patternFmt.FillBackgroundColor);
            patternFmt.FillBackgroundColor = (/*setter*/IndexedColors.RED.index);
            Assert.AreEqual(IndexedColors.RED.index, patternFmt.FillBackgroundColor);

            Assert.AreEqual(0, patternFmt.FillForegroundColor);
            patternFmt.FillForegroundColor = (/*setter*/IndexedColors.BLUE.index);
            Assert.AreEqual(IndexedColors.BLUE.index, patternFmt.FillForegroundColor);

            Assert.AreEqual(PatternFormatting.NO_FILL, patternFmt.FillPattern);
            patternFmt.FillPattern = (/*setter*/PatternFormatting.SOLID_FOREGROUND);
            Assert.AreEqual(PatternFormatting.SOLID_FOREGROUND, patternFmt.FillPattern);
            patternFmt.FillPattern = (/*setter*/PatternFormatting.NO_FILL);
            Assert.AreEqual(PatternFormatting.NO_FILL, patternFmt.FillPattern);
            patternFmt.FillPattern = (/*setter*/PatternFormatting.BRICKS);
            Assert.AreEqual(PatternFormatting.BRICKS, patternFmt.FillPattern);

            ConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            ConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.IsNotNull(cf);

            Assert.AreEqual(1, cf.NumberOfRules);

            PatternFormatting r1fp = cf.GetRule(0).PatternFormatting;
            Assert.IsNotNull(r1fp);

            Assert.AreEqual(IndexedColors.RED.index, r1fp.FillBackgroundColor);
            Assert.AreEqual(IndexedColors.BLUE.index, r1fp.FillForegroundColor);
            Assert.AreEqual(PatternFormatting.BRICKS, r1fp.FillPattern);
        }
        [TestMethod]
        public void TestCreateBorderFormatting()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            SheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            ConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.EQUAL, "7");
            BorderFormatting borderFmt = rule1.CreateBorderFormatting();

            Assert.AreEqual(BorderFormatting.BORDER_NONE, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderFormatting.BORDER_DOTTED);
            Assert.AreEqual(BorderFormatting.BORDER_DOTTED, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderFormatting.BORDER_NONE);
            Assert.AreEqual(BorderFormatting.BORDER_NONE, borderFmt.BorderBottom);
            borderFmt.BorderBottom = (/*setter*/BorderFormatting.BORDER_THICK);
            Assert.AreEqual(BorderFormatting.BORDER_THICK, borderFmt.BorderBottom);

            Assert.AreEqual(BorderFormatting.BORDER_NONE, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderFormatting.BORDER_DOTTED);
            Assert.AreEqual(BorderFormatting.BORDER_DOTTED, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderFormatting.BORDER_NONE);
            Assert.AreEqual(BorderFormatting.BORDER_NONE, borderFmt.BorderTop);
            borderFmt.BorderTop = (/*setter*/BorderFormatting.BORDER_THICK);
            Assert.AreEqual(BorderFormatting.BORDER_THICK, borderFmt.BorderTop);

            Assert.AreEqual(BorderFormatting.BORDER_NONE, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderFormatting.BORDER_DOTTED);
            Assert.AreEqual(BorderFormatting.BORDER_DOTTED, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderFormatting.BORDER_NONE);
            Assert.AreEqual(BorderFormatting.BORDER_NONE, borderFmt.BorderLeft);
            borderFmt.BorderLeft = (/*setter*/BorderFormatting.BORDER_THIN);
            Assert.AreEqual(BorderFormatting.BORDER_THIN, borderFmt.BorderLeft);

            Assert.AreEqual(BorderFormatting.BORDER_NONE, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderFormatting.BORDER_DOTTED);
            Assert.AreEqual(BorderFormatting.BORDER_DOTTED, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderFormatting.BORDER_NONE);
            Assert.AreEqual(BorderFormatting.BORDER_NONE, borderFmt.BorderRight);
            borderFmt.BorderRight = (/*setter*/BorderFormatting.BORDER_HAIR);
            Assert.AreEqual(BorderFormatting.BORDER_HAIR, borderFmt.BorderRight);

            ConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { CellRangeAddress.ValueOf("A1:A5") };

            sheetCF.AddConditionalFormatting(regions, cfRules);

            // Verification
            ConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.IsNotNull(cf);

            Assert.AreEqual(1, cf.NumberOfRules);

            BorderFormatting r1fp = cf.GetRule(0).BorderFormatting;
            Assert.IsNotNull(r1fp);
            Assert.AreEqual(BorderFormatting.BORDER_THICK, r1fp.BorderBottom);
            Assert.AreEqual(BorderFormatting.BORDER_THICK, r1fp.BorderTop);
            Assert.AreEqual(BorderFormatting.BORDER_THIN, r1fp.BorderLeft);
            Assert.AreEqual(BorderFormatting.BORDER_HAIR, r1fp.BorderRight);

        }
    }

}