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


using System.Collections.Generic;
using System.IO;

namespace TestCases.SS.UserModel
{
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    [TestFixture]
    public class ConditionalFormattingEvalTest
    {
        private XSSFWorkbook wb;
        private ISheet sheet;
        private XSSFFormulaEvaluator formulaEval;
        private ConditionalFormattingEvaluator cfe;
        private CellReference ref1;
        private List<EvaluationConditionalFormatRule> rules;

        [SetUp]
        public void OpenWB()
        {
            wb = XSSFTestDataSamples.OpenSampleWorkbook("ConditionalFormattingSamples.xlsx");
            formulaEval = new XSSFFormulaEvaluator(wb);
            cfe = new ConditionalFormattingEvaluator(wb, formulaEval);
        }

        [TearDown]
        public void CloseWB()
        {
            formulaEval = null;
            cfe = null;
            ref1 = null;
            rules = null;
            try
            {
                if(wb != null)
                    wb.Close();
            }
            catch(IOException)
            {
                // keep going, this shouldn't cancel things
            }
        }

        [Test]
        public void TestFormattingEvaluation()
        {
            sheet = wb.GetSheet("Products1");

            GetRulesFor(12, 1);
            ClassicAssert.AreEqual(1, rules.Count, "wrong # of rules for " + ref1);
            ClassicAssert.AreEqual("FFFFEB9C", GetColor(rules[0].Rule.PatternFormatting.FillBackgroundColorColor),
                "wrong bg color for " + ref1);
            ClassicAssert.IsFalse(rules[0].Rule.FontFormatting.IsItalic, "should not be italic " + ref1);

            GetRulesFor(16, 3);
            ClassicAssert.AreEqual(1, rules.Count, "wrong # of rules for " + ref1);
            ClassicAssert.AreEqual(0.7999816888943144d, GetTint(rules[0].Rule.PatternFormatting.FillBackgroundColorColor), 0.000000000000001,
                "wrong bg color for " + ref1);

            GetRulesFor(12, 3);
            ClassicAssert.AreEqual(0, rules.Count, "wrong # of rules for " + ref1);

            sheet = wb.GetSheet("Products2");

            GetRulesFor(15, 1);
            ClassicAssert.AreEqual(1, rules.Count, "wrong # of rules for " + ref1);
            ClassicAssert.AreEqual("FFFFEB9C", GetColor(rules[0].Rule.PatternFormatting.FillBackgroundColorColor)
                , "wrong bg color for " + ref1);

            GetRulesFor(20, 3);
            ClassicAssert.AreEqual(0, rules.Count, "wrong # of rules for " + ref1);

            // now change a cell value that's an input for the rules
            ICell cell = sheet.GetRow(1).GetCell(6);
            cell.SetCellValue("Dairy");
            formulaEval.NotifyUpdateCell(cell);
            cell = sheet.GetRow(4).GetCell(6);
            cell.SetCellValue(500);
            formulaEval.NotifyUpdateCell(cell);
            // need to throw away all evaluations, since we don't know how value changes may have affected format formulas
            cfe.ClearAllCachedValues();

            // test that the conditional validation evaluations changed
            GetRulesFor(15, 1);
            ClassicAssert.AreEqual(0, rules.Count, "wrong # of rules for " + ref1);

            GetRulesFor(20, 3);
            ClassicAssert.AreEqual(1, rules.Count, "wrong # of rules for " + ref1);
            ClassicAssert.AreEqual(0.7999816888943144d, GetTint(rules[0].Rule.PatternFormatting.FillBackgroundColorColor), 0.000000000000001,
                "wrong bg color for " + ref1);

            GetRulesFor(20, 1);
            ClassicAssert.AreEqual(1, rules.Count, "wrong # of rules for " + ref1);
            ClassicAssert.AreEqual("FFFFEB9C", GetColor(rules[0].Rule.PatternFormatting.FillBackgroundColorColor),
                "wrong bg color for " + ref1);

            sheet = wb.GetSheet("Book tour");

            GetRulesFor(8, 2);
            ClassicAssert.AreEqual(1, rules.Count, "wrong # of rules for " + ref1);

        }

        [Test]
        public void TestFormattingOnUndefinedCell()
        {
            wb = XSSFTestDataSamples.OpenSampleWorkbook("conditional_formatting_with_formula_on_second_sheet.xlsx");
            formulaEval = new XSSFFormulaEvaluator(wb);
            cfe = new ConditionalFormattingEvaluator(wb, formulaEval);

            sheet = wb.GetSheet("Sales Plan");
            GetRulesFor(9, 2);
            ClassicAssert.AreNotEqual(0, rules.Count, "No rules for " + ref1);
            ClassicAssert.AreEqual("FFFFFF00", GetColor(rules[0].Rule.PatternFormatting.FillBackgroundColorColor), 
                "wrong bg color for " + ref1);
        }

        private List<EvaluationConditionalFormatRule> GetRulesFor(int row, int col)
        {
            ref1 = new CellReference(sheet.SheetName, row, col, false, false);
            return rules = cfe.GetConditionalFormattingForCell(ref1);
        }

        private static string GetColor(IColor color)
        {
            XSSFColor c = XSSFColor.ToXSSFColor(color);
            return c.ARGBHex;
        }

        private static double GetTint(IColor color)
        {
            XSSFColor c = XSSFColor.ToXSSFColor(color);
            return c.Tint;
        }
    }
}
