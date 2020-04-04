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

namespace TestCases.HSSF.UserModel
{
    using System;

    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.HSSF.Record;
    using NUnit.Framework;
    using NPOI.SS.Util;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Record.CF;
    using TestCases.SS.UserModel;

    /**
* 
* @author Dmitriy Kumshayev
*/
    [TestFixture]
    public class TestHSSFConditionalFormatting : BaseTestConditionalFormatting
    {
        protected override void AssertColour(String hexExpected, IColor actual)
        {
            Assert.IsNotNull(actual, "Colour must be given");

            if (actual is HSSFColor) {
                HSSFColor colour = (HSSFColor)actual;
                Assert.AreEqual(hexExpected, colour.GetHexString());
            } else {
                HSSFExtendedColor colour = (HSSFExtendedColor)actual;
                if (hexExpected.Length == 8)
                {
                    Assert.AreEqual(hexExpected, colour.ARGBHex);
                }
                else
                {
                    Assert.AreEqual(hexExpected, colour.ARGBHex.Substring(2));
                }
            }
        }
        [Test]
        public void TestReadOffice2007()
        {
            testReadOffice2007("NewStyleConditionalFormattings.xls");
        }

        [Test]
        public void Test53691()
        {
            ISheetConditionalFormatting cf;
            IWorkbook wb = HSSFITestDataProvider.Instance.OpenSampleWorkbook("53691.xls");
            /*
            FileInputStream s = new FileInputStream("C:\\temp\\53691bbadfixed.xls");
            try {
                wb = new HSSFWorkbook(s);
            } finally {
                s.Close();
            }

            wb.RemoveSheetAt(1);*/

            // Initially it is good
            WriteTemp53691(wb, "agood");

            // clone sheet corrupts it
            ISheet sheet = wb.CloneSheet(0);
            WriteTemp53691(wb, "bbad");

            // removing the sheet Makes it good again
            wb.RemoveSheetAt(wb.GetSheetIndex(sheet));
            WriteTemp53691(wb, "cgood");

            // cloning again and removing the conditional formatting Makes it good again
            sheet = wb.CloneSheet(0);
            RemoveConditionalFormatting(sheet);
            WriteTemp53691(wb, "dgood");

            // cloning the conditional formatting manually Makes it bad again
            cf = sheet.SheetConditionalFormatting;
            ISheetConditionalFormatting scf = wb.GetSheetAt(0).SheetConditionalFormatting;
            for (int j = 0; j < scf.NumConditionalFormattings; j++)
            {
                cf.AddConditionalFormatting(scf.GetConditionalFormattingAt(j));
            }
            WriteTemp53691(wb, "ebad");

            // remove all conditional formatting for comparing BIFF output
            RemoveConditionalFormatting(sheet);
            RemoveConditionalFormatting(wb.GetSheetAt(0));
            WriteTemp53691(wb, "fgood");

            wb.Close();
        }

        private void RemoveConditionalFormatting(ISheet sheet)
        {
            ISheetConditionalFormatting cf = sheet.SheetConditionalFormatting;
            for (int j = 0; j < cf.NumConditionalFormattings; j++)
            {
                cf.RemoveConditionalFormatting(j);
            }
        }

        private void WriteTemp53691(IWorkbook wb, String suffix)
        {
            // assert that we can Write/read it in memory
            IWorkbook wbBack = HSSFITestDataProvider.Instance.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);
            wbBack.Close();
        }

        [Test]
        public void test52122()
        {
            IWorkbook workbook = new HSSFWorkbook();
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
            IWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)workbook);
            ISheet sheetBack = wbBack.GetSheetAt(0);
            ISheetConditionalFormatting sheetConditionalFormattingBack = sheetBack.SheetConditionalFormatting;
            Assert.IsNotNull(sheetConditionalFormattingBack);
            IConditionalFormatting formattingBack = sheetConditionalFormattingBack.GetConditionalFormattingAt(0);
            Assert.IsNotNull(formattingBack);
            IConditionalFormattingRule ruleBack = formattingBack.GetRule(0);
            Assert.IsNotNull(ruleBack);
            IPatternFormatting patternFormattingBack1 = ruleBack.PatternFormatting;
            Assert.IsNotNull(patternFormattingBack1);
            Assert.AreEqual(IndexedColors.Blue.Index, patternFormattingBack1.FillBackgroundColor);
            Assert.AreEqual(FillPattern.SolidForeground, patternFormattingBack1.FillPattern);
        }

    }
}