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

namespace TestCases.SS.Formula
{
    using System;

    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;

    /**
     * Tests Excel Table expressions (structured references)
     * @see <a href="https://support.office.com/en-us/article/Using-structured-references-with-Excel-tables-F5ED2452-2337-4F71-BED3-C8AE6D2B276E">
     *         Excel Structured Reference Syntax
     *      </a>
     */
    [TestFixture]
    public class TestStructuredReferences
    {

        /**
         * Test the regular expression used in INDIRECT() Evaluation to recognize structured references
         */
        [Test]
        public void TestTableExpressionSyntax()
        {
            Assert.IsTrue(Table.IsStructuredReference.Match("abc[col1]").Success, "Valid structured reference syntax didn't match expression");
            Assert.IsTrue(Table.IsStructuredReference.Match("_abc[col1]").Success, "Valid structured reference syntax didn't match expression");
            Assert.IsTrue(Table.IsStructuredReference.Match("_[col1]").Success, "Valid structured reference syntax didn't match expression");
            Assert.IsTrue(Table.IsStructuredReference.Match("\\[col1]").Success, "Valid structured reference syntax didn't match expression");
            Assert.IsTrue(Table.IsStructuredReference.Match("\\[col1]").Success, "Valid structured reference syntax didn't match expression");
            Assert.IsTrue(Table.IsStructuredReference.Match("\\[#This Row]").Success, "Valid structured reference syntax didn't match expression");
            Assert.IsTrue(Table.IsStructuredReference.Match("\\[ [col1], [col2] ]").Success, "Valid structured reference syntax didn't match expression");

            // can't have a space between the table name and open bracket
            Assert.IsFalse(Table.IsStructuredReference.Match("\\abc [ [col1], [col2] ]").Success, "Invalid structured reference syntax didn't fail expression");
        }

        [Test]
        public void TestTableFormulas()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            try
            {

                IFormulaEvaluator eval = new XSSFFormulaEvaluator(wb);
                XSSFSheet tableSheet = wb.GetSheet("Table") as XSSFSheet;
                XSSFSheet formulaSheet = wb.GetSheet("Formulas") as XSSFSheet;

                Confirm(eval, tableSheet.GetRow(5).GetCell(0), 49);
                Confirm(eval, formulaSheet.GetRow(0).GetCell(0), 209);
                Confirm(eval, formulaSheet.GetRow(1).GetCell(0), "one");

                // test changing a table value, to see if the caches are properly Cleared
                // Issue 59814

                // this test passes before the fix for 59814
                tableSheet.GetRow(1).GetCell(1).SetCellValue("ONEA");
                Confirm(eval, formulaSheet.GetRow(1).GetCell(0), "ONEA");

                // test Adding a row to a table, issue 59814
                IRow newRow = tableSheet.GetRow(7);
                if (newRow == null) newRow = tableSheet.CreateRow(7);
                newRow.CreateCell(0, CellType.Formula).CellFormula = (/*setter*/"\\_Prime.1[[#This Row],[@Number]]*\\_Prime.1[[#This Row],[@Number]]");
                newRow.CreateCell(1, CellType.String).SetCellValue("thirteen");
                newRow.CreateCell(2, CellType.Numeric).SetCellValue(13);

                // update Table
                XSSFTable table = wb.GetTable("\\_Prime.1");
                AreaReference newArea = new AreaReference(table.StartCellReference, new CellReference(table.EndRowIndex + 1, table.EndColIndex));
                String newAreaStr = newArea.FormatAsString();
                table.GetCTTable().@ref = (/*setter*/newAreaStr);
                table.GetCTTable().autoFilter.@ref = (/*setter*/newAreaStr);
                table.UpdateHeaders();
                table.UpdateReferences();

                // these fail before the fix for 59814
                Confirm(eval, tableSheet.GetRow(7).GetCell(0), 13 * 13);
                Confirm(eval, formulaSheet.GetRow(0).GetCell(0), 209 + 13 * 13);

            }
            finally
            {
                wb.Close();
            }
        }

        private static void Confirm(IFormulaEvaluator fe, ICell cell, double expectedResult)
        {
            fe.ClearAllCachedResultValues();
            CellValue cv = fe.Evaluate(cell);
            if (cv.CellType != CellType.Numeric)
            {
                Assert.Fail("expected numeric cell type but got " + cv.FormatAsString());
            }
            Assert.AreEqual(expectedResult, cv.NumberValue, 0.0);
        }

        private static void Confirm(IFormulaEvaluator fe, ICell cell, String expectedResult)
        {
            fe.ClearAllCachedResultValues();
            CellValue cv = fe.Evaluate(cell);
            if (cv.CellType != CellType.String)
            {
                Assert.Fail("expected String cell type but got " + cv.FormatAsString());
            }
            Assert.AreEqual(expectedResult, cv.StringValue);
        }
    }

}