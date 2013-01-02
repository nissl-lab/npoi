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

namespace TestCases.SS.Formula.Functions
{
    using System;
    using System.Text;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;

    /**
     * Tests for {@link Npv}
     *
     * @author Marcel May
     * @see <a href="http://office.microsoft.com/en-us/excel-help/npv-HP005209199.aspx">Excel Help</a>
     */
    [TestFixture]
    public class TestNpv
    {
        [Test]
        public void TestEvaluateInSheetExample2()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);

            sheet.CreateRow(1).CreateCell(0).SetCellValue(0.08d);
            sheet.CreateRow(2).CreateCell(0).SetCellValue(-40000d);
            sheet.CreateRow(3).CreateCell(0).SetCellValue(8000d);
            sheet.CreateRow(4).CreateCell(0).SetCellValue(9200d);
            sheet.CreateRow(5).CreateCell(0).SetCellValue(10000d);
            sheet.CreateRow(6).CreateCell(0).SetCellValue(12000d);
            sheet.CreateRow(7).CreateCell(0).SetCellValue(14500d);

            ICell cell = row.CreateCell(8);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            // Enumeration
            cell.CellFormula = ("NPV(A2, A4,A5,A6,A7,A8)+A3");
            fe.ClearAllCachedResultValues();
            fe.EvaluateFormulaCell(cell);
            double res = cell.NumericCellValue;
            Assert.AreEqual(1922.06d, Math.Round(res * 100d) / 100d);

            // Range
            cell.CellFormula = ("NPV(A2, A4:A8)+A3");

            fe.ClearAllCachedResultValues();
            fe.EvaluateFormulaCell(cell);
            res = cell.NumericCellValue;
            Assert.AreEqual(1922.06d, Math.Round(res * 100d) / 100d);
        }

        /**
         * Evaluate formulas with NPV and compare the result with
         * the cached formula result pre-calculated by Excel
         */
        [Test]
        public void TestNpvFromSpreadsheet()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("IrrNpvTestCaseData.xls");
            ISheet sheet = wb.GetSheet("IRR-NPV");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            StringBuilder failures = new StringBuilder();
            int failureCount = 0;
            // TODO YK: Formulas in rows 16 and 17 operate with ArrayPtg which isn't yet supported
            // FormulaEvaluator as of r1041407 throws "Unexpected ptg class (NPOI.SS.Formula.PTG.ArrayPtg)"
            for (int rownum = 9; rownum <= 15; rownum++)
            {
                IRow row = sheet.GetRow(rownum);
                ICell cellB = row.GetCell(1);
                try
                {
                    CellValue cv = fe.Evaluate(cellB);
                    assertFormulaResult(cv, cellB);
                }
                catch (Exception e)
                {
                    if (failures.Length > 0) failures.Append('\n');
                    failures.Append("Row[" + (cellB.RowIndex + 1) + "]: " + cellB.CellFormula + " ");
                    failures.Append(e.Message);
                    failureCount++;
                }
            }

            if (failures.Length > 0)
            {
                throw new AssertionException(failureCount + " IRR Evaluations failed:\n" + failures.ToString());
            }
        }

        private static void assertFormulaResult(CellValue cv, ICell cell)
        {
            double actualValue = cv.NumberValue;
            double expectedValue = cell.NumericCellValue; // cached formula result calculated by Excel
            Assert.AreEqual(expectedValue, actualValue, 1E-4); // should agree within 0.01%
        }
    }

}