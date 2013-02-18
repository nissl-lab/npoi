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
    using NPOI.SS.Formula.Functions;

    /**
     * Tests for {@link Irr}
     *
     * @author Marcel May
     */
    [TestFixture]
    public class TestIrr
    {
        [Test]
        public void TestIrr1()
        {
            // http://en.wikipedia.org/wiki/Internal_rate_of_return#Example
            double[] incomes = { -4000d, 1200d, 1410d, 1875d, 1050d };
            double irr = Irr.irr(incomes);
            double irrRounded = Math.Round(irr * 1000d) / 1000d;
            Assert.AreEqual(0.143d, irrRounded, "irr");

            // http://www.techonthenet.com/excel/formulas/irr.php
            incomes = new double[] { -7500d, 3000d, 5000d, 1200d, 4000d };
            irr = Irr.irr(incomes);
            irrRounded = Math.Round(irr * 100d) / 100d;
            Assert.AreEqual(0.28d, irrRounded, "irr");

            incomes = new double[] { -10000d, 3400d, 6500d, 1000d };
            irr = Irr.irr(incomes);
            irrRounded = Math.Round(irr * 100d) / 100d;
            Assert.AreEqual(0.05, irrRounded, "irr");

            incomes = new double[] { 100d, -10d, -110d };
            irr = Irr.irr(incomes);
            irrRounded = Math.Round(irr * 100d) / 100d;
            Assert.AreEqual(0.1, irrRounded, "irr");

            incomes = new double[] { -70000d, 12000, 15000 };
            irr = Irr.irr(incomes, -0.1);
            irrRounded = Math.Round(irr * 100d) / 100d;
            Assert.AreEqual(-0.44, irrRounded, "irr");
        }
        [Test]
        public void TestEvaluateInSheet()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);

            row.CreateCell(0).SetCellValue(-4000d);
            row.CreateCell(1).SetCellValue(1200d);
            row.CreateCell(2).SetCellValue(1410d);
            row.CreateCell(3).SetCellValue(1875d);
            row.CreateCell(4).SetCellValue(1050d);

            ICell cell = row.CreateCell(5);
            cell.CellFormula = ("IRR(A1:E1)");

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            fe.ClearAllCachedResultValues();
            fe.EvaluateFormulaCell(cell);
            double res = cell.NumericCellValue;
            Assert.AreEqual(0.143d, Math.Round(res * 1000d) / 1000d);
        }
        [Test]
        public void TestIrrFromSpreadsheet()
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
                ICell cellA = row.GetCell(0);
                try
                {
                    CellValue cv = fe.Evaluate(cellA);
                    assertFormulaResult(cv, cellA);
                }
                catch (Exception e)
                {
                    if (failures.Length > 0) failures.Append('\n');
                    failures.Append("Row[" + (cellA.RowIndex + 1) + "]: " + cellA.CellFormula + " ");
                    failures.Append(e.Message);
                    failureCount++;
                }

                ICell cellC = row.GetCell(2); //IRR-NPV relationship: NPV(IRR(values), values) = 0
                try
                {
                    CellValue cv = fe.Evaluate(cellC);
                    Assert.AreEqual(0, cv.NumberValue, 0.0001);  // should agree within 0.01%
                }
                catch (Exception e)
                {
                    if (failures.Length > 0) failures.Append('\n');
                    failures.Append("Row[" + (cellC.RowIndex + 1) + "]: " + cellC.CellFormula + " ");
                    failures.Append(e.Message);
                    failureCount++;
                }
            }

            if (failures.Length > 0)
            {
                throw new AssertionException(failureCount + " IRR assertions failed:\n" + failures.ToString());
            }

        }

        private static void assertFormulaResult(CellValue cv, ICell cell)
        {
            double actualValue = cv.NumberValue;
            double expectedValue = cell.NumericCellValue; // cached formula result calculated by Excel
            Assert.AreEqual(CellType.Numeric, cv.CellType, "Invalid formula result: " + cv.ToString());
            Assert.AreEqual(expectedValue, actualValue, 1E-4); // should agree within 0.01%
        }
    }

}