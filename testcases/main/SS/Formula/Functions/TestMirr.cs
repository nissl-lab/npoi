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
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NUnit.Framework;
    using TestCases.HSSF;
    using NPOI.SS.Formula.Functions;

    /**
     * Tests for {@link NPOI.SS.Formula.Functions.Mirr}
     *
     * @author Carlos Delgado (carlos dot del dot est at gmail dot com)
     * @author Cédric Walter (cedric dot walter at gmail dot com)
     * @see {@link NPOI.SS.Formula.Functions.TestIrr}
     */
    [TestFixture]
    public class TestMirr
    {

        [Test]
        public void TestMirrBasic()
        {
            Mirr mirr = new Mirr();
            double mirrValue;

            double financeRate = 0.12;
            double reinvestRate = 0.1;
            double[] values = { -120000d, 39000d, 30000d, 21000d, 37000d, 46000d, reinvestRate, financeRate };
            try
            {
                mirrValue = mirr.Evaluate(values);
            }
            catch (EvaluationException e)
            {
                throw new AssertFailedException("MIRR should not failed with these parameters" + e);
            }
            Assert.AreEqual(0.126094130366, mirrValue, 0.0000000001, "mirr");

            reinvestRate = 0.05;
            financeRate = 0.08;
            values = new double[] { -7500d, 3000d, 5000d, 1200d, 4000d, reinvestRate, financeRate };
            try
            {
                mirrValue = mirr.Evaluate(values);
            }
            catch (EvaluationException e)
            {
                throw new AssertFailedException("MIRR should not failed with these parameters" + e);
            }
            Assert.AreEqual(0.18736225093, mirrValue, 0.0000000001, "mirr");

            reinvestRate = 0.065;
            financeRate = 0.1;
            values = new double[] { -10000, 3400d, 6500d, 1000d, reinvestRate, financeRate };
            try
            {
                mirrValue = mirr.Evaluate(values);
            }
            catch (EvaluationException e)
            {
                throw new AssertFailedException("MIRR should not failed with these parameters" + e);
            }
            Assert.AreEqual(0.07039493966, mirrValue, 0.0000000001, "mirr");

            reinvestRate = 0.07;
            financeRate = 0.01;
            values = new double[] { -10000d, -3400d, -6500d, -1000d, reinvestRate, financeRate };
            try
            {
                mirrValue = mirr.Evaluate(values);
            }
            catch (EvaluationException e)
            {
                throw new AssertFailedException("MIRR should not failed with these parameters" + e);
            }
            Assert.AreEqual(-1, mirrValue, 0.0, "mirr");

        }

        [Test]
        public void TestMirrErrors_expectDIV0()
        {
            Mirr mirr = new Mirr();

            double reinvestRate = 0.05;
            double financeRate = 0.08;
            double[] incomes = { 120000d, 39000d, 30000d, 21000d, 37000d, 46000d, reinvestRate, financeRate };
            try
            {
                mirr.Evaluate(incomes);
            }
            catch (EvaluationException e)
            {
                Assert.AreEqual(ErrorEval.DIV_ZERO, e.GetErrorEval());
                return;
            }
            throw new AssertFailedException("MIRR should failed with all these positives values");
        }


        [Test]
        public void TestEvaluateInSheet()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);

            row.CreateCell(0).SetCellValue(-7500d);
            row.CreateCell(1).SetCellValue(3000d);
            row.CreateCell(2).SetCellValue(5000d);
            row.CreateCell(3).SetCellValue(1200d);
            row.CreateCell(4).SetCellValue(4000d);

            row.CreateCell(5).SetCellValue(0.05d);
            row.CreateCell(6).SetCellValue(0.08d);

            ICell cell = row.CreateCell(7);
            cell.CellFormula = (/*setter*/"MIRR(A1:E1, F1, G1)");

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            fe.ClearAllCachedResultValues();
            fe.EvaluateFormulaCell(cell);
            double res = cell.NumericCellValue;
            Assert.AreEqual(0.18736225093, res, 0.00000001);
        }

        [Test]
        public void TestMirrFromSpreadsheet()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("mirrTest.xls");
            ISheet sheet = wb.GetSheet("Mirr");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            StringBuilder failures = new StringBuilder();
            int failureCount = 0;
            int[] resultRows = { 9, 19, 29, 45 };

            foreach (int rowNum in resultRows)
            {
                IRow row1 = sheet.GetRow(rowNum);
                ICell cellA1 = row1.GetCell(0);
                try
                {
                    CellValue cv1 = fe.Evaluate(cellA1);
                    assertFormulaResult(cv1, cellA1);
                }
                catch (Exception e)
                {
                    if (failures.Length > 0) failures.Append('\n');
                    failures.Append("Row[").Append(cellA1.RowIndex + 1).Append("]: ").Append(cellA1.CellFormula).Append(" ");
                    failures.Append(e.Message);
                    failureCount++;
                }
            }

            IRow row = sheet.GetRow(37);
            ICell cellA = row.GetCell(0);
            CellValue cv = fe.Evaluate(cellA);
            Assert.AreEqual(ErrorEval.DIV_ZERO.ErrorCode, cv.ErrorValue);

            if (failures.Length > 0)
            {
                throw new AssertFailedException(failureCount + " IRR assertions failed:\n" + failures.ToString());
            }

        }

        private static void assertFormulaResult(CellValue cv, ICell cell)
        {
            double actualValue = cv.NumberValue;
            double expectedValue = cell.NumericCellValue; // cached formula result calculated by Excel
            Assert.AreEqual(CellType.Numeric, cv.CellType, "Invalid formula result: " + cv.ToString());
            Assert.AreEqual(expectedValue, actualValue, 1E-8);
        }
    }

}