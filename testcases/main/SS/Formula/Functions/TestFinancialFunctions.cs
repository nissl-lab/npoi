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

namespace TestCases.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Tests for financial functions backed by ExcelFinancialFunctions.
    /// Expected values verified against Excel.
    /// </summary>
    [TestFixture]
    public class TestFinancialFunctions
    {
        private static CellValue EvalFormula(string formula)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.CellFormula = formula;
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            return fe.Evaluate(cell);
        }

        [Test]
        public void TestFv()
        {
            // FV(rate, nper, pmt, pv, type)
            CellValue cv = EvalFormula("FV(0.05/12, 10*12, -100, -1000, 0)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(17175.24, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestPv()
        {
            // PV(rate, nper, pmt, fv, type)
            CellValue cv = EvalFormula("PV(0.1/12, 24, -500, 0, 0)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(10835.43, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestPmt()
        {
            // PMT(rate, nper, pv, fv, type)
            CellValue cv = EvalFormula("PMT(0.08/12, 10, 10000, 0, 0)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(-1037.03, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestNper()
        {
            // NPER(rate, pmt, pv, fv, type)
            CellValue cv = EvalFormula("NPER(0.05/12, -100, -1000, 10000, 0)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(73.95, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestRate()
        {
            // RATE(nper, pmt, pv, fv, type, guess)
            CellValue cv = EvalFormula("RATE(4*12, -200, 8000, 0, 0, 0.1)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(0.00770, cv.NumberValue, 0.0001);
        }

        [Test]
        public void TestIpmt()
        {
            // IPMT(rate, per, nper, pv, fv, type)
            CellValue cv = EvalFormula("IPMT(0.1/12, 1, 3*12, 8000, 0, 0)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(-66.67, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestPpmt()
        {
            // PPMT(rate, per, nper, pv, fv, type)
            CellValue cv = EvalFormula("PPMT(0.1/12, 1, 3*12, 8000, 0, 0)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(-191.47, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestIrr()
        {
            // IRR(values, guess)
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(-70000);
            row.CreateCell(1).SetCellValue(12000);
            row.CreateCell(2).SetCellValue(15000);
            row.CreateCell(3).SetCellValue(18000);
            row.CreateCell(4).SetCellValue(21000);
            ICell cell = row.CreateCell(5);
            cell.CellFormula = "IRR(A1:E1, -0.1)";
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(-0.0213, cv.NumberValue, 0.0001);
        }

        [Test]
        public void TestMirr()
        {
            // MIRR(values, finance_rate, reinvest_rate) via formula evaluation.
            // In NPOI's MIRR implementation, the MultiOperandNumericFunction collects all
            // arguments into a flat array. The code reads financeRate from values[n-1] (last)
            // and reinvestRate from values[n-2] (second-to-last). The EFF library's Mirr
            // function has its rate parameters in the opposite order to Excel's documentation,
            // so they are passed as Financial.Mirr(values, reinvestRate, financeRate).
            //
            // For formula MIRR(A1:F1, X, Y):
            //   values[n-2] = X, values[n-1] = Y
            //   financeRate = Y, reinvestRate = X
            //   call: Financial.Mirr(cashflows, X, Y)
            //
            // To get 0.12609... (matching the TestMirr unit tests), use MIRR(A1:F1, 0.1, 0.12)
            // so that the extracted financeRate=0.12, reinvestRate=0.1, and the EFF call is
            // Financial.Mirr(cashflows, 0.1, 0.12) = 0.12609...
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(-120000);
            row.CreateCell(1).SetCellValue(39000);
            row.CreateCell(2).SetCellValue(30000);
            row.CreateCell(3).SetCellValue(21000);
            row.CreateCell(4).SetCellValue(37000);
            row.CreateCell(5).SetCellValue(46000);
            ICell cell = row.CreateCell(6);
            cell.CellFormula = "MIRR(A1:F1, 0.1, 0.12)";
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(0.12609413, cv.NumberValue, 0.0000001);
        }

        [Test]
        public void TestSln()
        {
            // SLN(cost, salvage, life) - straight-line depreciation
            CellValue cv = EvalFormula("SLN(30000, 7500, 10)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(2250.0, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestSyd()
        {
            // SYD(cost, salvage, life, per) - sum of years digits depreciation
            CellValue cv = EvalFormula("SYD(30000, 7500, 10, 1)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(4090.91, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestDdb()
        {
            // DDB(cost, salvage, life, period, factor)
            CellValue cv = EvalFormula("DDB(2400, 300, 10*365, 1, 2)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(1.32, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestDb()
        {
            // DB(cost, salvage, life, period, month)
            CellValue cv = EvalFormula("DB(1000000, 100000, 6, 1, 7)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(186083.33, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestVdb()
        {
            // VDB(cost, salvage, life, start_period, end_period, factor, no_switch)
            CellValue cv = EvalFormula("VDB(2400, 300, 10*365, 0, 1, 2, TRUE)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(1.32, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestEffect()
        {
            // EFFECT(nominal_rate, npery)
            CellValue cv = EvalFormula("EFFECT(0.0525, 4)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(0.053543, cv.NumberValue, 0.000001);
        }

        [Test]
        public void TestNominal()
        {
            // NOMINAL(effect_rate, npery)
            CellValue cv = EvalFormula("NOMINAL(0.053543, 4)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(0.0525, cv.NumberValue, 0.0001);
        }

        [Test]
        public void TestXirr()
        {
            // XIRR(values, dates, guess)
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            IRow row2 = sheet.CreateRow(1);
            // Cash flows
            row.CreateCell(0).SetCellValue(-10000);
            row.CreateCell(1).SetCellValue(2750);
            row.CreateCell(2).SetCellValue(4250);
            row.CreateCell(3).SetCellValue(3250);
            row.CreateCell(4).SetCellValue(2750);
            // Dates as Excel serial numbers: 2008-01-01=39448, 2008-03-01=39508, 2008-10-30=39751, 2009-02-15=39859, 2009-04-01=39904
            row2.CreateCell(0).SetCellValue(39448);
            row2.CreateCell(1).SetCellValue(39508);
            row2.CreateCell(2).SetCellValue(39751);
            row2.CreateCell(3).SetCellValue(39859);
            row2.CreateCell(4).SetCellValue(39904);
            ICell cell = row.CreateCell(5);
            cell.CellFormula = "XIRR(A1:E1, A2:E2, 0.1)";
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(0.3734, cv.NumberValue, 0.001);
        }

        [Test]
        public void TestXnpv()
        {
            // XNPV(rate, values, dates)
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            IRow row2 = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue(-10000);
            row.CreateCell(1).SetCellValue(2750);
            row.CreateCell(2).SetCellValue(4250);
            row.CreateCell(3).SetCellValue(3250);
            row.CreateCell(4).SetCellValue(2750);
            row2.CreateCell(0).SetCellValue(39448);
            row2.CreateCell(1).SetCellValue(39508);
            row2.CreateCell(2).SetCellValue(39751);
            row2.CreateCell(3).SetCellValue(39859);
            row2.CreateCell(4).SetCellValue(39904);
            ICell cell = row.CreateCell(5);
            cell.CellFormula = "XNPV(0.09, A1:E1, A2:E2)";
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(2086.65, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestCumipmt()
        {
            // CUMIPMT(rate, nper, pv, start_period, end_period, type)
            CellValue cv = EvalFormula("CUMIPMT(0.1/12, 3*12, 75000, 1, 5, 0)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(-2974.16, cv.NumberValue, 0.01);
        }

        [Test]
        public void TestCumprinc()
        {
            // CUMPRINC(rate, nper, pv, start_period, end_period, type)
            CellValue cv = EvalFormula("CUMPRINC(0.1/12, 3*12, 75000, 1, 5, 0)");
            ClassicAssert.AreEqual(CellType.Numeric, cv.CellType);
            ClassicAssert.AreEqual(-9126.03, cv.NumberValue, 0.01);
        }
    }
}
