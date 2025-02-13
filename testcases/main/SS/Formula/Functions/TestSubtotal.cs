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
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;
    using NPOI.SS.Formula;



    /**
     * Tests for {@link Subtotal}
     *
     * @author Paul Tomlin
     */
    [TestFixture]
    public class TestSubtotal
    {
        private static int FUNCTION_AVERAGE = 1;
        private static int FUNCTION_COUNT = 2;
        private static int FUNCTION_MAX = 4;
        private static int FUNCTION_MIN = 5;
        private static int FUNCTION_PRODUCT = 6;
        private static int FUNCTION_STDEV = 7;
        private static int FUNCTION_SUM = 9;

        private static double[] TEST_VALUES0 = {
        1, 2,
        3, 4,
        5, 6,
        7, 8,
        9, 10
    };

        private static void ConfirmSubtotal(int function, double expected)
        {
            ValueEval[] values = new ValueEval[TEST_VALUES0.Length];
            for(int i = 0; i < TEST_VALUES0.Length; i++)
            {
                values[i] = new NumberEval(TEST_VALUES0[i]);
            }

            AreaEval arg1 = EvalFactory.CreateAreaEval("C1:D5", values);
            ValueEval[] args = { new NumberEval(function), arg1 };

            ValueEval result = new Subtotal().Evaluate(args, 0, 0);

            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(expected, ((NumberEval) result).NumberValue, 0.0);
        }
        [Test]
        public void TestBasics()
        {
            ConfirmSubtotal(FUNCTION_SUM, 55.0);
            ConfirmSubtotal(FUNCTION_AVERAGE, 5.5);
            ConfirmSubtotal(FUNCTION_COUNT, 10.0);
            ConfirmSubtotal(FUNCTION_MAX, 10.0);
            ConfirmSubtotal(FUNCTION_MIN, 1.0);
            ConfirmSubtotal(FUNCTION_PRODUCT, 3628800.0);
            ConfirmSubtotal(FUNCTION_STDEV, 3.0276503540974917);
        }
        [Test]
        public void TestAvg()
        {
            IWorkbook wb = new HSSFWorkbook();

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.CreateSheet();
            ICell a1 = sh.CreateRow(1).CreateCell(1);
            a1.SetCellValue(1);
            ICell a2 = sh.CreateRow(2).CreateCell(1);
            a2.SetCellValue(3);
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            a3.CellFormula = ("SUBTOTAL(1,B2:B3)");
            ICell a4 = sh.CreateRow(4).CreateCell(1);
            a4.SetCellValue(1);
            ICell a5 = sh.CreateRow(5).CreateCell(1);
            a5.SetCellValue(7);
            ICell a6 = sh.CreateRow(6).CreateCell(1);
            a6.CellFormula = ("SUBTOTAL(1,B2:B6)*2 + 2");
            ICell a7 = sh.CreateRow(7).CreateCell(1);
            a7.CellFormula = ("SUBTOTAL(1,B2:B7)");
            ICell a8 = sh.CreateRow(8).CreateCell(1);
            a8.CellFormula = ("SUBTOTAL(1,B2,B3,B4,B5,B6,B7,B8)");

            fe.EvaluateAll();

            Assert.AreEqual(2.0, a3.NumericCellValue);
            Assert.AreEqual(8.0, a6.NumericCellValue);
            Assert.AreEqual(3.0, a7.NumericCellValue);
            Assert.AreEqual(3.0, a8.NumericCellValue);
        }
        [Test]
        public void TestSum()
        {

            IWorkbook wb = new HSSFWorkbook();

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.CreateSheet();
            ICell a1 = sh.CreateRow(1).CreateCell(1);
            a1.SetCellValue(1);
            ICell a2 = sh.CreateRow(2).CreateCell(1);
            a2.SetCellValue(3);
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            a3.CellFormula = ("SUBTOTAL(9,B2:B3)");
            ICell a4 = sh.CreateRow(4).CreateCell(1);
            a4.SetCellValue(1);
            ICell a5 = sh.CreateRow(5).CreateCell(1);
            a5.SetCellValue(7);
            ICell a6 = sh.CreateRow(6).CreateCell(1);
            a6.CellFormula = ("SUBTOTAL(9,B2:B6)*2 + 2");
            ICell a7 = sh.CreateRow(7).CreateCell(0);
            a7.CellFormula = ("SUBTOTAL(9,B2:B7)");
            ICell a8 = sh.CreateRow(8).CreateCell(1);
            a8.CellFormula = ("SUBTOTAL(9,B2,B3,B4,B5,B6,B7,B8)");

            fe.EvaluateAll();

            Assert.AreEqual(4.0, a3.NumericCellValue);
            Assert.AreEqual(26.0, a6.NumericCellValue);
            Assert.AreEqual(12.0, a7.NumericCellValue);
            Assert.AreEqual(12.0, a8.NumericCellValue);
        }
        [Test]
        public void TestCount()
        {

            IWorkbook wb = new HSSFWorkbook();

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.CreateSheet();
            ICell a1 = sh.CreateRow(1).CreateCell(1);
            a1.SetCellValue(1);
            ICell a2 = sh.CreateRow(2).CreateCell(1);
            a2.SetCellValue(3);
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            a3.CellFormula = ("SUBTOTAL(2,B2:B3)");
            ICell a4 = sh.CreateRow(4).CreateCell(1);
            a4.SetCellValue("POI");                  // A4 is string and not counted
            ICell a5 = sh.CreateRow(5).CreateCell(1); // A5 is blank and not counted

            ICell a6 = sh.CreateRow(6).CreateCell(1);
            a6.CellFormula = ("SUBTOTAL(2,B2:B6)*2 + 2");
            ICell a7 = sh.CreateRow(7).CreateCell(1);
            a7.CellFormula = ("SUBTOTAL(2,B2:B7)");
            ICell a8 = sh.CreateRow(8).CreateCell(1);
            a8.CellFormula = ("SUBTOTAL(2,B2,B3,B4,B5,B6,B7,B8)");

            fe.EvaluateAll();

            Assert.AreEqual(2.0, a3.NumericCellValue);
            Assert.AreEqual(6.0, a6.NumericCellValue);
            Assert.AreEqual(2.0, a7.NumericCellValue);
            Assert.AreEqual(2.0, a8.NumericCellValue);
        }
        [Test]
        public void TestCounta()
        {

            IWorkbook wb = new HSSFWorkbook();

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.CreateSheet();
            ICell a1 = sh.CreateRow(1).CreateCell(1);
            a1.SetCellValue(1);
            ICell a2 = sh.CreateRow(2).CreateCell(1);
            a2.SetCellValue(3);
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            a3.CellFormula = ("SUBTOTAL(3,B2:B3)");
            ICell a4 = sh.CreateRow(4).CreateCell(1);
            a4.SetCellValue("POI");                  // A4 is string and not counted
            ICell a5 = sh.CreateRow(5).CreateCell(1); // A5 is blank and not counted

            ICell a6 = sh.CreateRow(6).CreateCell(1);
            a6.CellFormula = ("SUBTOTAL(3,B2:B6)*2 + 2");
            ICell a7 = sh.CreateRow(7).CreateCell(1);
            a7.CellFormula = ("SUBTOTAL(3,B2:B7)");
            ICell a8 = sh.CreateRow(8).CreateCell(1);
            a8.CellFormula = ("SUBTOTAL(3,B2,B3,B4,B5,B6,B7,B8)");

            fe.EvaluateAll();

            Assert.AreEqual(2.0, a3.NumericCellValue);
            Assert.AreEqual(8.0, a6.NumericCellValue);
            Assert.AreEqual(3.0, a7.NumericCellValue);
            Assert.AreEqual(3.0, a8.NumericCellValue);
        }
        [Test]
        public void TestMax()
        {

            IWorkbook wb = new HSSFWorkbook();

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.CreateSheet();
            ICell a1 = sh.CreateRow(1).CreateCell(1);
            a1.SetCellValue(1);
            ICell a2 = sh.CreateRow(2).CreateCell(1);
            a2.SetCellValue(3);
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            a3.CellFormula = ("SUBTOTAL(4,B2:B3)");
            ICell a4 = sh.CreateRow(4).CreateCell(1);
            a4.SetCellValue(1);
            ICell a5 = sh.CreateRow(5).CreateCell(1);
            a5.SetCellValue(7);
            ICell a6 = sh.CreateRow(6).CreateCell(1);
            a6.CellFormula = ("SUBTOTAL(4,B2:B6)*2 + 2");
            ICell a7 = sh.CreateRow(7).CreateCell(1);
            a7.CellFormula = ("SUBTOTAL(4,B2:B7)");
            ICell a8 = sh.CreateRow(8).CreateCell(1);
            a8.CellFormula = ("SUBTOTAL(4,B2,B3,B4,B5,B6,B7,B8)");

            fe.EvaluateAll();

            Assert.AreEqual(3.0, a3.NumericCellValue);
            Assert.AreEqual(16.0, a6.NumericCellValue);
            Assert.AreEqual(7.0, a7.NumericCellValue);
            Assert.AreEqual(7.0, a8.NumericCellValue);
        }
        [Test]
        public void TestMin()
        {

            IWorkbook wb = new HSSFWorkbook();

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.CreateSheet();
            ICell a1 = sh.CreateRow(1).CreateCell(1);
            a1.SetCellValue(1);
            ICell a2 = sh.CreateRow(2).CreateCell(1);
            a2.SetCellValue(3);
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            a3.CellFormula = ("SUBTOTAL(5,B2:B3)");
            ICell a4 = sh.CreateRow(4).CreateCell(1);
            a4.SetCellValue(1);
            ICell a5 = sh.CreateRow(5).CreateCell(1);
            a5.SetCellValue(7);
            ICell a6 = sh.CreateRow(6).CreateCell(1);
            a6.CellFormula = ("SUBTOTAL(5,B2:B6)*2 + 2");
            ICell a7 = sh.CreateRow(7).CreateCell(1);
            a7.CellFormula = ("SUBTOTAL(5,B2:B7)");
            ICell a8 = sh.CreateRow(8).CreateCell(1);
            a8.CellFormula = ("SUBTOTAL(5,B2,B3,B4,B5,B6,B7,B8)");

            fe.EvaluateAll();

            Assert.AreEqual(1.0, a3.NumericCellValue);
            Assert.AreEqual(4.0, a6.NumericCellValue);
            Assert.AreEqual(1.0, a7.NumericCellValue);
            Assert.AreEqual(1.0, a8.NumericCellValue);
        }
        [Test]
        public void TestStdev()
        {

            IWorkbook wb = new HSSFWorkbook();

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.CreateSheet();
            ICell a1 = sh.CreateRow(1).CreateCell(1);
            a1.SetCellValue(1);
            ICell a2 = sh.CreateRow(2).CreateCell(1);
            a2.SetCellValue(3);
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            a3.CellFormula = ("SUBTOTAL(7,B2:B3)");
            ICell a4 = sh.CreateRow(4).CreateCell(1);
            a4.SetCellValue(1);
            ICell a5 = sh.CreateRow(5).CreateCell(1);
            a5.SetCellValue(7);
            ICell a6 = sh.CreateRow(6).CreateCell(1);
            a6.CellFormula = ("SUBTOTAL(7,B2:B6)*2 + 2");
            ICell a7 = sh.CreateRow(7).CreateCell(1);
            a7.CellFormula = ("SUBTOTAL(7,B2:B7)");
            ICell a8 = sh.CreateRow(8).CreateCell(1);
            a8.CellFormula = ("SUBTOTAL(7,B2,B3,B4,B5,B6,B7,B8)");

            fe.EvaluateAll();

            Assert.AreEqual(1.41421, a3.NumericCellValue, 0.00001);
            Assert.AreEqual(7.65685, a6.NumericCellValue, 0.00001);
            Assert.AreEqual(2.82842, a7.NumericCellValue, 0.00001);
            Assert.AreEqual(2.82842, a8.NumericCellValue, 0.00001);
        }

        [Test]
        public void TestStdevp()
        {
            using(IWorkbook wb = new HSSFWorkbook())
            {
                IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

                ISheet sh = wb.CreateSheet();
                ICell a1 = sh.CreateRow(1).CreateCell(1);
                a1.SetCellValue(1);
                ICell a2 = sh.CreateRow(2).CreateCell(1);
                a2.SetCellValue(3);
                ICell a3 = sh.CreateRow(3).CreateCell(1);
                a3.SetCellFormula("SUBTOTAL(8,B2:B3)");
                ICell a4 = sh.CreateRow(4).CreateCell(1);
                a4.SetCellValue(1);
                ICell a5 = sh.CreateRow(5).CreateCell(1);
                a5.SetCellValue(7);
                ICell a6 = sh.CreateRow(6).CreateCell(1);
                a6.SetCellFormula("SUBTOTAL(8,B2:B6)*2 + 2");
                ICell a7 = sh.CreateRow(7).CreateCell(1);
                a7.SetCellFormula("SUBTOTAL(8,B2:B7)");
                ICell a8 = sh.CreateRow(8).CreateCell(1);
                a8.SetCellFormula("SUBTOTAL(8,B2,B3,B4,B5,B6,B7,B8)");

                fe.EvaluateAll();

                Assert.AreEqual(1.0, a3.NumericCellValue, 0.00001);
                Assert.AreEqual(6.898979, a6.NumericCellValue, 0.00001);
                Assert.AreEqual(2.44949, a7.NumericCellValue, 0.00001);
                Assert.AreEqual(2.44949, a8.NumericCellValue, 0.00001);
            }
        }

        [Test]
        public void TestVar()
        {


            using(IWorkbook wb = new HSSFWorkbook())
            {
                IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

                ISheet sh = wb.CreateSheet();
                ICell a1 = sh.CreateRow(1).CreateCell(1);
                a1.SetCellValue(1);
                ICell a2 = sh.CreateRow(2).CreateCell(1);
                a2.SetCellValue(3);
                ICell a3 = sh.CreateRow(3).CreateCell(1);
                a3.SetCellFormula("SUBTOTAL(10,B2:B3)");
                ICell a4 = sh.CreateRow(4).CreateCell(1);
                a4.SetCellValue(1);
                ICell a5 = sh.CreateRow(5).CreateCell(1);
                a5.SetCellValue(7);
                ICell a6 = sh.CreateRow(6).CreateCell(1);
                a6.SetCellFormula("SUBTOTAL(10,B2:B6)*2 + 2");
                ICell a7 = sh.CreateRow(7).CreateCell(1);
                a7.SetCellFormula("SUBTOTAL(10,B2:B7)");
                ICell a8 = sh.CreateRow(8).CreateCell(1);
                a8.SetCellFormula("SUBTOTAL(10,B2,B3,B4,B5,B6,B7,B8)");

                fe.EvaluateAll();

                Assert.AreEqual(2.0, a3.NumericCellValue);
                Assert.AreEqual(18.0, a6.NumericCellValue);
                Assert.AreEqual(8.0, a7.NumericCellValue);
                Assert.AreEqual(8.0, a8.NumericCellValue);
            }
        }

        [Test]
        public void TestVarp()
        {


            using(IWorkbook wb = new HSSFWorkbook())
            {
                IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

                ISheet sh = wb.CreateSheet();
                ICell a1 = sh.CreateRow(1).CreateCell(1);
                a1.SetCellValue(1);
                ICell a2 = sh.CreateRow(2).CreateCell(1);
                a2.SetCellValue(3);
                ICell a3 = sh.CreateRow(3).CreateCell(1);
                a3.SetCellFormula("SUBTOTAL(11,B2:B3)");
                ICell a4 = sh.CreateRow(4).CreateCell(1);
                a4.SetCellValue(1);
                ICell a5 = sh.CreateRow(5).CreateCell(1);
                a5.SetCellValue(7);
                ICell a6 = sh.CreateRow(6).CreateCell(1);
                a6.SetCellFormula("SUBTOTAL(11,B2:B6)*2 + 2");
                ICell a7 = sh.CreateRow(7).CreateCell(1);
                a7.SetCellFormula("SUBTOTAL(11,B2:B7)");
                ICell a8 = sh.CreateRow(8).CreateCell(1);
                a8.SetCellFormula("SUBTOTAL(11,B2,B3,B4,B5,B6,B7,B8)");

                fe.EvaluateAll();

                Assert.AreEqual(1.0, a3.NumericCellValue);
                Assert.AreEqual(14.0, a6.NumericCellValue);
                Assert.AreEqual(6.0, a7.NumericCellValue);
                Assert.AreEqual(6.0, a8.NumericCellValue);
            }
        }

        [Test]
        public void Test50209()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sh = wb.CreateSheet();
            ICell a1 = sh.CreateRow(1).CreateCell(1);
            a1.SetCellValue(1);
            ICell a2 = sh.CreateRow(2).CreateCell(1);
            a2.CellFormula = ("SUBTOTAL(9,B2)");
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            a3.CellFormula = ("SUBTOTAL(9,B2:B3)");

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            fe.EvaluateAll();
            Assert.AreEqual(1.0, a2.NumericCellValue);
            Assert.AreEqual(1.0, a3.NumericCellValue);
        }

        private static void ConfirmExpectedResult(IFormulaEvaluator Evaluator, String msg, ICell cell, double expected)
        {

            CellValue value = Evaluator.Evaluate(cell);
            if(value.ErrorValue != 0)
                throw new Exception(msg + ": " + value.FormatAsString());
            Assert.AreEqual(expected, value.NumberValue, msg);
        }
        [Test]
        public void TestFunctionsFromTestSpreadsheet()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("SubtotalsNested.xls");
            ISheet sheet = workbook.GetSheetAt(0);
            IFormulaEvaluator evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

            Assert.AreEqual(10.0, sheet.GetRow(1).GetCell(1).NumericCellValue, "B2");
            Assert.AreEqual(20.0, sheet.GetRow(2).GetCell(1).NumericCellValue, "B3");

            //Test simple subtotal over one area
            ICell cellA3 = sheet.GetRow(3).GetCell(1);
            ConfirmExpectedResult(evaluator, "B4", cellA3, 30.0);

            //Test existence of the second area
            Assert.IsNotNull(sheet.GetRow(1).GetCell(2), "C2 must not be null");
            Assert.AreEqual(7.0, sheet.GetRow(1).GetCell(2).NumericCellValue, "C2");

            ICell cellC1 = sheet.GetRow(1).GetCell(3);
            ICell cellC2 = sheet.GetRow(2).GetCell(3);
            ICell cellC3 = sheet.GetRow(3).GetCell(3);

            //Test Functions SUM, COUNT and COUNTA calculation of SUBTOTAL
            //a) areas A and B are used
            //b) first 2 subtotals don't consider the value of nested subtotal in A3
            ConfirmExpectedResult(evaluator, "SUBTOTAL(SUM;B2:B8;C2:C8)", cellC1, 37.0);
            ConfirmExpectedResult(evaluator, "SUBTOTAL(COUNT;B2:B8;C2:C8)", cellC2, 3.0);
            ConfirmExpectedResult(evaluator, "SUBTOTAL(COUNTA;B2:B8;C2:C8)", cellC3, 5.0);
        }

        [Test]
        public void TestUnimplemented()
        {
            IWorkbook wb = new HSSFWorkbook();
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            ISheet sh = wb.CreateSheet();
            ICell a3 = sh.CreateRow(3).CreateCell(1);
            
            a3.CellFormula = ("SUBTOTAL(0,B2:B3)");
            fe.EvaluateAll();
            Assert.AreEqual(FormulaError.VALUE.Code, a3.ErrorCellValue);
            try
            {
                a3.CellFormula = ("SUBTOTAL(9)");
                Assert.Fail("Should catch an exception here");
            }
            catch(FormulaParseException)
            {
                // expected here
            }
            try
            {
                a3.CellFormula = ("SUBTOTAL()");
                Assert.Fail("Should catch an exception here");
            }
            catch(FormulaParseException)
            {
                // expected here
            }
            Subtotal subtotal = new Subtotal();
            Assert.AreEqual(ErrorEval.VALUE_INVALID, subtotal.Evaluate(new ValueEval[] { }, 0, 0));
        }

    }

}