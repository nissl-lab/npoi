/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
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

namespace TestCases.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NUnit.Framework;
    using TestCases.HSSF;

    /**
     * Test cases for SUMIFS()
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestSumifs
    {
        private static OperationEvaluationContext EC = new OperationEvaluationContext(null, null, 0, 1, 0, null);

        private static ValueEval InvokeSumifs(ValueEval[] args, OperationEvaluationContext ec)
        {
            return new Sumifs().Evaluate(args, EC);
        }

        private static void ConfirmDouble(double expected, ValueEval actualEval)
        {
            if (!(actualEval is NumericValueEval))
            {
                throw new AssertionException("Expected numeric result");
            }
            NumericValueEval nve = (NumericValueEval)actualEval;
            Assert.AreEqual(expected, nve.NumberValue, 0);
        }

        private static void Confirm(double expectedResult, ValueEval[] args)
        {
            ConfirmDouble(expectedResult, InvokeSumifs(args, EC));
        }

        /**
         *  Example 1 from
         *  http://office.microsoft.com/en-us/excel-help/sumifs-function-HA010047504.aspx
         */
        [Test]
        public void TestExample1()
        {
            // mimic test sample from http://office.microsoft.com/en-us/excel-help/sumifs-function-HA010047504.aspx
            ValueEval[] a2a9 = new ValueEval[]
            {
                new NumberEval(5),
                new NumberEval(4),
                new NumberEval(15),
                new NumberEval(3),
                new NumberEval(22),
                new NumberEval(12),
                new NumberEval(10),
                new NumberEval(33)
            };

            ValueEval[] b2b9 = new ValueEval[]
            {
                new StringEval("Apples"),
                new StringEval("Apples"),
                new StringEval("Artichokes"),
                new StringEval("Artichokes"),
                new StringEval("Bananas"),
                new StringEval("Bananas"),
                new StringEval("Carrots"),
                new StringEval("Carrots"),
            };

            ValueEval[] c2c9 = new ValueEval[]
            {
                new NumberEval(1),
                new NumberEval(2),
                new NumberEval(1),
                new NumberEval(2),
                new NumberEval(1),
                new NumberEval(2),
                new NumberEval(1),
                new NumberEval(2)
            };

            // "=SUMIFS(A2:A9, B2:B9, "=A*", C2:C9, 1)"
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A9", a2a9),
                EvalFactory.CreateAreaEval("B2:B9", b2b9),
                new StringEval("A*"),
                EvalFactory.CreateAreaEval("C2:C9", c2c9),
                new NumberEval(1),
            };
            Confirm(20.0, args);

            // "=SUMIFS(A2:A9, B2:B9, "<>Bananas", C2:C9, 1)"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A9", a2a9),
                EvalFactory.CreateAreaEval("B2:B9", b2b9),
                new StringEval("<>Bananas"),
                EvalFactory.CreateAreaEval("C2:C9", c2c9),
                new NumberEval(1),
            };
            Confirm(30.0, args);

            // a test case that returns ErrorEval.VALUE_INVALID :
            // the dimensions of the first and second criteria ranges are different
            // "=SUMIFS(A2:A9, B2:B8, "<>Bananas", C2:C9, 1)"
            args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("A2:A9", a2a9),
                EvalFactory.CreateAreaEval("B2:B8", new ValueEval[]
                {
                        new StringEval("Apples"),
                        new StringEval("Apples"),
                        new StringEval("Artichokes"),
                        new StringEval("Artichokes"),
                        new StringEval("Bananas"),
                        new StringEval("Bananas"),
                        new StringEval("Carrots"),
                }),
                new StringEval("<>Bananas"),
                EvalFactory.CreateAreaEval("C2:C9", c2c9),
                new NumberEval(1),
            };
            Assert.AreEqual(ErrorEval.VALUE_INVALID, InvokeSumifs(args, EC));
        }

        /**
         *  Example 2 from
         *  http://office.microsoft.com/en-us/excel-help/sumifs-function-HA010047504.aspx
         */
        [Test]
        public void TestExample2()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            ValueEval[] b2e2 = new ValueEval[]
            {
                new NumberEval(100),
                new NumberEval(390),
                new NumberEval(8321),
                new NumberEval(500)
            };

            // 1%	0.5%	3%	4%
            ValueEval[] b3e3 = new ValueEval[]
            {
                new NumberEval(0.01),
                new NumberEval(0.005),
                new NumberEval(0.03),
                new NumberEval(0.04)
            };

            // 1%	1.3%	2.1%	2%
            ValueEval[] b4e4 = new ValueEval[]
            {
                new NumberEval(0.01),
                new NumberEval(0.013),
                new NumberEval(0.021),
                new NumberEval(0.02)
            };

            // 0.5%	3%	1%	4%
            ValueEval[] b5e5 = new ValueEval[]
            {
                new NumberEval(0.005),
                new NumberEval(0.03),
                new NumberEval(0.01),
                new NumberEval(0.04)
            };

            // "=SUMIFS(B2:E2, B3:E3, ">3%", B4:E4, ">=2%")"
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("B2:E2", b2e2),
                EvalFactory.CreateAreaEval("B3:E3", b3e3),
                new StringEval(">0.03"), // 3% in the MSFT example
                EvalFactory.CreateAreaEval("B4:E4", b4e4),
                new StringEval(">=0.02"),   // 2% in the MSFT example
            };

            Confirm(500.0, args);
        }

        /**
         *  Example 3 from
         *  http://office.microsoft.com/en-us/excel-help/sumifs-function-HA010047504.aspx
         */
        [Test]
        public void TestExample3()
        {
            //3.3	0.8	5.5	5.5
            ValueEval[] b2e2 = new ValueEval[]
            {
                new NumberEval(3.3),
                new NumberEval(0.8),
                new NumberEval(5.5),
                new NumberEval(5.5)
            };

            // 55	39	39	57.5
            ValueEval[] b3e3 = new ValueEval[]
            {
                new NumberEval(55),
                new NumberEval(39),
                new NumberEval(39),
                new NumberEval(57.5)
            };

            // 6.5	19.5	6	6.5
            ValueEval[] b4e4 = new ValueEval[]
            {
                new NumberEval(6.5),
                new NumberEval(19.5),
                new NumberEval(6),
                new NumberEval(6.5)
            };

            // "=SUMIFS(B2:E2, B3:E3, ">=40", B4:E4, "<10")"
            ValueEval[] args = new ValueEval[]
            {
                EvalFactory.CreateAreaEval("B2:E2", b2e2),
                EvalFactory.CreateAreaEval("B3:E3", b3e3),
                new StringEval(">=40"),
                EvalFactory.CreateAreaEval("B4:E4", b4e4),
                new StringEval("<10"),
            };
            Confirm(8.8, args);
        }

        /**
         *  Example 5 from
         *  http://office.microsoft.com/en-us/excel-help/sumifs-function-HA010047504.aspx
         *
         *  Criteria entered as reference and by using wildcard characters
         */
        [Test]
        public void TestFromFile()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("sumifs.xls");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            HSSFSheet example1 = (HSSFSheet)wb.GetSheet("Example 1");
            HSSFCell ex1cell1 = (HSSFCell)example1.GetRow(10).GetCell(2);
            fe.Evaluate(ex1cell1);
            Assert.AreEqual(20.0, ex1cell1.NumericCellValue);
            HSSFCell ex1cell2 = (HSSFCell)example1.GetRow(11).GetCell(2);
            fe.Evaluate(ex1cell2);
            Assert.AreEqual(30.0, ex1cell2.NumericCellValue);

            HSSFSheet example2 = (HSSFSheet)wb.GetSheet("Example 2");
            HSSFCell ex2cell1 = (HSSFCell)example2.GetRow(6).GetCell(2);
            fe.Evaluate(ex2cell1);
            Assert.AreEqual(500.0, ex2cell1.NumericCellValue);
            HSSFCell ex2cell2 = (HSSFCell)example2.GetRow(7).GetCell(2);
            fe.Evaluate(ex2cell2);
            Assert.AreEqual(8711.0, ex2cell2.NumericCellValue);

            HSSFSheet example3 = (HSSFSheet)wb.GetSheet("Example 3");
            HSSFCell ex3cell = (HSSFCell)example3.GetRow(5).GetCell(2);
            fe.Evaluate(ex3cell);
            Assert.AreEqual(8, 8, ex3cell.NumericCellValue);

            HSSFSheet example4 = (HSSFSheet)wb.GetSheet("Example 4");
            HSSFCell ex4cell = (HSSFCell)example4.GetRow(8).GetCell(2);
            fe.Evaluate(ex4cell);
            Assert.AreEqual(3.5, ex4cell.NumericCellValue);

            HSSFSheet example5 = (HSSFSheet)wb.GetSheet("Example 5");
            HSSFCell ex5cell = (HSSFCell)example5.GetRow(8).GetCell(2);
            fe.Evaluate(ex5cell);
            Assert.AreEqual(625000.0, ex5cell.NumericCellValue);
        }
        [Test]
        public void TestBug56655()
        {
            ValueEval[] a2a9 = new ValueEval[] {
                new NumberEval(5),
                new NumberEval(4),
                new NumberEval(15),
                new NumberEval(3),
                new NumberEval(22),
                new NumberEval(12),
                new NumberEval(10),
                new NumberEval(33)
        };

            ValueEval[] args = new ValueEval[]{
                EvalFactory.CreateAreaEval("A2:A9", a2a9),
                ErrorEval.VALUE_INVALID,
                new StringEval("A*"),
        };

            ValueEval result = InvokeSumifs(args, EC);
            Assert.IsTrue(result is ErrorEval, "Expect to have an error when an input is an invalid value, but had: " + result.GetType());

            args = new ValueEval[]{
                EvalFactory.CreateAreaEval("A2:A9", a2a9),
                EvalFactory.CreateAreaEval("A2:A9", a2a9),
                ErrorEval.VALUE_INVALID,
        };

            result = InvokeSumifs(args, EC);
            Assert.IsTrue(result is ErrorEval, "Expect to have an error when an input is an invalid value, but had: " + result.GetType());
        }
        [Test]
        public void TestBug56655b()
        {
            /*
                    setCellFormula(sheet, 0, 0, "B1*C1");
                    sheet.getRow(0).createCell(1).setCellValue("A");
                    setCellFormula(sheet, 1, 0, "B1*C1");
                    sheet.getRow(1).createCell(1).setCellValue("A");
                    setCellFormula(sheet, 0, 3, "SUMIFS(A:A,A:A,A2)");
             */
            ValueEval[] a0a1 = new ValueEval[] {
                NumberEval.ZERO,
                NumberEval.ZERO
        };

            ValueEval[] args = new ValueEval[]{
                EvalFactory.CreateAreaEval("A0:A1", a0a1),
                EvalFactory.CreateAreaEval("A0:A1", a0a1),
                ErrorEval.VALUE_INVALID
        };

            ValueEval result = InvokeSumifs(args, EC);
            Assert.IsTrue(result is ErrorEval, "Expect to have an error when an input is an invalid value, but had: " + result.GetType().Name);
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }
        [Test]
        public void TestBug56655c()
        {
            /*
                    setCellFormula(sheet, 0, 0, "B1*C1");
                    sheet.getRow(0).createCell(1).setCellValue("A");
                    setCellFormula(sheet, 1, 0, "B1*C1");
                    sheet.getRow(1).createCell(1).setCellValue("A");
                    setCellFormula(sheet, 0, 3, "SUMIFS(A:A,A:A,A2)");
             */
            ValueEval[] a0a1 = new ValueEval[] {
                NumberEval.ZERO,
                NumberEval.ZERO
        };

            ValueEval[] args = new ValueEval[]{
                EvalFactory.CreateAreaEval("A0:A1", a0a1),
                EvalFactory.CreateAreaEval("A0:A1", a0a1),
                ErrorEval.NAME_INVALID
        };

            ValueEval result = InvokeSumifs(args, EC);
            Assert.IsTrue(result is ErrorEval, "Expect to have an error when an input is an invalid value, but had: " + result.GetType());
            Assert.AreEqual(ErrorEval.NAME_INVALID, result);
        }
    }
}