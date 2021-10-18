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
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.UserModel;
    using NUnit.Framework;

    /**
     * Test cases for SUMPRODUCT()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestSumif
    {
        private static NumberEval _30 = new NumberEval(30);
        private static NumberEval _40 = new NumberEval(40);
        private static NumberEval _50 = new NumberEval(50);
        private static NumberEval _60 = new NumberEval(60);

        private static ValueEval invokeSumif(int rowIx, int colIx, params ValueEval[] args)
        {
            return new Sumif().Evaluate(args, rowIx, colIx);
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

        [Test]
        public void TestBasic()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            ValueEval[] arg0values = new ValueEval[] { _30, _30, _40, _40, _50, _50 };
            ValueEval[] arg2values = new ValueEval[] { _30, _40, _50, _60, _60, _60 };

            AreaEval arg0;
            AreaEval arg2;

            arg0 = EvalFactory.CreateAreaEval("A3:B5", arg0values);
            arg2 = EvalFactory.CreateAreaEval("D1:E3", arg2values);

            Confirm(60.0, arg0, new NumberEval(30.0));
            Confirm(70.0, arg0, new NumberEval(30.0), arg2);
            Confirm(100.0, arg0, new StringEval(">45"));
            Confirm(100.0, arg0, new StringEval(">=45"));
            Confirm(100.0, arg0, new StringEval(">=50.0"));
            Confirm(140.0, arg0, new StringEval("<45"));
            Confirm(140.0, arg0, new StringEval("<=45"));
            Confirm(140.0, arg0, new StringEval("<=40.0"));
            Confirm(160.0, arg0, new StringEval("<>40.0"));
            Confirm(80.0, arg0, new StringEval("=40.0"));
        }

        private static void Confirm(double expectedResult, params ValueEval[] args)
        {
            ConfirmDouble(expectedResult, invokeSumif(-1, -1, args));
        }


        /**
         * Test for bug observed near svn r882931
         */
        [Test]
        public void TestCriteriaArgRange()
        {
            ValueEval[] arg0values = new ValueEval[] { _50, _60, _50, _50, _50, _30, };
            ValueEval[] arg1values = new ValueEval[] { _30, _40, _50, _60, };

            AreaEval arg0;
            AreaEval arg1;
            ValueEval ve;

            arg0 = EvalFactory.CreateAreaEval("A3:B5", arg0values);
            arg1 = EvalFactory.CreateAreaEval("A2:D2", arg1values); // single row range

            ve = invokeSumif(0, 2, arg0, arg1);  // invoking from cell C1
            if (ve is NumberEval)
            {
                NumberEval ne = (NumberEval)ve;
                if (ne.NumberValue == 30.0)
                {
                    throw new AssertionException("identified error in SUMIF - criteria arg not Evaluated properly");
                }
            }

            ConfirmDouble(200, ve);

            arg0 = EvalFactory.CreateAreaEval("C1:D3", arg0values);
            arg1 = EvalFactory.CreateAreaEval("B1:B4", arg1values); // single column range

            ve = invokeSumif(3, 0, arg0, arg1); // invoking from cell A4

            ConfirmDouble(60, ve);
        }

        [Test]
        public void TestEvaluateException()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, invokeSumif(-1, -1, BlankEval.instance, new NumberEval(30.0)));
            Assert.AreEqual(ErrorEval.VALUE_INVALID, invokeSumif(-1, -1, BlankEval.instance, new NumberEval(30.0), new NumberEval(30.0)));
            Assert.AreEqual(ErrorEval.VALUE_INVALID, invokeSumif(-1, -1, new NumberEval(30.0), BlankEval.instance, new NumberEval(30.0)));
            Assert.AreEqual(ErrorEval.VALUE_INVALID, invokeSumif(-1, -1, new NumberEval(30.0), new NumberEval(30.0), BlankEval.instance));
        }

        [Test]
        public void TestMicrosoftExample1()
        {
            IWorkbook wb = initWorkbook1();

            IFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(A2:A5,\">160000\",B2:B5)", 63000);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(A2:A5,\">160000\")", 900000);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(A2:A5,300000,B2:B5)", 21000);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(A2:A5,\">\" & C2,B2:B5)", 49000);
        }

        [Test]
        public void TestMicrosoftExample1WithNA()
        {
            IWorkbook wb = initWorkbook1WithNA();
            IFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            Util.Utils.AssertError(fe, cell, "SUMIF(A2:A6,\">160000\",B2:B6)", FormulaError.NA);
        }

        [Test]
        public void TestMicrosoftExample1WithBooleanAndString()
        {
            IWorkbook wb = initWorkbook1WithBooleanAndString();

            IFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(A2:A7,\">160000\",B2:B7)", 63000);
        }

        [Test]
        public void TestMicrosoftExample2()
        {
            IWorkbook wb = initWorkbook2();
            IFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(A2:A7,\"Fruits\",C2:C7)", 2000);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(A2:A7,\"Vegetables\",C2:C7)", 12000);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(B2:B7,\"*es\",C2:C7)", 4300);
            Util.Utils.AssertDouble(fe, cell, "SUMIF(A2:A7,\"\",C2:C7)", 400);
        }

        //see https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b
        private IWorkbook initWorkbook1()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            Util.Utils.AddRow(sheet, 0, "Property Value", "Commission", "Data");
            Util.Utils.AddRow(sheet, 1, 100000, 7000, 250000);
            Util.Utils.AddRow(sheet, 2, 200000, 14000);
            Util.Utils.AddRow(sheet, 3, 300000, 21000);
            Util.Utils.AddRow(sheet, 4, 400000, 28000);
            return wb;
        }

        private IWorkbook initWorkbook1WithNA()
        {
            IWorkbook wb = initWorkbook1();
            ISheet sheet = wb.GetSheetAt(0);
            Util.Utils.AddRow(sheet, 5, 500000, FormulaError.NA);
            return wb;
        }

        private IWorkbook initWorkbook1WithBooleanAndString()
        {
            IWorkbook wb = initWorkbook1();
            ISheet sheet = wb.GetSheetAt(0);
            Util.Utils.AddRow(sheet, 5, 500000, true);
            Util.Utils.AddRow(sheet, 6, 600000, "abc");
            return wb;
        }

        private IWorkbook initWorkbook2()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            Util.Utils.AddRow(sheet, 0, "Category", "Food", "Sales");
            Util.Utils.AddRow(sheet, 1, "Vegetables", "Tomatoes", 2300);
            Util.Utils.AddRow(sheet, 2, "Vegetables", "Celery", 5500);
            Util.Utils.AddRow(sheet, 3, "Fruits", "Oranges", 800);
            Util.Utils.AddRow(sheet, 4, null, "Butter", 400);
            Util.Utils.AddRow(sheet, 5, "Vegetables", "Carrots", 4200);
            Util.Utils.AddRow(sheet, 6, "Fruits", "Apples", 1200);
            return wb;
        }


    }
}