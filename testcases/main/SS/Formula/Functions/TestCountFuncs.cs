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

using NPOI.SS.Util;

namespace TestCases.SS.Formula.Functions
{

    using NPOI.HSSF;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using System;
    using TestCases.HSSF;

    /**
     * Test cases for COUNT(), COUNTA() COUNTIF(), COUNTBLANK()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestCountFuncs
    {

        private static String NULL = null;

        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        [Test]
        public void TestCountBlank()
        {

            AreaEval range;
            ValueEval[] values;

            values = new ValueEval[] {
				new NumberEval(0),
				new StringEval(""),	// note - does not match blank
				BoolEval.TRUE,
				BoolEval.FALSE,
				ErrorEval.DIV_ZERO,
				BlankEval.instance,
		};
            range = EvalFactory.CreateAreaEval("A1:B3", values);
            ConfirmCountBlank(1, range);

            values = new ValueEval[] {
				new NumberEval(0),
				new StringEval(""),	// note - does not match blank
				BlankEval.instance,
				BoolEval.FALSE,
				BoolEval.TRUE,
				BlankEval.instance,
		};
            range = EvalFactory.CreateAreaEval("A1:B3", values);
            ConfirmCountBlank(2, range);
        }
        [Test]
        public void TestCountA()
        {

            ValueEval[] args;

            args = new ValueEval[] {
			new NumberEval(0),
		};
            ConfirmCountA(1, args);

            args = new ValueEval[] {
			new NumberEval(0),
			new NumberEval(0),
			new StringEval(""),
		};
            ConfirmCountA(3, args);

            args = new ValueEval[] {
			EvalFactory.CreateAreaEval("D2:F5", new ValueEval[12]),
		};
            ConfirmCountA(12, args);

            args = new ValueEval[] {
			EvalFactory.CreateAreaEval("D1:F5", new ValueEval[15]),
			EvalFactory.CreateRefEval("A1"),
			EvalFactory.CreateAreaEval("A1:G6", new ValueEval[42]),
			new NumberEval(0),
		};
            ConfirmCountA(59, args);
        }
        [Test]
        public void TestCountIf()
        {

            AreaEval range;
            ValueEval[] values;

            // when criteria is a bool value
            values = new ValueEval[] {
				new NumberEval(0),
				new StringEval("TRUE"),	// note - does not match bool TRUE
				BoolEval.TRUE,
				BoolEval.FALSE,
				BoolEval.TRUE,
				BlankEval.instance,
		};
            range = EvalFactory.CreateAreaEval("A1:B3", values);
            ConfirmCountIf(2, range, BoolEval.TRUE);

            // when criteria is numeric
            values = new ValueEval[] {
				new NumberEval(0),
				new StringEval("2"),
				new StringEval("2.001"),
				new NumberEval(2),
				new NumberEval(2),
				BoolEval.TRUE,
		};
            range = EvalFactory.CreateAreaEval("A1:B3", values);
            ConfirmCountIf(3, range, new NumberEval(2));
            // note - same results when criteria is a string that Parses as the number with the same value
            ConfirmCountIf(3, range, new StringEval("2.00"));

            // when criteria is an expression (starting with a comparison operator)
            ConfirmCountIf(2, range, new StringEval(">1"));
            // when criteria is an expression (starting with a comparison operator)
            ConfirmCountIf(2, range, new StringEval(">0.5"));
        }
        [Test]
        public void TestCriteriaPredicateNe_Bug46647()
        {
            IMatchPredicate mp = Countif.CreateCriteriaPredicate(new StringEval("<>aa"), 0, 0);
            StringEval seA = new StringEval("aa"); // this should not match the criteria '<>aa'
            StringEval seB = new StringEval("bb"); // this should match
            if (mp.Matches(seA) && !mp.Matches(seB))
            {
                throw new AssertionException("Identified bug 46647");
            }
            Assert.IsFalse(mp.Matches(seA));
            Assert.IsTrue(mp.Matches(seB));

            // general Tests for not-equal (<>) operator
            AreaEval range;
            ValueEval[] values;

            values = new ValueEval[] {
				new StringEval("aa"),
				new StringEval("def"),
				new StringEval("aa"),
				new StringEval("ghi"),
				new StringEval("aa"),
				new StringEval("aa"),
		};

            range = EvalFactory.CreateAreaEval("A1:A6", values);
            ConfirmCountIf(2, range, new StringEval("<>aa"));

            values = new ValueEval[] {
				new StringEval("ab"),
				new StringEval("aabb"),
				new StringEval("aa"), // match
				new StringEval("abb"),
				new StringEval("aab"),
				new StringEval("ba"), // match
		};

            range = EvalFactory.CreateAreaEval("A1:A6", values);
            ConfirmCountIf(2, range, new StringEval("<>a*b"));


            values = new ValueEval[] {
				new NumberEval(222),
				new NumberEval(222),
				new NumberEval(111),
				new StringEval("aa"),
				new StringEval("111"),
		};

            range = EvalFactory.CreateAreaEval("A1:A5", values);
            ConfirmCountIf(4, range, new StringEval("<>111"));
        }

        /**
         * String criteria in COUNTIF are case insensitive;
         * for example, the string "apples" and the string "APPLES" will match the same cells.
         */
        [Test]
        public void TestCaseInsensitiveStringComparison()
        {
            AreaEval range;
            ValueEval[] values;

            values = new ValueEval[] {
                new StringEval("no"),
                new StringEval("NO"),
                new StringEval("No"),
                new StringEval("Yes")
        };

            range = EvalFactory.CreateAreaEval("A1:A4", values);
            ConfirmCountIf(3, range, new StringEval("no"));
            ConfirmCountIf(3, range, new StringEval("NO"));
            ConfirmCountIf(3, range, new StringEval("No"));
        }

        /**
         * special case where the criteria argument is a cell reference
         */
        [Test]
        public void TestCountIfWithCriteriaReference()
        {

            ValueEval[] values = {
				new NumberEval(22),
				new NumberEval(25),
				new NumberEval(21),
				new NumberEval(25),
				new NumberEval(25),
				new NumberEval(25),
		};
            AreaEval arg0 = EvalFactory.CreateAreaEval("C1:C6", values);

            ValueEval criteriaArg = EvalFactory.CreateRefEval("A1", new NumberEval(25));
            ValueEval[] args = { arg0, criteriaArg, };

            double actual = NumericFunctionInvoker.Invoke(new Countif(), args);
            Assert.AreEqual(4, actual, 0D);
        }

        private static void ConfirmCountA(int expected, ValueEval[] args)
        {
            double result = NumericFunctionInvoker.Invoke(new Counta(), args);
            Assert.AreEqual(expected, result, 0);
        }
        private static void ConfirmCountIf(int expected, AreaEval range, ValueEval criteria)
        {

            ValueEval[] args = { range, criteria, };
            double result = NumericFunctionInvoker.Invoke(new Countif(), args);
            Assert.AreEqual(expected, result, 0);
        }
        private static void ConfirmCountBlank(int expected, AreaEval range)
        {

            ValueEval[] args = { range };
            double result = NumericFunctionInvoker.Invoke(new Countblank(), args);
            Assert.AreEqual(expected, result, 0);
        }

        private static IMatchPredicate CreateCriteriaPredicate(ValueEval ev)
        {
            return Countif.CreateCriteriaPredicate(ev, 0, 0);
        }

        /**
         * the criteria arg is mostly handled by {@link OperandResolver#getSingleValue(NPOI.SS.Formula.Eval.ValueEval, int, int)}}
         */
        [Test]
        public void TestCountifAreaCriteria()
        {
            int srcColIx = 2; // anything but column A

            ValueEval v0 = new NumberEval(2.0);
            ValueEval v1 = new StringEval("abc");
            ValueEval v2 = ErrorEval.DIV_ZERO;

            AreaEval ev = EvalFactory.CreateAreaEval("A10:A12", new ValueEval[] { v0, v1, v2, });

            IMatchPredicate mp;
            mp = Countif.CreateCriteriaPredicate(ev, 9, srcColIx);
            ConfirmPredicate(true, mp, srcColIx);
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(false, mp, ErrorEval.DIV_ZERO);

            mp = Countif.CreateCriteriaPredicate(ev, 10, srcColIx);
            ConfirmPredicate(false, mp, srcColIx);
            ConfirmPredicate(true, mp, "abc");
            ConfirmPredicate(false, mp, ErrorEval.DIV_ZERO);

            mp = Countif.CreateCriteriaPredicate(ev, 11, srcColIx);
            ConfirmPredicate(false, mp, srcColIx);
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(true, mp, ErrorEval.DIV_ZERO);
            ConfirmPredicate(false, mp, ErrorEval.VALUE_INVALID);

            // tricky: indexing outside of A10:A12
            // even this #VALUE! error Gets used by COUNTIF as valid criteria
            mp = Countif.CreateCriteriaPredicate(ev, 12, srcColIx);
            ConfirmPredicate(false, mp, srcColIx);
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(false, mp, ErrorEval.DIV_ZERO);
            ConfirmPredicate(true, mp, ErrorEval.VALUE_INVALID);
        }
        [Test]
        public void TestCountifEmptyStringCriteria()
        {
            IMatchPredicate mp;

            // pred '=' matches blank cell but not empty string
            mp = CreateCriteriaPredicate(new StringEval("="));
            ConfirmPredicate(false, mp, "");
            ConfirmPredicate(true, mp, NULL);

            // pred '' matches both blank cell but not empty string
            mp = CreateCriteriaPredicate(new StringEval(""));
            ConfirmPredicate(true, mp, "");
            ConfirmPredicate(true, mp, NULL);

            // pred '<>' matches empty string but not blank cell
            mp = CreateCriteriaPredicate(new StringEval("<>"));
            ConfirmPredicate(false, mp, NULL);
            ConfirmPredicate(true, mp, "");
        }
        [Test]
        public void TestCountifComparisons()
        {
            IMatchPredicate mp;

            mp = CreateCriteriaPredicate(new StringEval(">5"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(false, mp, 5);
            ConfirmPredicate(true, mp, 6);

            mp = CreateCriteriaPredicate(new StringEval("<=5"));
            ConfirmPredicate(true, mp, 4);
            ConfirmPredicate(true, mp, 5);
            ConfirmPredicate(false, mp, 6);
            ConfirmPredicate(false, mp, "4.9");
            ConfirmPredicate(false, mp, "4.9t");
            ConfirmPredicate(false, mp, "5.1");
            ConfirmPredicate(false, mp, NULL);

            mp = CreateCriteriaPredicate(new StringEval("=abc"));
            ConfirmPredicate(true, mp, "abc");

            mp = CreateCriteriaPredicate(new StringEval("=42"));
            ConfirmPredicate(false, mp, 41);
            ConfirmPredicate(true, mp, 42);
            ConfirmPredicate(true, mp, "42");

            mp = CreateCriteriaPredicate(new StringEval(">abc"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(true, mp, "abd");

            mp = CreateCriteriaPredicate(new StringEval(">4t3"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(false, mp, 500);
            ConfirmPredicate(true, mp, "500");
            ConfirmPredicate(true, mp, "4t4");
        }

        /**
         * the criteria arg value can be an error code (the error does not
         * propagate to the COUNTIF result).
         */
        [Test]
        public void TestCountifErrorCriteria()
        {
            IMatchPredicate mp;

            mp = CreateCriteriaPredicate(new StringEval("#REF!"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(false, mp, "#REF!");
            ConfirmPredicate(true, mp, ErrorEval.REF_INVALID);

            mp = CreateCriteriaPredicate(new StringEval("<#VALUE!"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(false, mp, "#DIV/0!");
            ConfirmPredicate(false, mp, "#REF!");
            ConfirmPredicate(true, mp, ErrorEval.DIV_ZERO);
            ConfirmPredicate(false, mp, ErrorEval.REF_INVALID);

            // not quite an error literal, should be treated as plain text
            mp = CreateCriteriaPredicate(new StringEval("<=#REF!a"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(true, mp, "#DIV/0!");
            ConfirmPredicate(true, mp, "#REF!");
            ConfirmPredicate(false, mp, ErrorEval.DIV_ZERO);
            ConfirmPredicate(false, mp, ErrorEval.REF_INVALID);
        }

        /**
        * Bug #51498 - Check that CountIf behaves correctly for GTE, LTE
        *  and NEQ cases
        */
        [Test]
        public void TestCountifBug51498()
        {
            int REF_COL = 4;
            int EVAL_COL = 3;

            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("51498.xls");
            IFormulaEvaluator evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            ISheet sheet = workbook.GetSheetAt(0);

            // numeric criteria
            for (int i = 0; i < 8; i++)
            {
                CellValue expected = evaluator.Evaluate(sheet.GetRow(i).GetCell(REF_COL));
                CellValue actual = evaluator.Evaluate(sheet.GetRow(i).GetCell(EVAL_COL));
                Assert.AreEqual(expected.FormatAsString(), actual.FormatAsString());
            }

            // boolean criteria
            for (int i = 0; i < 8; i++)
            {
                HSSFCell cellFmla = (HSSFCell)sheet.GetRow(i).GetCell(8);
                HSSFCell cellRef = (HSSFCell)sheet.GetRow(i).GetCell(9);

                double expectedValue = cellRef.NumericCellValue;
                double actualValue = evaluator.Evaluate(cellFmla).NumberValue;

                Assert.AreEqual(expectedValue, actualValue, 0.0001,
                    "Problem with a formula at " +
                                new CellReference(cellFmla).FormatAsString() + "[" + cellFmla.CellFormula + "] ");
            }

            // string criteria
            for (int i = 1; i < 9; i++)
            {
                ICell cellFmla = sheet.GetRow(i).GetCell(13);
                ICell cellRef = sheet.GetRow(i).GetCell(14);

                double expectedValue = cellRef.NumericCellValue;
                double actualValue = evaluator.Evaluate(cellFmla).NumberValue;

                Assert.AreEqual(expectedValue, actualValue, 0.0001,
                    "Problem with a formula at " +
                                new CellReference(cellFmla).FormatAsString() + "[" + cellFmla.CellFormula + "] ");
            }
        }

        [Test]
        public void TestWildCards()
        {
            IMatchPredicate mp;

            mp = CreateCriteriaPredicate(new StringEval("a*b"));
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(true, mp, "ab");
            ConfirmPredicate(true, mp, "axxb");
            ConfirmPredicate(false, mp, "xab");

            mp = CreateCriteriaPredicate(new StringEval("a?b"));
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(false, mp, "ab");
            ConfirmPredicate(false, mp, "axxb");
            ConfirmPredicate(false, mp, "xab");
            ConfirmPredicate(true, mp, "axb");

            mp = CreateCriteriaPredicate(new StringEval("a~?"));
            ConfirmPredicate(false, mp, "a~a");
            ConfirmPredicate(false, mp, "a~?");
            ConfirmPredicate(true, mp, "a?");

            mp = CreateCriteriaPredicate(new StringEval("~*a"));
            ConfirmPredicate(false, mp, "~aa");
            ConfirmPredicate(false, mp, "~*a");
            ConfirmPredicate(true, mp, "*a");

            mp = CreateCriteriaPredicate(new StringEval("12?12"));
            ConfirmPredicate(false, mp, 12812);
            ConfirmPredicate(true, mp, "12812");
            ConfirmPredicate(false, mp, "128812");
        }
        [Test]
        public void TestNotQuiteWildCards()
        {
            IMatchPredicate mp;

            // make sure special reg-ex chars are treated like normal chars
            mp = CreateCriteriaPredicate(new StringEval("a.b"));
            ConfirmPredicate(false, mp, "aab");
            ConfirmPredicate(true, mp, "a.b");


            mp = CreateCriteriaPredicate(new StringEval("a~b"));
            ConfirmPredicate(false, mp, "ab");
            ConfirmPredicate(false, mp, "axb");
            ConfirmPredicate(false, mp, "a~~b");
            ConfirmPredicate(true, mp, "a~b");

            mp = CreateCriteriaPredicate(new StringEval(">a*b"));
            ConfirmPredicate(false, mp, "a(b");
            ConfirmPredicate(true, mp, "aab");
            ConfirmPredicate(false, mp, "a*a");
            ConfirmPredicate(true, mp, "a*c");
        }

        private static void ConfirmPredicate(bool expectedResult, IMatchPredicate matchPredicate, int value)
        {
            Assert.AreEqual(expectedResult, matchPredicate.Matches(new NumberEval(value)));
        }
        private static void ConfirmPredicate(bool expectedResult, IMatchPredicate matchPredicate, String value)
        {
            ValueEval ev = (value == null) ? BlankEval.instance : (ValueEval)new StringEval(value);
            Assert.AreEqual(expectedResult, matchPredicate.Matches(ev));
        }
        private static void ConfirmPredicate(bool expectedResult, IMatchPredicate matchPredicate, ErrorEval value)
        {
            Assert.AreEqual(expectedResult, matchPredicate.Matches(value));
        }
        [Test]
        public void TestCountifFromSpreadsheet()
        {
            TestCountFunctionFromSpreadsheet("countifExamples.xls", 1, 2, 3, "countif");
        }

        /**
         * Two COUNTIF examples taken from
         * http://office.microsoft.com/en-us/excel-help/countif-function-HP010069840.aspx?CTT=5&origin=HA010277524
         */
        [Test]
        public void TestCountifExamples()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("countifExamples.xls");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            HSSFSheet sheet1 = (HSSFSheet)wb.GetSheet("MSDN Example 1");
            for (int rowIx = 7; rowIx <= 12; rowIx++)
            {
                HSSFRow row = (HSSFRow)sheet1.GetRow(rowIx - 1);
                HSSFCell cellA = (HSSFCell)row.GetCell(0);  // cell containing a formula with COUNTIF
                Assert.AreEqual(CellType.Formula, cellA.CellType);
                HSSFCell cellC = (HSSFCell)row.GetCell(2);  // cell with a reference value
                Assert.AreEqual(CellType.Numeric, cellC.CellType);

                CellValue cv = fe.Evaluate(cellA);
                double actualValue = cv.NumberValue;
                double expectedValue = cellC.NumericCellValue;
                Assert.AreEqual(expectedValue, actualValue, 0.0001,
                    "Problem with a formula at  " + new CellReference(cellA).FormatAsString()
                                + ": " + cellA.CellFormula + " :"
                        + "Expected = (" + expectedValue + ") Actual=(" + actualValue + ") ");
            }

            HSSFSheet sheet2 = (HSSFSheet)wb.GetSheet("MSDN Example 2");
            for (int rowIx = 9; rowIx <= 14; rowIx++)
            {
                HSSFRow row = (HSSFRow)sheet2.GetRow(rowIx - 1);
                HSSFCell cellA = (HSSFCell)row.GetCell(0);  // cell containing a formula with COUNTIF
                Assert.AreEqual(CellType.Formula, cellA.CellType);
                HSSFCell cellC = (HSSFCell)row.GetCell(2);  // cell with a reference value
                Assert.AreEqual(CellType.Numeric, cellC.CellType);

                CellValue cv = fe.Evaluate(cellA);
                double actualValue = cv.NumberValue;
                double expectedValue = cellC.NumericCellValue;

                Assert.AreEqual(expectedValue, actualValue, 0.0001,
                    "Problem with a formula at " +
                                new CellReference(cellA).FormatAsString() + "[" + cellA.CellFormula + "]: "
                                + "Expected = (" + expectedValue + ") Actual=(" + actualValue + ") ");

            }
        }

        [Test]
        public void TestCountBlankFromSpreadsheet()
        {
            TestCountFunctionFromSpreadsheet("countblankExamples.xls", 1, 3, 4, "countblank");
        }

        private static void TestCountFunctionFromSpreadsheet(String FILE_NAME, int START_ROW_IX, int COL_IX_ACTUAL, int COL_IX_EXPECTED, String functionName)
        {

            int failureCount = 0;
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook(FILE_NAME);
            ISheet sheet = wb.GetSheetAt(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            int maxRow = sheet.LastRowNum;
            for (int rowIx = START_ROW_IX; rowIx < maxRow; rowIx++)
            {
                IRow row = sheet.GetRow(rowIx);
                if (row == null)
                {
                    continue;
                }
                ICell cell = row.GetCell(COL_IX_ACTUAL);
                CellValue cv = fe.Evaluate(cell);
                double actualValue = cv.NumberValue;
                double expectedValue = row.GetCell(COL_IX_EXPECTED).NumericCellValue;
                if (actualValue != expectedValue)
                {
                    System.Console.Error.WriteLine("Problem with Test case on row " + (rowIx + 1) + " "
                            + "Expected = (" + expectedValue + ") Actual=(" + actualValue + ") ");
                    failureCount++;
                }
            }

            if (failureCount > 0)
            {
                throw new AssertionException(failureCount + " " + functionName
                        + " Evaluations failed. See stderr for more details");
            }
        }
    }

}