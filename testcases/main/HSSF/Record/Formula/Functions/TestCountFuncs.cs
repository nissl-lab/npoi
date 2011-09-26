/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


namespace TestCases.HSSF.Record.Formula.Functions
{
    using System;
    using NPOI.HSSF.Record.Formula.Functions;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Record.Formula.Eval;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Test cases for COUNT(), COUNTA() COUNTIF(), COUNTBLANK()
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestCountFuncs
    {

        public TestCountFuncs()
        {

        }
        [TestMethod]
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
			    EvalFactory.CreateAreaEval("D1:F5",new ValueEval[15]),	// 15
			    EvalFactory.CreateRefEval("A1"),	
			    EvalFactory.CreateAreaEval("A1:G6", new ValueEval[42]),	// 42
			    new NumberEval(0),
		    };
            ConfirmCountA(59, args);
        }
        [TestMethod]
        public void TestCountIf()
        {

            AreaEval range;
            ValueEval[] values;

            // when criteria is1 a bool value
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

            // when criteria is1 numeric
            values = new ValueEval[] {
				new NumberEval(0),	
				new StringEval("2"),	
				new StringEval("2.001"),	
				new NumberEval(2),	
				new NumberEval(2),	
				BoolEval.TRUE,
		    };
            range =EvalFactory.CreateAreaEval("A1:B3", values);
            ConfirmCountIf(3, range, new NumberEval(2));
            // note - same results when criteria is1 a string that parses as the number with the same value
            ConfirmCountIf(3, range, new StringEval("2.00"));

            //if (false)
            //{ // not supported yet: 
            //    // when criteria is1 an expression (starting with a comparison operator)
            //    ConfirmCountIf(4, range, new StringEval(">1"));
            //}
        }
        /**
         * special case where the criteria argument is1 a cell reference
         */
        [TestMethod]
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
        [TestMethod]
        public void TestCountIfEmptyStringCriteria()
        {
            I_MatchPredicate mp;

            // pred '=' Matches blank cell but not empty string
            mp = Countif.CreateCriteriaPredicate(new StringEval("="));
            ConfirmPredicate(false, mp, "");
            //ConfirmPredicate(true, mp, null);

            // pred '' Matches both blank cell but not empty string
            mp = Countif.CreateCriteriaPredicate(new StringEval(""));
            ConfirmPredicate(true, mp, "");
            //ConfirmPredicate(true, mp, null);

            // pred '<>' Matches empty string but not blank cell
            mp = Countif.CreateCriteriaPredicate(new StringEval("<>"));
            ConfirmPredicate(false, mp, null);
            ConfirmPredicate(true, mp, "");
        }
        [TestMethod]
        public void TestCountifComparisons()
        {
            I_MatchPredicate mp;

            mp = Countif.CreateCriteriaPredicate(new StringEval(">5"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(false, mp, 5);
            ConfirmPredicate(true, mp, 6);

            mp = Countif.CreateCriteriaPredicate(new StringEval("<=5"));
            ConfirmPredicate(true, mp, 4);
            ConfirmPredicate(true, mp, 5);
            ConfirmPredicate(false, mp, 6);
            ConfirmPredicate(true, mp, "4.9");
            ConfirmPredicate(false, mp, "4.9t");
            ConfirmPredicate(false, mp, "5.1");
            ConfirmPredicate(false, mp, null);

            mp = Countif.CreateCriteriaPredicate(new StringEval("=abc"));
            ConfirmPredicate(true, mp, "abc");

            mp = Countif.CreateCriteriaPredicate(new StringEval("=42"));
            ConfirmPredicate(false, mp, 41);
            ConfirmPredicate(true, mp, 42);
            ConfirmPredicate(true, mp, "42");

            mp = Countif.CreateCriteriaPredicate(new StringEval(">abc"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(true, mp, "abd");

            mp = Countif.CreateCriteriaPredicate(new StringEval(">4t3"));
            ConfirmPredicate(false, mp, 4);
            ConfirmPredicate(false, mp, 500);
            ConfirmPredicate(true, mp, "500");
            ConfirmPredicate(true, mp, "4t4");
        }
        [TestMethod]
        public void TestWildCards()
        {
            I_MatchPredicate mp;

            mp = Countif.CreateCriteriaPredicate(new StringEval("a*b"));
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(true, mp, "ab");
            ConfirmPredicate(true, mp, "axxb");
            ConfirmPredicate(false, mp, "xab");

            mp = Countif.CreateCriteriaPredicate(new StringEval("a?b"));
            ConfirmPredicate(false, mp, "abc");
            ConfirmPredicate(false, mp, "ab");
            ConfirmPredicate(false, mp, "axxb");
            ConfirmPredicate(false, mp, "xab");
            ConfirmPredicate(true, mp, "axb");

            mp = Countif.CreateCriteriaPredicate(new StringEval("a~?"));
            ConfirmPredicate(false, mp, "a~a");
            ConfirmPredicate(false, mp, "a~?");
            ConfirmPredicate(true, mp, "a?");

            mp = Countif.CreateCriteriaPredicate(new StringEval("~*a"));
            ConfirmPredicate(false, mp, "~aa");
            ConfirmPredicate(false, mp, "~*a");
            ConfirmPredicate(true, mp, "*a");

            mp = Countif.CreateCriteriaPredicate(new StringEval("12?12"));
            ConfirmPredicate(false, mp, 12812);
            ConfirmPredicate(true, mp, "12812");
            ConfirmPredicate(false, mp, "128812");
        }
        [TestMethod]
        public void TestNotQuiteWildCards()
        {
            I_MatchPredicate mp;

            // make sure special reg-ex chars are treated like normal chars
            mp = Countif.CreateCriteriaPredicate(new StringEval("a.b"));
            ConfirmPredicate(false, mp, "aab");
            ConfirmPredicate(true, mp, "a.b");


            mp = Countif.CreateCriteriaPredicate(new StringEval("a~b"));
            ConfirmPredicate(false, mp, "ab");
            ConfirmPredicate(false, mp, "axb");
            ConfirmPredicate(false, mp, "a~~b");
            ConfirmPredicate(true, mp, "a~b");

            mp = Countif.CreateCriteriaPredicate(new StringEval(">a*b"));
            ConfirmPredicate(false, mp, "a(b");
            ConfirmPredicate(true, mp, "aab");
            ConfirmPredicate(false, mp, "a*a");
            ConfirmPredicate(true, mp, "a*c");
        }

        private static void ConfirmPredicate(bool expectedResult, I_MatchPredicate matchPredicate, int value)
        {
            Assert.AreEqual(expectedResult, matchPredicate.Matches(new NumberEval(value)));
        }
        private static void ConfirmPredicate(bool expectedResult, I_MatchPredicate matchPredicate, String value)
        {
            ValueEval ev = value == null ? (ValueEval)BlankEval.instance : new StringEval(value);
            Assert.AreEqual(expectedResult, matchPredicate.Matches(ev));
        }
        [TestMethod]
        public void TestCountifFromSpreadsheet() {
		    String FILE_NAME = "countifExamples.xls";
		    int START_ROW_IX = 1;
		    int COL_IX_ACTUAL = 2;
		    int COL_IX_EXPECTED = 3;

		    int failureCount = 0;
		    HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook(FILE_NAME);
		    NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);
		    HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
		    int maxRow = sheet.LastRowNum;
		    for (int rowIx=START_ROW_IX; rowIx<maxRow; rowIx++) {
			    NPOI.SS.UserModel.IRow row = sheet.GetRow(rowIx);
			    if(row == null) {
				    continue;
			    }
			    ICell cell = row.GetCell(COL_IX_ACTUAL);
                NPOI.SS.UserModel.CellValue cv = fe.Evaluate(cell);
			    double actualValue = cv.NumberValue;
			    double expectedValue = row.GetCell(COL_IX_EXPECTED).NumericCellValue;
			    if (actualValue != expectedValue) {
				    Console.Error.WriteLine("Problem with Test case on row " + (rowIx+1) + " "
						    + "Expected = (" + expectedValue + ") Actual=(" + actualValue + ") ");
				    failureCount++;
			    }
		    }

		    if (failureCount > 0) {
			    throw new AssertFailedException(failureCount + " countif evaluations failed. See stderr for more details");
		    }
	    }
    }
}