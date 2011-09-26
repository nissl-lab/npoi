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

namespace TestCases.HSSF.Model
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.UserModel;
    using TestCases.SS.Formula;
    using NPOI.SS.UserModel;

    /**
     * Test the low level formula Parser functionality. High level tests are to
     * be done via usermodel/Cell.SetFormulaValue().
     */
    [TestClass]
    public class TestFormulaParser
    {

        /**
         * @return Parsed token array alReady Confirmed not <c>null</c>
         */
        /* package */
        public static Ptg[] ParseFormula(String formula)
        {
            Ptg[] result = HSSFFormulaParser.Parse(formula, null);
            Assert.IsNotNull(result, "Ptg array should not be null");
            return result;
        }
        [TestMethod]
        public void TestSimpleFormula()
        {
            Ptg[] ptgs = ParseFormula("2+2");
            Assert.AreEqual(3, ptgs.Length);
        }
        [TestMethod]
        public void TestFormulaWithSpace1()
        {
            Ptg[] ptgs = ParseFormula(" 2 + 2 ");
            Assert.AreEqual(3, ptgs.Length);
            Assert.IsTrue((ptgs[0] is IntPtg));
            Assert.IsTrue((ptgs[1] is IntPtg));
            Assert.IsTrue((ptgs[2] is AddPtg));
        }
        [TestMethod]
        public void TestFormulaWithSpace2()
        {
            Ptg[] ptgs = ParseFormula("2+ sum( 3 , 4) ");
            Assert.AreEqual(5, ptgs.Length);
        }
        [TestMethod]
        public void TestFormulaWithSpaceNRef()
        {
            Ptg[] ptgs = ParseFormula("sum( A2:A3 )");
            Assert.AreEqual(2, ptgs.Length);
        }
        [TestMethod]
        public void TestFormulaWithString()
        {
            Ptg[] ptgs = ParseFormula("\"hello\" & \"world\" ");
            Assert.AreEqual(3, ptgs.Length);
        }
        [TestMethod]
        public void TestTRUE()
        {
            Ptg[] ptgs = ParseFormula("TRUE");
            Assert.AreEqual(1, ptgs.Length);
            BoolPtg flag = (BoolPtg)ptgs[0];
            Assert.AreEqual(true, flag.Value);
        }
        [TestMethod]
        public void TestSumIf()
        {
            Ptg[] ptgs = ParseFormula("SUMIF(A1:A5,\">4000\",B1:B5)");
            Assert.AreEqual(4, ptgs.Length);
        }

        /**
         * Bug Reported by xt-jens.riis@nokia.com (Jens Riis)
         * Refers to Bug <a href="http://issues.apache.org/bugzilla/show_bug.cgi?id=17582">#17582</a>
         *
         */
        [TestMethod]
        public void TestNonAlphaFormula()
        {
            String currencyCell = "F3";
            Ptg[] ptgs = ParseFormula("\"TOTAL[\"&" + currencyCell + "&\"]\"");
            Assert.AreEqual(5, ptgs.Length);
            Assert.IsTrue((ptgs[0] is StringPtg), "Ptg[0] is1 a string");
            StringPtg firstString = (StringPtg)ptgs[0];

            Assert.AreEqual("TOTAL[", firstString.Value);
            //the PTG order isn't 100% correct but it still works - dmui
        }
        [TestMethod]
        public void TestMacroFunction()
        {
            // testNames.xls contains a VB function called 'myFunc'
            HSSFWorkbook w = HSSFTestDataSamples.OpenSampleWorkbook("testNames.xls");
            HSSFEvaluationWorkbook book = HSSFEvaluationWorkbook.Create(w);

            Ptg[] ptg = HSSFFormulaParser.Parse("myFunc()", w);

            // the name Gets encoded as the first arg
            NamePtg tname = (NamePtg)ptg[0];
            Assert.AreEqual("myFunc", tname.ToFormulaString(book));

            AbstractFunctionPtg tfunc = (AbstractFunctionPtg)ptg[1];
            Assert.IsTrue(tfunc.IsExternalFunction);
        }
        [TestMethod]
        public void TestEmbeddedSlash()
        {
            Ptg[] ptgs = ParseFormula("HYPERLINK(\"http://www.jakarta.org\",\"Jakarta\")");
            Assert.IsTrue(ptgs[0] is StringPtg, "first ptg is1 string");
            Assert.IsTrue(ptgs[1] is StringPtg, "second ptg is1 string");
        }
        [TestMethod]
        public void TestConcatenate()
        {
            Ptg[] ptgs = ParseFormula("CONCATENATE(\"first\",\"second\")");
            Assert.IsTrue(ptgs[0] is StringPtg, "first ptg is1 string");
            Assert.IsTrue(ptgs[1] is StringPtg, "second ptg is1 string");
        }
        [TestMethod]
        public void TestWorksheetReferences()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            wb.CreateSheet("NoQuotesNeeded");
            wb.CreateSheet("Quotes Needed Here &#$@");

            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell;

            cell = row.CreateCell((short)0);
            cell.CellFormula = ("NoQuotesNeeded!A1");

            cell = row.CreateCell((short)1);
            cell.CellFormula = ("'Quotes Needed Here &#$@'!A1");
        }
        [TestMethod]
        public void TestUnaryMinus()
        {
            Ptg[] ptgs = ParseFormula("-A1");
            Assert.AreEqual(2, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "first ptg is1 reference");
            Assert.IsTrue(ptgs[1] is UnaryMinusPtg, "second ptg is1 Minus");
        }
        [TestMethod]
        public void TestUnaryPlus()
        {
            Ptg[] ptgs = ParseFormula("+A1");
            Assert.AreEqual(2, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "first ptg is1 reference");
            Assert.IsTrue(ptgs[1] is UnaryPlusPtg, "second ptg is1 Plus");
        }
        [TestMethod]
        public void TestLeadingSpaceInString()
        {
            String value = "  hi  ";
            Ptg[] ptgs = ParseFormula("\"" + value + "\"");

            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is StringPtg, "ptg0 is1 a StringPtg");
            Assert.IsTrue(((StringPtg)ptgs[0]).Value.Equals(value), "ptg0 contains exact value");
        }
        [TestMethod]
        public void TestLookupAndMatchFunctionArgs()
        {
            Ptg[] ptgs = ParseFormula("lookup(A1, A3:A52, B3:B52)");

            Assert.AreEqual(4, ptgs.Length);
            Assert.IsTrue(ptgs[0].PtgClass == Ptg.CLASS_VALUE, "ptg0 has Value class");

            ptgs = ParseFormula("match(A1, A3:A52)");

            Assert.AreEqual(3, ptgs.Length);
            Assert.IsTrue(ptgs[0].PtgClass == Ptg.CLASS_VALUE, "ptg0 has Value class");
        }

        /** bug 33160*/
        [TestMethod]
        public void TestLargeInt()
        {
            Ptg[] ptgs = ParseFormula("40");
            Assert.IsTrue(ptgs[0] is IntPtg, "ptg is1 Int, is1 " + ptgs[0].GetType());

            ptgs = ParseFormula("40000");
            Assert.IsTrue(ptgs[0] is IntPtg, "ptg should be  IntPtg, is1 " + ptgs[0].GetType());
        }

        /** bug 33160, testcase by Amol Deshmukh*/
        [TestMethod]
        public void TestSimpleLongFormula()
        {
            Ptg[] ptgs = ParseFormula("40000/2");
            Assert.AreEqual(3, ptgs.Length);
            Assert.IsTrue((ptgs[0] is IntPtg), "IntPtg");
            Assert.IsTrue((ptgs[1] is IntPtg), "IntPtg");
            Assert.IsTrue((ptgs[2] is DividePtg), "DividePtg");
        }

        /** bug 35027, underscore in sheet name */
        [TestMethod]
        public void TestUnderscore()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            wb.CreateSheet("Cash_Flow");

            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell;

            cell = row.CreateCell((short)0);
            cell.CellFormula = ("Cash_Flow!A1");
        }

        // bug 38396 : Formula with exponential numbers not Parsed correctly.
        [TestMethod]
        public void TestExponentialParsing()
        {
            Ptg[] ptgs;
            ptgs = ParseFormula("1.3E21/2");
            Assert.AreEqual(3, ptgs.Length);
            Assert.IsTrue((ptgs[0] is NumberPtg), "NumberPtg");
            Assert.IsTrue((ptgs[1] is IntPtg), "IntPtg");
            Assert.IsTrue((ptgs[2] is DividePtg), "DividePtg");

            ptgs = ParseFormula("1322E21/2");
            Assert.AreEqual(3, ptgs.Length);
            Assert.IsTrue((ptgs[0] is NumberPtg), "NumberPtg");
            Assert.IsTrue((ptgs[1] is IntPtg), "IntPtg");
            Assert.IsTrue((ptgs[2] is DividePtg), "DividePtg");

            ptgs = ParseFormula("1.3E1/2");
            Assert.AreEqual(3, ptgs.Length);
            Assert.IsTrue((ptgs[0] is NumberPtg), "NumberPtg");
            Assert.IsTrue((ptgs[1] is IntPtg), "IntPtg");
            Assert.IsTrue((ptgs[2] is DividePtg), "DividePtg");
        }
        [TestMethod]
        public void TestExponentialInSheet()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            wb.CreateSheet("Cash_Flow");

            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            String formula = null;

            cell.CellFormula = ("1.3E21/3");
            formula = cell.CellFormula;
            Assert.AreEqual("1.3E+21/3", formula, "Exponential formula string");

            cell.CellFormula = ("-1.3E21/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-1.3E+21/3", formula, "Exponential formula string");

            cell.CellFormula = ("1322E21/3");
            formula = cell.CellFormula;
            Assert.AreEqual("1.322E+24/3", formula, "Exponential formula string");

            cell.CellFormula = ("-1322E21/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-1.322E+24/3", formula, "Exponential formula string");

            cell.CellFormula = ("1.3E1/3");
            formula = cell.CellFormula;
            Assert.AreEqual("13/3", formula, "Exponential formula string");

            cell.CellFormula = ("-1.3E1/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-13/3", formula, "Exponential formula string");

            cell.CellFormula = ("1.3E-4/3");
            formula = cell.CellFormula;
            Assert.AreEqual("0.00013/3", formula, "Exponential formula string");

            cell.CellFormula = ("-1.3E-4/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-0.00013/3", formula, "Exponential formula string");

            cell.CellFormula = ("13E-15/3");
            formula = cell.CellFormula;
            Assert.AreEqual("1.3E-14/3", formula, "Exponential formula string");

            cell.CellFormula = ("-13E-15/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-1.3E-14/3", formula, "Exponential formula string");

            cell.CellFormula = ("1.3E3/3");
            formula = cell.CellFormula;
            Assert.AreEqual("1300/3", formula, "Exponential formula string");

            cell.CellFormula = ("-1.3E3/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-1300/3", formula, "Exponential formula string");

            cell.CellFormula = ("1300000000000000/3");
            formula = cell.CellFormula;
            Assert.AreEqual("1.3E+15/3", formula, "Exponential formula string");

            cell.CellFormula = ("-1300000000000000/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-1.3E+15/3", formula, "Exponential formula string");

            cell.CellFormula = ("-10E-1/3.1E2*4E3/3E4");
            formula = cell.CellFormula;
            Assert.AreEqual("-1/310*4000/30000", formula, "Exponential formula string");
        }
        [TestMethod]
        public void TestNumbers()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            wb.CreateSheet("Cash_Flow");

            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            String formula = null;

            // starts from decimal point

            cell.CellFormula = (".1");
            formula = cell.CellFormula;
            Assert.AreEqual("0.1", formula);

            cell.CellFormula = ("+.1");
            formula = cell.CellFormula;
            Assert.AreEqual("+0.1", formula);

            cell.CellFormula = ("-.1");
            formula = cell.CellFormula;
            Assert.AreEqual("-0.1", formula);

            // has exponent

            cell.CellFormula = ("10E1");
            formula = cell.CellFormula;
            Assert.AreEqual("100", formula);

            cell.CellFormula = ("10E+1");
            formula = cell.CellFormula;
            Assert.AreEqual("100", formula);

            cell.CellFormula = ("10E-1");
            formula = cell.CellFormula;
            Assert.AreEqual("1", formula);
        }
        [TestMethod]
        public void TestRanges()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            wb.CreateSheet("Cash_Flow");

            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            String formula = null;

            cell.CellFormula = ("A1.A2");
            formula = cell.CellFormula;
            Assert.AreEqual("A1:A2", formula);

            cell.CellFormula = ("A1..A2");
            formula = cell.CellFormula;
            Assert.AreEqual("A1:A2", formula);

            cell.CellFormula = ("A1...A2");
            formula = cell.CellFormula;
            Assert.AreEqual("A1:A2", formula);
        }

        /**
         * Test for bug observable at svn revision 618865 (5-Feb-2008)<br/>
         * a formula consisting of a single no-arg function got rendered without the function braces
         */
        [TestMethod]
        public void TestToFormulaStringZeroArgFunction()
        {
            HSSFWorkbook book = new HSSFWorkbook();

            Ptg[] ptgs = {
				new FuncPtg(10),
		};
            Assert.AreEqual("NA()", HSSFFormulaParser.ToFormulaString(book, ptgs));
        }
        [TestMethod]
        public void TestPercent()
        {
            Ptg[] ptgs;
            ptgs = ParseFormula("5%");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(ptgs[0].GetType(), typeof(IntPtg));
            Assert.AreEqual(ptgs[1].GetType(), typeof(PercentPtg));

            // spaces OK
            ptgs = ParseFormula(" 250 % ");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(ptgs[0].GetType(), typeof(IntPtg));
            Assert.AreEqual(ptgs[1].GetType(), typeof(PercentPtg));


            // double percent OK
            ptgs = ParseFormula("12345.678%%");
            Assert.AreEqual(3, ptgs.Length);
            Assert.AreEqual(ptgs[0].GetType(), typeof(NumberPtg));
            Assert.AreEqual(ptgs[1].GetType(), typeof(PercentPtg));
            Assert.AreEqual(ptgs[2].GetType(), typeof(PercentPtg));

            // percent of a bracketed expression
            ptgs = ParseFormula("(A1+35)%*B1%");
            Assert.AreEqual(8, ptgs.Length);
            Assert.AreEqual(ptgs[4].GetType(), typeof(PercentPtg));
            Assert.AreEqual(ptgs[6].GetType(), typeof(PercentPtg));

            // percent of a text quantity
            ptgs = ParseFormula("\"8.75\"%");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(ptgs[0].GetType(), typeof(StringPtg));
            Assert.AreEqual(ptgs[1].GetType(), typeof(PercentPtg));

            // percent to the power of
            ptgs = ParseFormula("50%^3");
            Assert.AreEqual(4, ptgs.Length);
            Assert.AreEqual(ptgs[0].GetType(), typeof(IntPtg));
            Assert.AreEqual(ptgs[1].GetType(), typeof(PercentPtg));
            Assert.AreEqual(ptgs[2].GetType(), typeof(IntPtg));
            Assert.AreEqual(ptgs[3].GetType(), typeof(PowerPtg));

            //
            // things that Parse OK but would *evaluate* to an error

            ptgs = ParseFormula("\"abc\"%");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(ptgs[0].GetType(), typeof(StringPtg));
            Assert.AreEqual(ptgs[1].GetType(), typeof(PercentPtg));

            ptgs = ParseFormula("#N/A%");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(ptgs[0].GetType(), typeof(ErrPtg));
            Assert.AreEqual(ptgs[1].GetType(), typeof(PercentPtg));
        }

        /**
         * Tests combinations of various operators in the absence of brackets
         */
        [TestMethod]
        public void TestPrecedenceAndAssociativity()
        {

            Type[] expClss;

            // TRUE=TRUE=2=2  evaluates to FALSE
            expClss = new Type[] { typeof(BoolPtg), typeof(BoolPtg), typeof(EqualPtg),
				typeof(IntPtg), typeof(EqualPtg), typeof(IntPtg), typeof(EqualPtg),  };
            ConfirmTokenClasses("TRUE=TRUE=2=2", expClss);


            //  2^3^2	evaluates to 64 not 512
            expClss = new Type[] { typeof(IntPtg), typeof(IntPtg), typeof(PowerPtg),
				typeof(IntPtg), typeof(PowerPtg), };
            ConfirmTokenClasses("2^3^2", expClss);

            // "abc" & 2 + 3 & "def"   evaluates to "abc5def"
            expClss = new Type[] { typeof(StringPtg), typeof(IntPtg), typeof(IntPtg),
				typeof(AddPtg), typeof(ConcatPtg), typeof(StringPtg), typeof(ConcatPtg), };
            ConfirmTokenClasses("\"abc\"&2+3&\"def\"", expClss);


            //  (1 / 2) - (3 * 4)
            expClss = new Type[] { typeof(IntPtg), typeof(IntPtg), typeof(DividePtg),
				typeof(IntPtg), typeof(IntPtg), typeof(MultiplyPtg), typeof(SubtractPtg), };
            ConfirmTokenClasses("1/2-3*4", expClss);

            // 2 * (2^2)
            expClss = new Type[] { typeof(IntPtg), typeof(IntPtg), typeof(IntPtg), typeof(PowerPtg), typeof(MultiplyPtg), };
            // NOT: (2 *2) ^ 2 -> int int multiply int power
            ConfirmTokenClasses("2*2^2", expClss);

            //  2^200% -> 2 not 1.6E58
            expClss = new Type[] { typeof(IntPtg), typeof(IntPtg), typeof(PercentPtg), typeof(PowerPtg), };
            ConfirmTokenClasses("2^200%", expClss);
        }

        /* package */
        public static Ptg[] ConfirmTokenClasses(String formula, Type[] expectedClasses)
        {
            Ptg[] ptgs = ParseFormula(formula);
            Assert.AreEqual(expectedClasses.Length, ptgs.Length);
            for (int i = 0; i < expectedClasses.Length; i++)
            {
                if (expectedClasses[i] != ptgs[i].GetType())
                {
                    Assert.Fail("difference at token[" + i + "]: expected ("
                        + expectedClasses[i].Name + ") but got ("
                        + ptgs[i].GetType().Name + ")");
                }
            }
            return ptgs;
        }
        [TestMethod]
        public void TestPower()
        {
            ConfirmTokenClasses("2^5", new Type[] { typeof(IntPtg), typeof(IntPtg), typeof(PowerPtg), });
        }

        private static Ptg ParseSingleToken(String formula, Type ptgClass)
        {
            Ptg[] ptgs = ParseFormula(formula);
            Assert.AreEqual(1, ptgs.Length);
            Ptg result = ptgs[0];
            Assert.AreEqual(ptgClass, result.GetType());
            return result;
        }
        [TestMethod]
        public void TestParseNumber()
        {
            IntPtg ip;

            // bug 33160
            ip = (IntPtg)ParseSingleToken("40", typeof(IntPtg));
            Assert.AreEqual(40, ip.Value);
            ip = (IntPtg)ParseSingleToken("40000", typeof(IntPtg));
            Assert.AreEqual(40000, ip.Value);

            // check the upper edge of the IntPtg range:
            ip = (IntPtg)ParseSingleToken("65535", typeof(IntPtg));
            Assert.AreEqual(65535, ip.Value);
            NumberPtg np = (NumberPtg)ParseSingleToken("65536", typeof(NumberPtg));
            Assert.AreEqual(65536, np.Value, 0);

            np = (NumberPtg)ParseSingleToken("65534.6", typeof(NumberPtg));
            Assert.AreEqual(65534.6, np.Value, 0);
        }
        [TestMethod]
        public void TestMissingArgs()
        {

            Type[] expClss;

            expClss = new Type[] { 
				typeof(RefPtg), 
				typeof(AttrPtg), // tAttrIf
				typeof(MissingArgPtg), 
				typeof(AttrPtg), // tAttrSkip
				typeof(RefPtg),
				typeof(AttrPtg), // tAttrSkip
				typeof(FuncVarPtg), 
		};

            ConfirmTokenClasses("if(A1, ,C1)", expClss);

            expClss = new Type[] { typeof(MissingArgPtg), typeof(AreaPtg), typeof(MissingArgPtg),
				typeof(FuncVarPtg), };
            ConfirmTokenClasses("counta( , A1:B2, )", expClss);
        }
        [TestMethod]
        public void TestParseErrorLiterals()
        {

            ConfirmParseErrorLiteral(ErrPtg.NULL_INTERSECTION, "#NULL!");
            ConfirmParseErrorLiteral(ErrPtg.DIV_ZERO, "#DIV/0!");
            ConfirmParseErrorLiteral(ErrPtg.VALUE_INVALID, "#VALUE!");
            ConfirmParseErrorLiteral(ErrPtg.REF_INVALID, "#REF!");
            ConfirmParseErrorLiteral(ErrPtg.NAME_INVALID, "#NAME?");
            ConfirmParseErrorLiteral(ErrPtg.NUM_ERROR, "#NUM!");
            ConfirmParseErrorLiteral(ErrPtg.N_A, "#N/A");
        }

        private static void ConfirmParseErrorLiteral(ErrPtg expectedToken, String formula)
        {
            Assert.AreEqual(expectedToken, ParseSingleToken(formula, typeof(ErrPtg)));
        }

        /**
         * To aid Readability the parameters have been encoded with single quotes instead of double
         * quotes.  This method converts single quotes to double quotes before performing the Parse
         * and result check.
         */
        private static void ConfirmStringParse(String singleQuotedValue)
        {
            // formula: internal quotes become double double, surround with double quotes
            String formula = '"' + singleQuotedValue.Replace("'", "\"\"") + '"';
            String expectedValue = singleQuotedValue.Replace('\'', '"');

            StringPtg sp = (StringPtg)ParseSingleToken(formula, typeof(StringPtg));
            Assert.AreEqual(expectedValue, sp.Value);
        }
        [TestMethod]
        public void TestParseStringLiterals_bug28754()
        {

            StringPtg sp;
            try
            {
                sp = (StringPtg)ParseSingleToken("\"test\"\"ing\"", typeof(StringPtg));
            }
            catch (Exception e)
            {
                if (e.Message.StartsWith("Cannot Parse"))
                {
                    throw new AssertFailedException("Identified bug 28754a");
                }
                throw e;
            }
            Assert.AreEqual("test\"ing", sp.Value);

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            cell.CellFormula = ("right(\"test\"\"ing\", 3)");
            String actualCellFormula = cell.CellFormula;
            if ("RIGHT(\"test\"ing\",3)".Equals(actualCellFormula))
            {
                throw new AssertFailedException("Identified bug 28754b");
            }
            Assert.AreEqual("RIGHT(\"test\"\"ing\",3)", actualCellFormula);
        }
        [TestMethod]
        public void TestParseStringLiterals()
        {
            ConfirmStringParse("goto considered harmful");

            ConfirmStringParse("goto 'considered' harmful");

            ConfirmStringParse("");
            ConfirmStringParse("'");
            ConfirmStringParse("''");
            ConfirmStringParse("' '");
            ConfirmStringParse(" ' ");
        }
        [TestMethod]
        public void TestParseSumIfSum()
        {
            String formulaString;
            Ptg[] ptgs;
            ptgs = ParseFormula("sum(5, 2, if(3>2, sum(A1:A2), 6))");
            formulaString = HSSFFormulaParser.ToFormulaString(null, ptgs);
            Assert.AreEqual("SUM(5,2,IF(3>2,SUM(A1:A2),6))", formulaString);

            ptgs = ParseFormula("if(1<2,sum(5, 2, if(3>2, sum(A1:A2), 6)),4)");
            formulaString = HSSFFormulaParser.ToFormulaString(null, ptgs);
            Assert.AreEqual("IF(1<2,SUM(5,2,IF(3>2,SUM(A1:A2),6)),4)", formulaString);
        }
        [TestMethod]
        public void TestParserErrors()
        {
            ParseExpectedException("1 2");
            ParseExpectedException(" 12 . 345  ");
            ParseExpectedException("1 .23  ");

            ParseExpectedException("sum(#NAME)");
            ParseExpectedException("1 + #N / A * 2");
            ParseExpectedException("#value?");
            ParseExpectedException("#DIV/ 0+2");


            ParseExpectedException("IF(TRUE)");
            ParseExpectedException("countif(A1:B5, C1, D1)");
        }

        private static void ParseExpectedException(String formula)
        {
            try
            {
                ParseFormula(formula);
                throw new AssertFailedException("expected Parse exception");
            }
            catch (Exception e)
            {
                FormulaParserTestHelper.ConfirmParseException(e);
            }
        }
        [TestMethod]
        public void TestSetFormulaWithRowBeyond32768_Bug44539()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            cell.CellFormula = ("SUM(A32769:A32770)");
            if ("SUM(A-32767:A-32766)".Equals(cell.CellFormula))
            {
                Assert.Fail("Identified bug 44539");
            }
            Assert.AreEqual("SUM(A32769:A32770)", cell.CellFormula);
        }
        [TestMethod]
        public void TestSpaceAtStartOfFormula()
        {
            // Simulating cell formula of "= 4" (note space)
            // The same Ptg array can be observed if an excel file is1 saved with that exact formula

            AttrPtg spacePtg = AttrPtg.CreateSpace(AttrPtg.SpaceType.SPACE_BEFORE, 1);
            Ptg[] ptgs = { spacePtg, new IntPtg(4), };
            String formulaString;
            try
            {
                formulaString = HSSFFormulaParser.ToFormulaString(null, ptgs);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals("too much stuff left on the stack", StringComparison.InvariantCultureIgnoreCase))
                {
                    throw new AssertFailedException("Identified bug 44609");
                }
                // else some unexpected error
                throw e;
            }
            // FormulaParser strips spaces anyway
            Assert.AreEqual("4", formulaString);

            ptgs = new Ptg[] { new IntPtg(3), spacePtg, new IntPtg(4), spacePtg, AddPtg.instance, };
            formulaString = HSSFFormulaParser.ToFormulaString(null, ptgs);
            Assert.AreEqual("3+4", formulaString);
        }

        /**
         * Checks some internal error detecting logic ('stack underflow error' in toFormulaString)
         */
        [TestMethod]
        public void TestTooFewOperandArgs()
        {
            // Simulating badly encoded cell formula of "=/1"
            // Not sure if Excel could ever produce this
            Ptg[] ptgs = {
				// Excel would probably have put tMissArg here
				new IntPtg(1),
				DividePtg.instance,
		};
            try
            {
                HSSFFormulaParser.ToFormulaString(null, ptgs);
                Assert.Fail("Expected exception was not thrown");
            }
            catch (InvalidOperationException e)
            {
                // expected during successful test
                Assert.IsTrue(e.Message.StartsWith("Too few arguments supplied to operation"));
            }
        }
        /**
         * Make sure that POI uses the right Func Ptg when encoding formulas.  Functions with variable
         * number of args should Get FuncVarPtg, functions with fixed args should Get FuncPtg.<p/>
         * 
         * Prior to the fix for bug 44675 POI would encode FuncVarPtg for all functions.  In many cases
         * Excel tolerates the wrong Ptg and evaluates the formula OK (e.g. SIN), but in some cases 
         * (e.g. COUNTIF) Excel Assert.Fails to evaluate the formula, giving '#VALUE!' instead. 
         */
        [TestMethod]
        public void TestFuncPtgSelection()
        {

            Ptg[] ptgs;
            ptgs = ParseFormula("countif(A1:A2, 1)");
            Assert.AreEqual(3, ptgs.Length);
            if (typeof(FuncVarPtg) == ptgs[2].GetType())
            {
                throw new AssertFailedException("Identified bug 44675");
            }
            Assert.AreEqual(typeof(FuncPtg), ptgs[2].GetType());
            ptgs = ParseFormula("sin(1)");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(typeof(FuncPtg), ptgs[1].GetType());
        }
        [TestMethod]
        public void TestWrongNumberOfFunctionArgs()
        {
            ConfirmArgCountMsg("sin()", "Too few arguments to function 'SIN'. Expected 1 but got 0.");
            ConfirmArgCountMsg("countif(1, 2, 3, 4)", "Too many arguments to function 'COUNTIF'. Expected 2 but got 4.");
            ConfirmArgCountMsg("index(1, 2, 3, 4, 5, 6)", "Too many arguments to function 'INDEX'. At most 4 were expected but got 6.");
            ConfirmArgCountMsg("vlookup(1, 2)", "Too few arguments to function 'VLOOKUP'. At least 3 were expected but got 2.");
        }

        private static void ConfirmArgCountMsg(String formula, String expectedMessage)
        {
            HSSFWorkbook book = new HSSFWorkbook();
            try
            {
                HSSFFormulaParser.Parse(formula, book);
                throw new AssertFailedException("Didn't Get Parse exception as expected");
            }
            catch (Exception e)
            {
                Assert.AreEqual(expectedMessage, e.Message);
            }
        }
        [TestMethod]
        public void TestParseErrorExpecteMsg()
        {

            try
            {
                ParseFormula("round(3.14;2)");
                throw new AssertFailedException("Didn't Get Parse exception as expected");
            }
            catch (Exception e)
            {
                Assert.AreEqual("Parse error near char 10 ';' in specified formula 'round(3.14;2)'. Expected ',' or ')'", e.Message);
            }
        }
    }
}