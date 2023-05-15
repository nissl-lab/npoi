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
    using NUnit.Framework;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    using TestCases.HSSF.UserModel;
    using TestCases.SS.Formula;
    using NPOI.SS.Formula.Constant;
    using System.IO;
    using System.Globalization;

    /**
     * Test the low level formula Parser functionality. High level Tests are to
     * be done via usermodel/Cell.SetFormulaValue().
     */
    [TestFixture]
    public class TestFormulaParser
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void PrepareCulture()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }


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
        [Test]
        public void TestSimpleFormula()
        {
            Ptg[] ptgs = ParseFormula("2+2");
            Assert.AreEqual(3, ptgs.Length);
        }
        [Test]
        public void TestFormulaWithSpace1()
        {
            Ptg[] ptgs = ParseFormula(" 2 + 2 ");
            Assert.AreEqual(3, ptgs.Length);
            Assert.IsTrue((ptgs[0] is IntPtg));
            Assert.IsTrue((ptgs[1] is IntPtg));
            Assert.IsTrue((ptgs[2] is AddPtg));
        }
        [Test]
        public void TestFormulaWithSpace2()
        {
            Ptg[] ptgs = ParseFormula("2+ sum( 3 , 4) ");
            Assert.AreEqual(5, ptgs.Length);
        }
        [Test]
        public void TestFormulaWithSpaceNRef()
        {
            Ptg[] ptgs = ParseFormula("sum( A2:A3 )");
            Assert.AreEqual(2, ptgs.Length);
        }
        [Test]
        public void TestFormulaWithString()
        {
            Ptg[] ptgs = ParseFormula("\"hello\" & \"world\" ");
            Assert.AreEqual(3, ptgs.Length);
        }
        [Test]
        public void TestTRUE()
        {
            Ptg[] ptgs = ParseFormula("TRUE");
            Assert.AreEqual(1, ptgs.Length);
            BoolPtg flag = (BoolPtg)ptgs[0];
            Assert.AreEqual(true, flag.Value);
        }
        [Test]
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
        [Test]
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
        [Test]
        public void TestMacroFunction()
        {
            // testNames.xls contains a VB function called 'myFunc'
            String testFile = "testNames.xls";
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook(testFile);
            try
            {
                HSSFEvaluationWorkbook book = HSSFEvaluationWorkbook.Create(wb);

                //Expected ptg stack: [NamePtg(myFunc), StringPtg(arg), (additional operands go here...), FunctionPtg(myFunc)]
                Ptg[] ptg = FormulaParser.Parse("myFunc(\"arg\")", book, FormulaType.Cell, -1);
                Assert.AreEqual(3, ptg.Length);

                // the name gets encoded as the first operand on the stack
                NamePtg tname = (NamePtg)ptg[0];
                Assert.AreEqual("myFunc", tname.ToFormulaString(book));

                // the function's arguments are pushed onto the stack from left-to-right as OperandPtgs
                StringPtg arg = (StringPtg)ptg[1];
                Assert.AreEqual("arg", arg.Value);

                // The external FunctionPtg is the last Ptg added to the stack
                // During formula evaluation, this Ptg pops off the the appropriate number of
                // arguments (getNumberOfOperands()) and pushes the result on the stack
                AbstractFunctionPtg tfunc = (AbstractFunctionPtg)ptg[2]; //FuncVarPtg
                Assert.IsTrue(tfunc.IsExternalFunction);

                // confirm formula parsing is case-insensitive
                FormulaParser.Parse("mYfUnC(\"arg\")", book, FormulaType.Cell, -1);

                // confirm formula parsing doesn't care about argument count or type
                // this should only throw an error when evaluating the formula.
                FormulaParser.Parse("myFunc()", book, FormulaType.Cell, -1);
                FormulaParser.Parse("myFunc(\"arg\", 0, TRUE)", book, FormulaType.Cell, -1);

                // A completely unknown formula name (not saved in workbook) should still be parseable and renderable
                // but will throw an NotImplementedFunctionException or return a #NAME? error value if evaluated.
                FormulaParser.Parse("yourFunc(\"arg\")", book, FormulaType.Cell, -1);

                // Verify that myFunc and yourFunc were successfully added to Workbook names
                HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb);
                try
                {
                    // HSSFWorkbook/EXCEL97-specific side-effects user-defined function names must be added to Workbook's defined names in order to be saved.
                    Assert.IsNotNull(wb2.GetName("myFunc"));
                    assertEqualsIgnoreCase("myFunc", wb2.GetName("myFunc").NameName);
                    Assert.IsNotNull(wb2.GetName("yourFunc"));
                    assertEqualsIgnoreCase("yourFunc", wb2.GetName("yourFunc").NameName);

                    // Manually check to make sure file isn't corrupted
                    // TODO: develop a process for occasionally manually reviewing workbooks
                    // to verify workbooks are not corrupted
                    /*
                    FileInfo fileIn = HSSFTestDataSamples.GetSampleFile(testFile);
                    FileInfo reSavedFile = new FileInfo(fileIn.FullName.Replace(".xls", "-saved.xls"));
                    FileStream fos = new FileStream(reSavedFile.FullName, FileMode.Create, FileAccess.ReadWrite);
                    wb2.Write(fos);
                    fos.Close();
                    */
                }
                finally
                {
                    wb2.Close();
                }
            }
            finally
            {
                wb.Close();
            }
        }

        private static void assertEqualsIgnoreCase(String expected, String actual)
        {
            CultureInfo cultureUS = CultureInfo.GetCultureInfo("en-US");
            Assert.AreEqual(expected.ToLower(cultureUS), actual.ToLower(cultureUS));
        }

        [Test]
        public void TestEmbeddedSlash()
        {
            Ptg[] ptgs = ParseFormula("HYPERLINK(\"http://www.jakarta.org\",\"Jakarta\")");
            Assert.IsTrue(ptgs[0] is StringPtg, "first ptg is1 string");
            Assert.IsTrue(ptgs[1] is StringPtg, "second ptg is1 string");
        }
        [Test]
        public void TestConcatenate()
        {
            Ptg[] ptgs = ParseFormula("CONCATENATE(\"first\",\"second\")");
            Assert.IsTrue(ptgs[0] is StringPtg, "first ptg is1 string");
            Assert.IsTrue(ptgs[1] is StringPtg, "second ptg is1 string");
        }
        [Test]
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

            wb.Close();
        }
        [Test]
        public void TestUnaryMinus()
        {
            Ptg[] ptgs = ParseFormula("-A1");
            Assert.AreEqual(2, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "first ptg is1 reference");
            Assert.IsTrue(ptgs[1] is UnaryMinusPtg, "second ptg is1 Minus");
        }
        [Test]
        public void TestUnaryPlus()
        {
            Ptg[] ptgs = ParseFormula("+A1");
            Assert.AreEqual(2, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "first ptg is1 reference");
            Assert.IsTrue(ptgs[1] is UnaryPlusPtg, "second ptg is1 Plus");
        }
        [Test]
        public void TestLeadingSpaceInString()
        {
            String value = "  hi  ";
            Ptg[] ptgs = ParseFormula("\"" + value + "\"");

            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is StringPtg, "ptg0 is1 a StringPtg");
            Assert.IsTrue(((StringPtg)ptgs[0]).Value.Equals(value), "ptg0 contains exact value");
        }
        [Test]
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
        [Test]
        public void TestLargeInt()
        {
            Ptg[] ptgs = ParseFormula("40");
            Assert.IsTrue(ptgs[0] is IntPtg, "ptg is1 Int, is1 " + ptgs[0].GetType());

            ptgs = ParseFormula("40000");
            Assert.IsTrue(ptgs[0] is IntPtg, "ptg should be  IntPtg, is1 " + ptgs[0].GetType());
        }

        /** bug 33160, Testcase by Amol Deshmukh*/
        [Test]
        public void TestSimpleLongFormula()
        {
            Ptg[] ptgs = ParseFormula("40000/2");
            Assert.AreEqual(3, ptgs.Length);
            Assert.IsTrue((ptgs[0] is IntPtg), "IntPtg");
            Assert.IsTrue((ptgs[1] is IntPtg), "IntPtg");
            Assert.IsTrue((ptgs[2] is DividePtg), "DividePtg");
        }

        /** bug 35027, underscore in sheet name */
        [Test]
        public void TestUnderscore()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            wb.CreateSheet("Cash_Flow");

            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell;

            cell = row.CreateCell((short)0);
            cell.CellFormula = ("Cash_Flow!A1");

            wb.Close();
        }
        /** bug 49725, defined names with underscore */
        [Test]
        public void TestNamesWithUnderscore()
        {
            HSSFWorkbook wb = new HSSFWorkbook(); //or new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("NamesWithUnderscore");

            IName nm;

            nm = wb.CreateName();
            nm.NameName = ("DA6_LEO_WBS_Number");
            nm.RefersToFormula = ("33");

            nm = wb.CreateName();
            nm.NameName = ("DA6_LEO_WBS_Name");
            nm.RefersToFormula = ("33");

            nm = wb.CreateName();
            nm.NameName = ("A1_");
            nm.RefersToFormula = ("22");

            nm = wb.CreateName();
            nm.NameName = ("_A1");
            nm.RefersToFormula = ("11");

            nm = wb.CreateName();
            nm.NameName = ("A_1");
            nm.RefersToFormula = ("44");

            nm = wb.CreateName();
            nm.NameName = ("A_1_");
            nm.RefersToFormula = ("44");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            cell.CellFormula = ("DA6_LEO_WBS_Number*2");
            Assert.AreEqual("DA6_LEO_WBS_Number*2", cell.CellFormula);

            cell.CellFormula = ("(A1_*_A1+A_1)/A_1_");
            Assert.AreEqual("(A1_*_A1+A_1)/A_1_", cell.CellFormula);

            cell.CellFormula = ("INDEX(DA6_LEO_WBS_Name,MATCH($A3,DA6_LEO_WBS_Number,0))");
            Assert.AreEqual("INDEX(DA6_LEO_WBS_Name,MATCH($A3,DA6_LEO_WBS_Number,0))", cell.CellFormula);

            wb.Close();
        }
        // bug 38396 : Formula with exponential numbers not Parsed correctly.
        [Test]
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
        /// <summary>
        /// Tests the exponential in sheet.
        /// </summary>
        [Test]
        public void TestExponentialInSheet()
        {
            // This Test depends on the american culture.
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

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
            Assert.AreEqual("0.000000000000013/3", formula, "Exponential formula string");

            cell.CellFormula = ("-13E-15/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-0.000000000000013/3", formula, "Exponential formula string");

            cell.CellFormula = ("1.3E3/3");
            formula = cell.CellFormula;
            Assert.AreEqual("1300/3", formula, "Exponential formula string");

            cell.CellFormula = ("-1.3E3/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-1300/3", formula, "Exponential formula string");

            cell.CellFormula = ("1300000000000000/3");
            formula = cell.CellFormula;
            Assert.AreEqual("1300000000000000/3", formula, "Exponential formula string");

            cell.CellFormula = ("-1300000000000000/3");
            formula = cell.CellFormula;
            Assert.AreEqual("-1300000000000000/3", formula, "Exponential formula string");

            cell.CellFormula = ("-10E-1/3.1E2*4E3/3E4");
            formula = cell.CellFormula;
            Assert.AreEqual("-1/310*4000/30000", formula, "Exponential formula string");

            wb.Close();
        }
        [Test]
        public void TestNumbers()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");

            HSSFWorkbook wb = new HSSFWorkbook();

            wb.CreateSheet("Cash_Flow");

            ISheet sheet = wb.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            String formula = null;

            // starts from decimal point

            cell.CellFormula = (".1");
            formula = cell.CellFormula;
            Assert.AreEqual("0.1", formula);

            cell.CellFormula = ("+.1");
            formula = cell.CellFormula;
            Assert.AreEqual("0.1", formula);

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

            wb.Close();
        }
        [Test]
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

            wb.Close();
        }

        [Test]
        public void TestMultiSheetReference()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            wb.CreateSheet("Cash_Flow");
            wb.CreateSheet("Test Sheet");

            HSSFSheet sheet = wb.CreateSheet("Test") as HSSFSheet;
            HSSFRow row = sheet.CreateRow(0) as HSSFRow;
            HSSFCell cell = row.CreateCell(0) as HSSFCell;
            String formula = null;

            // References to a single cell:

            // One sheet
            cell.CellFormula = (/*setter*/"Cash_Flow!A1");
            formula = cell.CellFormula;
            Assert.AreEqual("Cash_Flow!A1", formula);

            // Then the other
            cell.CellFormula = (/*setter*/"\'Test Sheet\'!A1");
            formula = cell.CellFormula;
            Assert.AreEqual("\'Test Sheet\'!A1", formula);

            // Now both
            cell.CellFormula = (/*setter*/"Cash_Flow:\'Test Sheet\'!A1");
            formula = cell.CellFormula;
            Assert.AreEqual("Cash_Flow:\'Test Sheet\'!A1", formula);

            // References to a range (area) of cells:

            // One sheet
            cell.CellFormula = ("Cash_Flow!A1:B2");
            formula = cell.CellFormula;
            Assert.AreEqual("Cash_Flow!A1:B2", formula);

            // Then the other
            cell.CellFormula = ("\'Test Sheet\'!A1:B2");
            formula = cell.CellFormula;
            Assert.AreEqual("\'Test Sheet\'!A1:B2", formula);

            // Now both
            cell.CellFormula = ("Cash_Flow:\'Test Sheet\'!A1:B2");
            formula = cell.CellFormula;
            Assert.AreEqual("Cash_Flow:\'Test Sheet\'!A1:B2", formula);

            wb.Close();
        }


        /**
         * Test for bug observable at svn revision 618865 (5-Feb-2008)<br/>
         * a formula consisting of a single no-arg function got rendered without the function braces
         */
        [Test]
        public void TestToFormulaStringZeroArgFunction()
        {
            HSSFWorkbook book = new HSSFWorkbook();

            Ptg[] ptgs = {
                FuncPtg.Create(10),
        };
            Assert.AreEqual("NA()", HSSFFormulaParser.ToFormulaString(book, ptgs));

            book.Close();
        }
        [Test]
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
        [Test]
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
        public static Ptg[] ConfirmTokenClasses(String formula, params Type[] expectedClasses)
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


        /**
         * There may be multiple ways to encode an expression involving {@link UnaryPlusPtg}
         * or {@link UnaryMinusPtg}.  These may be perfectly equivalent from a formula
         * evaluation perspective, or formula rendering.  However, differences in the way
         * POI encodes formulas may cause unnecessary confusion.  These non-critical tests
         * check that POI follows the same encoding rules as Excel.
         */
        [Test]
        public void TestExactEncodingOfUnaryPlusAndMinus()
        {
            // as tested in Excel:
            ConfirmUnary("-3", -3, typeof(NumberPtg));
            ConfirmUnary("--4", -4, typeof(NumberPtg), typeof(UnaryMinusPtg));
            ConfirmUnary("+++5", 5, typeof(IntPtg), typeof(UnaryPlusPtg), typeof(UnaryPlusPtg));
            ConfirmUnary("++-6", -6, typeof(NumberPtg), typeof(UnaryPlusPtg), typeof(UnaryPlusPtg));

            // Spaces muck things up a bit.  It would be clearer why the following cases are
            // reasonable if POI encoded tAttrSpace in the right places.
            // Otherwise these differences look capricious.
            ConfirmUnary("+ 12", 12, typeof(IntPtg), typeof(UnaryPlusPtg));
            ConfirmUnary("- 13", 13, typeof(IntPtg), typeof(UnaryMinusPtg));
        }
        private static void ConfirmUnary(String formulaText, double val, params Type[] expectedTokenTypes)
        {
            Ptg[] ptgs = ParseFormula(formulaText);
            ConfirmTokenClasses(ptgs, expectedTokenTypes);
            Ptg ptg0 = ptgs[0];
            if (ptg0 is IntPtg)
            {
                IntPtg intPtg = (IntPtg)ptg0;
                Assert.AreEqual((int)val, intPtg.Value);
            }
            else if (ptg0 is NumberPtg)
            {
                NumberPtg numberPtg = (NumberPtg)ptg0;
                Assert.AreEqual(val, numberPtg.Value, 0.0);
            }
            else
            {
                Assert.Fail("bad ptg0 " + ptg0);
            }
        }
        [Test]
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
        [Test]
        public void TestParseNumber()
        {
            // This Test depends on the american culture.
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

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
        [Test]
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
        [Test]
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
        [Test]
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
                    Assert.Fail("Identified bug 28754a");
                }
                throw;
            }
            Assert.AreEqual("test\"ing", sp.Value);

            HSSFWorkbook wb = new HSSFWorkbook();
            try
            {
                NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet();
                wb.SetSheetName(0, "Sheet1");

                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell((short)0);
                cell.CellFormula = ("right(\"test\"\"ing\", 3)");
                String actualCellFormula = cell.CellFormula;
                if ("RIGHT(\"test\"ing\",3)".Equals(actualCellFormula))
                {
                    Assert.Fail("Identified bug 28754b");
                }
                Assert.AreEqual("RIGHT(\"test\"\"ing\",3)", actualCellFormula);
            }
            finally
            {
                wb.Close();
            }
        }
        [Test]
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
        [Test]
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
        [Test]
        public void TestParserErrors()
        {
            ParseExpectedException(" 12 . 345  ");
            ParseExpectedException("1 .23  ");

            ParseExpectedException("sum(#NAME)");
            ParseExpectedException("1 + #N / A * 2");
            ParseExpectedException("#value?");
            ParseExpectedException("#DIV/ 0+2");


            ParseExpectedException("IF(TRUE)");
            ParseExpectedException("countif(A1:B5, C1, D1)");

            ParseExpectedException("(");
            ParseExpectedException(")");
            ParseExpectedException("+");
            ParseExpectedException("42+");

            ParseExpectedException("IF(");
        }

        private static void ParseExpectedException(String formula)
        {
            try
            {
                ParseFormula(formula);
                Assert.Fail("Expected FormulaParseException: " + formula);
            }
            catch (FormulaParseException e)
            {
                // expected during successful test
                Assert.IsNotNull(e.Message);
            }
        }
        [Test]
        public void TestSetFormulaWithRowBeyond32768_Bug44539()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            wb.SetSheetName(0, "Sheet1");

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell((short)0);
            cell.CellFormula = ("SUM(A32769:A32770)");
            if ("SUM(A-32767:A-32766)".Equals(cell.CellFormula))
            {
                Assert.Fail("Identified bug 44539");
            }
            Assert.AreEqual("SUM(A32769:A32770)", cell.CellFormula);

            wb.Close();
        }
        [Test]
        public void TestSpaceAtStartOfFormula()
        {
            // Simulating cell formula of "= 4" (note space)
            // The same Ptg array can be observed if an excel file is1 saved with that exact formula

            AttrPtg spacePtg = AttrPtg.CreateSpace(AttrPtg.SpaceType.SpaceBefore, 1);
            Ptg[] ptgs = { spacePtg, new IntPtg(4), };
            String formulaString;
            try
            {
                formulaString = HSSFFormulaParser.ToFormulaString(null, ptgs);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals("too much stuff left on the stack", StringComparison.OrdinalIgnoreCase))
                {
                    Assert.Fail("Identified bug 44609");
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
        [Test]
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
                // expected during successful Test
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
        [Test]
        public void TestFuncPtgSelection()
        {

            Ptg[] ptgs;
            ptgs = ParseFormula("countif(A1:A2, 1)");
            Assert.AreEqual(3, ptgs.Length);
            if (typeof(FuncVarPtg) == ptgs[2].GetType())
            {
                Assert.Fail("Identified bug 44675");
            }
            Assert.AreEqual(typeof(FuncPtg), ptgs[2].GetType());
            ptgs = ParseFormula("sin(1)");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(typeof(FuncPtg), ptgs[1].GetType());
        }
        [Test]
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
                Assert.Fail("Didn't Get Parse exception as expected");
            }
            catch (Exception e)
            {
                Assert.AreEqual(expectedMessage, e.Message);
            }
        }

        [Test]
        public void TestRange_bug46643()
        {
            String formula = "Sheet1!A1:Sheet1!B3";
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");
            Ptg[] ptgs = FormulaParser.Parse(formula, HSSFEvaluationWorkbook.Create(wb), FormulaType.Cell, -1, -1);

            if (ptgs.Length == 3)
            {
                ConfirmTokenClasses(ptgs, typeof(Ref3DPtg), typeof(Ref3DPtg), typeof(RangePtg));
                Assert.Fail("Identified bug 46643");
            }

            ConfirmTokenClasses(ptgs,
                    typeof(MemFuncPtg),
                    typeof(Ref3DPtg),
                    typeof(Ref3DPtg),
                    typeof(RangePtg)
            );
            MemFuncPtg mf = (MemFuncPtg)ptgs[0];
            Assert.AreEqual(15, mf.LenRefSubexpression);
        }
        /* package */
        private static void ConfirmTokenClasses(Ptg[] ptgs, params Type[] expectedClasses)
        {
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
        }
        [Test]
        public void TestParseErrorExpectedMsg()
        {

            try
            {
                ParseFormula("round(3.14;2)");
                Assert.Fail("Didn't get parse exception as expected");
            }
            catch (FormulaParseException e)
            {
                ConfirmParseException(e,
                        "Parse error near char 10 ';' in specified formula 'round(3.14;2)'. Expected ',' or ')'");
            }

            try
            {
                ParseFormula(" =2+2");
                Assert.Fail("Didn't get parse exception as expected");
            }
            catch (FormulaParseException e)
            {
                ConfirmParseException(e,
                        "The specified formula ' =2+2' starts with an equals sign which is not allowed.");
            }
        }

        /**
         * this function name has a dot in it.
         */
        [Test]
        public void TestParseErrorTypeFunction()
        {

            Ptg[] ptgs;
            try
            {
                ptgs = ParseFormula("error.type(A1)");
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("Invalid Formula cell reference: 'error'"))
                {
                    Assert.Fail("Identified bug 45334");
                }
                throw e;
            }
            ConfirmTokenClasses(ptgs, typeof(RefPtg), typeof(FuncPtg));
            Assert.AreEqual("ERROR.TYPE", ((FuncPtg)ptgs[1]).Name);
        }
        [Test]
        public void TestNamedRangeThatLooksLikeCell()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IName name = wb.CreateName();
            name.RefersToFormula = ("Sheet1!B1");
            name.NameName = ("pfy1");

            Ptg[] ptgs;
            try
            {
                ptgs = HSSFFormulaParser.Parse("count(pfy1)", wb);
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("Specified colIx (1012) is out of range"))
                {
                    Assert.Fail("Identified bug 45354");
                }
                wb.Close();
                throw e;
            }
            ConfirmTokenClasses(ptgs, typeof(NamePtg), typeof(FuncVarPtg));

            ICell cell = sheet.CreateRow(0).CreateCell(0);
            cell.CellFormula = ("count(pfy1)");
            Assert.AreEqual("COUNT(pfy1)", cell.CellFormula);
            try
            {
                cell.CellFormula = ("count(pf1)");
                Assert.Fail("Expected formula parse execption");
            }
            catch (FormulaParseException e)
            {
                ConfirmParseException(e,
                        "Specified named range 'pf1' does not exist in the current workbook.");
            }
            cell.CellFormula = ("count(fp1)"); // plain cell ref, col is in range
            wb.Close();
        }
        [Test]
        public void TestParseAreaRefHighRow_bug45358()
        {
            Ptg[] ptgs;
            AreaI aptg;

            HSSFWorkbook book = new HSSFWorkbook();
            book.CreateSheet("Sheet1");

            ptgs = HSSFFormulaParser.Parse("Sheet1!A10:A40000", book);
            aptg = (AreaI)ptgs[0];
            if (aptg.LastRow == -25537)
            {
                Assert.Fail("Identified bug 45358");
            }
            Assert.AreEqual(39999, aptg.LastRow);

            ptgs = HSSFFormulaParser.Parse("Sheet1!A10:A65536", book);
            aptg = (AreaI)ptgs[0];
            Assert.AreEqual(65535, aptg.LastRow);

            // plain area refs should be ok too
            ptgs = ParseFormula("A10:A65536");
            aptg = (AreaI)ptgs[0];
            Assert.AreEqual(65535, aptg.LastRow);
            book.Close();
        }
        [Test]
        public void TestParseArray()
        {
            Ptg[] ptgs;
            ptgs = ParseFormula("mode({1,2,2,#REF!;FALSE,3,3,2})");
            ConfirmTokenClasses(ptgs, typeof(ArrayPtg), typeof(FuncVarPtg));
            Assert.AreEqual("{1,2,2,#REF!;FALSE,3,3,2}", ptgs[0].ToFormulaString());

            ArrayPtg aptg = (ArrayPtg)ptgs[0];
            Object[,] values = aptg.GetTokenArrayValues();
            Assert.AreEqual(ErrorConstant.ValueOf(FormulaError.REF.Code), values[0, 3]);
            Assert.AreEqual(false, values[1, 0]);
        }
        [Test]
        public void TestParseStringElementInArray()
        {
            Ptg[] ptgs;
            ptgs = ParseFormula("MAX({\"5\"},3)");
            ConfirmTokenClasses(ptgs, typeof(ArrayPtg), typeof(IntPtg), typeof(FuncVarPtg));
            object element = ((ArrayPtg)ptgs[0]).GetTokenArrayValues()[0, 0];
            if (element is UnicodeString)
            {
                // this would cause ClassCastException below
                Assert.Fail("Wrong encoding of array element value");
            }
            Assert.AreEqual(typeof(String), element.GetType());

            // make sure the formula encodes OK
            int encSize = Ptg.GetEncodedSize(ptgs);
            byte[] data = new byte[encSize];
            Ptg.SerializePtgs(ptgs, data, 0);
            byte[] expData = HexRead.ReadFromString(
                    "20 00 00 00 00 00 00 00 " // tArray
                    + "1E 03 00 "      // tInt(3)
                    + "42 02 07 00 "   // tFuncVar(MAX) 2-arg
                    + "00 00 00 "      // Array data: 1 col, 1 row
                    + "02 01 00 00 35" // elem (type=string, len=1, "5")
            );
            Assert.IsTrue(Arrays.Equals(expData, data));
            int initSize = Ptg.GetEncodedSizeWithoutArrayData(ptgs);
            Ptg[] ptgs2 = Ptg.ReadTokens(initSize, new LittleEndianByteArrayInputStream(data));
            ConfirmTokenClasses(ptgs2, typeof(ArrayPtg), typeof(IntPtg), typeof(FuncVarPtg));
        }
        [Test]
        public void TestParseArrayNegativeElement()
        {
            Ptg[] ptgs;
            try
            {
                ptgs = ParseFormula("{-42}");
            }
            catch (FormulaParseException e)
            {
                if (e.Message.StartsWith("Parse error near char 1 '-' in specified formula '{-42}'. Expected ")) // "Integer" in Java or "int" in C#
                {
                    Assert.Fail("Identified bug - failed to parse negative array element.");
                }
                throw e;
            }
            ConfirmTokenClasses(ptgs, typeof(ArrayPtg));
            Object element = ((ArrayPtg)ptgs[0]).GetTokenArrayValues()[0, 0];

            Assert.AreEqual(-42.0, (Double)element, 0.0);

            // Should be able to handle whitespace between unary minus and digits (Excel
            // accepts this formula after presenting the user with a Confirmation dialog).
            ptgs = ParseFormula("{- 5}");
            element = ((ArrayPtg)ptgs[0]).GetTokenArrayValues()[0, 0];
            Assert.AreEqual(-5.0, (Double)element, 0.0);
        }
        [Test]
        public void TestRangeOperator()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            ICell cell = sheet.CreateRow(0).CreateCell(0);

            wb.SetSheetName(0, "Sheet1");
            cell.CellFormula = ("Sheet1!B$4:Sheet1!$C1"); // explicit range ':' operator
            Assert.AreEqual("Sheet1!B$4:Sheet1!$C1", cell.CellFormula);

            cell.CellFormula = ("Sheet1!B$4:$C1"); // plain area ref
            Assert.AreEqual("Sheet1!B1:$C$4", cell.CellFormula); // note - area ref is normalised

            cell.CellFormula = ("Sheet1!$C1...B$4"); // different syntax for plain area ref
            Assert.AreEqual("Sheet1!B1:$C$4", cell.CellFormula);

            // with funny sheet name
            wb.SetSheetName(0, "A1...A2");
            cell.CellFormula = ("A1...A2!B1");
            Assert.AreEqual("A1...A2!B1", cell.CellFormula);

            wb.Close();
        }
        [Test]
        public void TestBooleanNamedSheet()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("true");
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            cell.CellFormula = ("'true'!B2");

            Assert.AreEqual("'true'!B2", cell.CellFormula);

            wb.Close();
        }
        [Test]
        public void TestParseExternalWorkbookReference()
        {
            HSSFWorkbook wbA = HSSFTestDataSamples.OpenSampleWorkbook("multibookFormulaA.xls");
            ICell cell = wbA.GetSheetAt(0).GetRow(0).GetCell(0);

            // make sure formula in sample is as expected
            Assert.AreEqual("[multibookFormulaB.xls]BSheet1!B1", cell.CellFormula);
            Ptg[] expectedPtgs = FormulaExtractor.GetPtgs(cell);
            ConfirmSingle3DRef(expectedPtgs, 1);

            // now try (re-)parsing the formula
            Ptg[] actualPtgs = HSSFFormulaParser.Parse("[multibookFormulaB.xls]BSheet1!B1", wbA);
            ConfirmSingle3DRef(actualPtgs, 1); // externalSheetIndex 1 -> BSheet1

            // try parsing a formula pointing to a different external sheet
            Ptg[] otherPtgs = HSSFFormulaParser.Parse("[multibookFormulaB.xls]AnotherSheet!B1", wbA);
            ConfirmSingle3DRef(otherPtgs, 0); // externalSheetIndex 0 -> AnotherSheet

            // try setting the same formula in a cell
            cell.CellFormula = ("[multibookFormulaB.xls]AnotherSheet!B1");
            Assert.AreEqual("[multibookFormulaB.xls]AnotherSheet!B1", cell.CellFormula);

            wbA.Close();
        }
        private static void ConfirmSingle3DRef(Ptg[] ptgs, int expectedExternSheetIndex)
        {
            Assert.AreEqual(1, ptgs.Length);
            Ptg ptg0 = ptgs[0];
            Assert.IsTrue(ptg0 is Ref3DPtg);
            Assert.AreEqual(expectedExternSheetIndex, ((Ref3DPtg)ptg0).ExternSheetIndex);
        }
        [Test]
        public void TestUnion()
        {
            String formula = "Sheet1!$B$2:$C$3,OFFSET(Sheet1!$E$2:$E$4,1,Sheet1!$A$1),Sheet1!$D$6";
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");
            Ptg[] ptgs = FormulaParser.Parse(formula, HSSFEvaluationWorkbook.Create(wb), FormulaType.Cell, -1, -1);

            ConfirmTokenClasses(ptgs,
                    // TODO - AttrPtg), // Excel prepends this
                    typeof(MemFuncPtg),
                    typeof(Area3DPtg),
                    typeof(Area3DPtg),
                    typeof(IntPtg),
                    typeof(Ref3DPtg),
                    typeof(FuncVarPtg),
                    typeof(UnionPtg),
                    typeof(Ref3DPtg),
                    typeof(UnionPtg)
            );
            MemFuncPtg mf = (MemFuncPtg)ptgs[0];
            Assert.AreEqual(45, mf.LenRefSubexpression);

            // We don't check the type of the operands.
            ConfirmTokenClasses("1,2", typeof(MemAreaPtg), typeof(IntPtg), typeof(IntPtg), typeof(UnionPtg));

            wb.Close();
        }

        [Test]
        public void TestIntersection()
        {
            String formula = "Sheet1!$B$2:$C$3 OFFSET(Sheet1!$E$2:$E$4, 1,Sheet1!$A$1) Sheet1!$D$6";
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");
            Ptg[] ptgs = FormulaParser.Parse(formula, HSSFEvaluationWorkbook.Create(wb), FormulaType.Cell, -1);
            ConfirmTokenClasses(ptgs,
                // TODO - AttrPtg.class, // Excel prepends this
                typeof(MemFuncPtg),
                typeof(Area3DPtg),
                typeof(Area3DPtg),
                typeof(IntPtg),
                typeof(Ref3DPtg),
                typeof(FuncVarPtg),
                typeof(IntersectionPtg),
                typeof(Ref3DPtg),
                typeof(IntersectionPtg)
        );
            MemFuncPtg mf = (MemFuncPtg)ptgs[0];
            Assert.AreEqual(45, mf.LenRefSubexpression);
            // This used to be an error but now parses.  Union has the same behaviour.
            ConfirmTokenClasses("1 2", typeof(MemAreaPtg), typeof(IntPtg), typeof(IntPtg), typeof(IntersectionPtg));

            wb.Close();
        }
        [Test]
        public void TestComparisonInParen()
        {
            ConfirmTokenClasses("(A1 > B2)",
                typeof(RefPtg),
                typeof(RefPtg),
                typeof(GreaterThanPtg),
                typeof(ParenthesisPtg)
            );
        }
        [Test]
        public void TestUnionInParen()
        {
            ConfirmTokenClasses("(A1:B2,B2:C3)",
              typeof(MemAreaPtg),
                  typeof(AreaPtg),
                  typeof(AreaPtg),
                  typeof(UnionPtg),
                  typeof(ParenthesisPtg)
                );
        }
        [Test]
        public void TestIntersectionInParen()
        {
            ConfirmTokenClasses("(A1:B2 B2:C3)",
                typeof(MemAreaPtg),
                    typeof(AreaPtg),
                    typeof(AreaPtg),
                    typeof(IntersectionPtg),
                    typeof(ParenthesisPtg)
                );
        }


        /** Named ranges with backslashes, e.g. 'POI\\2009' */
        [Test]
        public void TestBackSlashInNames()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            IName name = wb.CreateName();
            name.NameName = ("POI\\2009");
            name.RefersToFormula = ("Sheet1!$A$1");

            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);

            ICell cell_C1 = row.CreateCell(2);
            cell_C1.CellFormula = ("POI\\2009");
            Assert.AreEqual("POI\\2009", cell_C1.CellFormula);

            ICell cell_D1 = row.CreateCell(2);
            cell_D1.CellFormula = ("NOT(POI\\2009=\"3.5-final\")");
            Assert.AreEqual("NOT(POI\\2009=\"3.5-final\")", cell_D1.CellFormula);

            wb.Close();
        }

        /**
         * TODO - delete equiv Test:
         * {@link BaseTestBugzillaIssues#test42448()}
         */
        [Test]
        public void TestParseAbnormalSheetNamesAndRanges_bug42448()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("A");
            try
            {
                HSSFFormulaParser.Parse("SUM(A!C7:A!C67)", wb);
            }
            catch (IndexOutOfRangeException)
            {
                Assert.Fail("Identified bug 42448");
            }
            // the exact example from the bugzilla description:
            HSSFFormulaParser.Parse("SUMPRODUCT(A!C7:A!C67, B8:B68) / B69", wb);

            wb.Close();
        }
        [Test]
        public void TestRangeFuncOperand_bug46951()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Ptg[] ptgs;
            try
            {
                ptgs = HSSFFormulaParser.Parse("SUM(C1:OFFSET(C1,0,B1))", wb);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Specified named range 'OFFSET' does not exist in the current workbook."))
                {
                    Assert.Fail("Identified bug 46951");
                }
                wb.Close();
                throw;
            }
            ConfirmTokenClasses(ptgs,
                typeof(MemFuncPtg), // [len=23]
                typeof(RefPtg), // [C1]
                typeof(RefPtg), // [C1]
                typeof(IntPtg), // [0]
                typeof(RefPtg), // [B1]
                typeof(FuncVarPtg), // [OFFSET nArgs=3]
                typeof(RangePtg), //
                typeof(AttrPtg) // [sum ]
            );
            wb.Close();
        }
        [Test]
        public void TestUnionOfFullCollFullRowRef()
        {
            Ptg[] ptgs;
            ptgs = ParseFormula("3:4");
            ptgs = ParseFormula("$Z:$AC");
            ConfirmTokenClasses(ptgs, typeof(AreaPtg));
            ptgs = ParseFormula("B:B");

            ptgs = ParseFormula("$11:$13");
            ConfirmTokenClasses(ptgs, typeof(AreaPtg));

            ptgs = ParseFormula("$A:$A,$1:$4");
            ConfirmTokenClasses(ptgs, typeof(MemAreaPtg),
                    typeof(AreaPtg),
                    typeof(AreaPtg),
                    typeof(UnionPtg)
            );

            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");
            ptgs = HSSFFormulaParser.Parse("Sheet1!$A:$A,Sheet1!$1:$4", wb);
            ConfirmTokenClasses(ptgs, typeof(MemFuncPtg),
                    typeof(Area3DPtg),
                    typeof(Area3DPtg),
                    typeof(UnionPtg)
            );

            ptgs = HSSFFormulaParser.Parse("'Sheet1'!$A:$A,'Sheet1'!$1:$4", wb);
            ConfirmTokenClasses(ptgs,
                    typeof(MemFuncPtg),
                    typeof(Area3DPtg),
                    typeof(Area3DPtg),
                    typeof(UnionPtg)
            );
            wb.Close();
        }

        [Test]
        public void TestExplicitRangeWithTwoSheetNames()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");
            Ptg[] ptgs = HSSFFormulaParser.Parse("Sheet1!F1:Sheet1!G2", wb);
            ConfirmTokenClasses(ptgs,
                    typeof(MemFuncPtg),
                    typeof(Ref3DPtg),
                    typeof(Ref3DPtg),
                    typeof(RangePtg)
            );
            MemFuncPtg mf;
            mf = (MemFuncPtg)ptgs[0];
            Assert.AreEqual(15, mf.LenRefSubexpression);
            wb.Close();
        }

        /**
         * Checks that the area-ref and explicit range operators get the right associativity
         * and that the {@link MemFuncPtg} / {@link MemAreaPtg} is added correctly
         */
        [Test]
        public void TestComplexExplicitRangeEncodings()
        {
            Ptg[] ptgs;
            ptgs = ParseFormula("SUM(OFFSET(A1,0,0):B2:C3:D4:E5:OFFSET(F6,1,1):G7)");
            ConfirmTokenClasses(ptgs,
                // AttrPtg), // [volatile ] // POI doesn't do this yet (Apr 2009)
                typeof(MemFuncPtg), // len 57
                typeof(RefPtg), // [A1]
                typeof(IntPtg), // [0]
                typeof(IntPtg), // [0]
                typeof(FuncVarPtg), // [OFFSET nArgs=3]
                typeof(AreaPtg), // [B2:C3]
                typeof(RangePtg),
                typeof(AreaPtg), // [D4:E5]
                typeof(RangePtg),
                typeof(RefPtg), // [F6]
                typeof(IntPtg), // [1]
                typeof(IntPtg), // [1]
                typeof(FuncVarPtg), // [OFFSET nArgs=3]
                typeof(RangePtg),
                typeof(RefPtg), // [G7]
                typeof(RangePtg),
                typeof(AttrPtg) // [sum ]
            );

            MemFuncPtg mf = (MemFuncPtg)ptgs[0];
            Assert.AreEqual(57, mf.LenRefSubexpression);
            Assert.AreEqual("D4:E5", ((AreaPtgBase)ptgs[7]).ToFormulaString());
            Assert.IsTrue(((AttrPtg)ptgs[16]).IsSum);

            ptgs = ParseFormula("SUM(A1:B2:C3:D4)");
            ConfirmTokenClasses(ptgs,
                    // AttrPtg), // [volatile ] // POI doesn't do this yet (Apr 2009)
                    typeof(MemAreaPtg), // len 19
                    typeof(AreaPtg), // [A1:B2]
                    typeof(AreaPtg), // [C3:D4]
                    typeof(RangePtg),
                    typeof(AttrPtg) // [sum ]
            );
            MemAreaPtg ma = (MemAreaPtg)ptgs[0];
            Assert.AreEqual(19, ma.LenRefSubexpression);
        }


        /**
         * Mostly Confirming that erroneous conditions are detected.  Actual error message wording is not critical.
         *
         */
        [Test]
        public void TestEdgeCaseParserErrors()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");

            ConfirmParseError(wb, "A1:ROUND(B1,1)", "The RHS of the range operator ':' at position 3 is not a proper reference.");

            ConfirmParseError(wb, "Sheet1!!!", "Parse error near char 7 '!' in specified formula 'Sheet1!!!'. Expected number, string, defined name, or data table");
            ConfirmParseError(wb, "Sheet1!.Name", "Parse error near char 7 '.' in specified formula 'Sheet1!.Name'. Expected number, string, defined name, or data table");
            ConfirmParseError(wb, "Sheet1!Sheet1", "Specified name 'Sheet1' for sheet Sheet1 not found");

            ConfirmParseError(wb, "Sheet1!F:Sheet1!G", "'Sheet1!F' is not a proper reference.");
            ConfirmParseError(wb, "Sheet1!F..foobar", "Complete area reference expected after sheet name at index 11.");
            ConfirmParseError(wb, "Sheet1!A .. B", "Dotted range (full row or column) expression 'A .. B' must not contain whitespace.");
            ConfirmParseError(wb, "Sheet1!A...B", "Dotted range (full row or column) expression 'A...B' must have exactly 2 dots.");
            ConfirmParseError(wb, "Sheet1!A foobar", "Second part of cell reference expected after sheet name at index 10.");

            ConfirmParseError(wb, "foobar", "Specified named range 'foobar' does not exist in the current workbook.");
            ConfirmParseError(wb, "A1:1", "The RHS of the range operator ':' at position 3 is not a proper reference.");
        }
        private static void ConfirmParseError(HSSFWorkbook wb, String formula, String expectedMessage)
        {

            try
            {
                HSSFFormulaParser.Parse(formula, wb);
                Assert.Fail("Expected formula parse execption");
            }
            catch (FormulaParseException e)
            {
                ConfirmParseException(e, expectedMessage);
            }
        }

        /**
         * In bug 47078, POI had trouble evaluating a defined name flagged as 'complex'.
         * POI should also be able to parse such defined names.
         */
        [Test]
        public void TestParseComplexName()
        {

            // Mock up a spreadsheet to match the critical details of the sample
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");
            IName definedName = wb.CreateName();
            definedName.NameName = ("foo");
            definedName.RefersToFormula = ("Sheet1!B2");

            // Set the complex flag - POI doesn't usually manipulate this flag
            NameRecord nameRec = TestHSSFName.GetNameRecord(definedName);
            nameRec.OptionFlag = ((short)0x10); // 0x10 -> complex

            Ptg[] result;
            try
            {
                result = HSSFFormulaParser.Parse("1+foo", wb);
            }
            catch (FormulaParseException e)
            {
                if (e.Message.Equals("Specified name 'foo' is not a range as expected."))
                {
                    Assert.Fail("Identified bug 47078c");
                }
                wb.Close();
                throw e;
            }
            ConfirmTokenClasses(result, typeof(IntPtg), typeof(NamePtg), typeof(AddPtg));
            wb.Close();
        }

        /**
         * Zero is not a valid row number so cell references like 'A0' are not valid.
         * Actually, they should be treated like defined names.
         * <br/>
         * In addition, leading zeros (on the row component) should be removed from cell
         * references during parsing.
         */
        [Test]
        public void TestZeroRowRefs()
        {
            String badCellRef = "B0"; // bad because zero is not a valid row number
            String leadingZeroCellRef = "B000001"; // this should get parsed as "B1"
            HSSFWorkbook wb = new HSSFWorkbook();

            try
            {
                HSSFFormulaParser.Parse(badCellRef, wb);
                Assert.Fail("Identified bug 47312b - Shouldn't be able to parse cell ref '"
                        + badCellRef + "'.");
            }
            catch (FormulaParseException e)
            {
                // expected during successful Test
                ConfirmParseException(e, "Specified named range '"
                        + badCellRef + "' does not exist in the current workbook.");
            }

            Ptg[] ptgs;
            try
            {
                ptgs = HSSFFormulaParser.Parse(leadingZeroCellRef, wb);
                Assert.AreEqual("B1", ((RefPtg)ptgs[0]).ToFormulaString());
            }
            catch (FormulaParseException e)
            {
                ConfirmParseException(e, "Specified named range '"
                        + leadingZeroCellRef + "' does not exist in the current workbook.");
                // close but no cigar
                Assert.Fail("Identified bug 47312c - '"
                        + leadingZeroCellRef + "' should parse as 'B1'.");
            }

            // create a defined name called 'B0' and try again
            IName n = wb.CreateName();
            n.NameName = ("B0");
            n.RefersToFormula = ("1+1");
            ptgs = HSSFFormulaParser.Parse("B0", wb);
            ConfirmTokenClasses(ptgs, typeof(NamePtg));
            wb.Close();
        }

        private static void ConfirmParseException(FormulaParseException e, String expMsg)
        {
            Assert.AreEqual(expMsg, e.Message);
        }

        [Test]
        public void Test57196_Formula()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Ptg[] ptgs = HSSFFormulaParser.Parse("DEC2HEX(HEX2DEC(O8)-O2+D2)", wb, FormulaType.Cell, -1);
            Assert.IsNotNull(ptgs, "Ptg array should not be null");

            ConfirmTokenClasses(ptgs,
                typeof(NameXPtg), // ??
                typeof(NameXPtg), // ??
                typeof(RefPtg), // O8
                typeof(FuncVarPtg), // HEX2DEC
                typeof(RefPtg), // O2
                typeof(SubtractPtg),
                typeof(RefPtg),   // D2
                typeof(AddPtg),
                typeof(FuncVarPtg) // DEC2HEX
            );

            RefPtg o8 = (RefPtg)ptgs[2];
            FuncVarPtg hex2Dec = (FuncVarPtg)ptgs[3];
            RefPtg o2 = (RefPtg)ptgs[4];
            RefPtg d2 = (RefPtg)ptgs[6];
            FuncVarPtg dec2Hex = (FuncVarPtg)ptgs[8];

            Assert.AreEqual("O8", o8.ToFormulaString());
            Assert.AreEqual(255, hex2Dec.FunctionIndex);
            //Assert.AreEqual("", hex2Dec.ToString());
            Assert.AreEqual("O2", o2.ToFormulaString());
            Assert.AreEqual("D2", d2.ToFormulaString());
            Assert.AreEqual(255, dec2Hex.FunctionIndex);
            wb.Close();
        }

    }
}
