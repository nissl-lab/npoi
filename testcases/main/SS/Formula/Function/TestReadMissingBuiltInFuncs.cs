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

namespace TestCases.SS.Formula.Function
{
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using System;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NPOI.Util;

    /**
     * Tests Reading from a sample spreadsheet some built-in functions that were not properly
     * registered in POI as of bug #44675, #44733 (March/April 2008).
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestReadMissingBuiltInFuncs
    {

        /**
         * This spreadsheet has examples of calls to the interesting built-in functions in cells A1:A7
         */
        private static String SAMPLE_SPREADSHEET_FILE_NAME = "missingFuncs44675.xls";
        private static ISheet _sheet;

        private static ISheet GetSheet()
        {
            if (_sheet == null)
            {
                HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook(SAMPLE_SPREADSHEET_FILE_NAME);
                _sheet = wb.GetSheetAt(0);
            }
            return _sheet;
        }
        [Test]
        public void TestDatedif()
        {

            String formula;
            try
            {
                formula = GetCellFormula(0);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.StartsWith("Too few arguments"))
                {
                    if (e.Message.IndexOf("AttrPtg") > 0)
                    {
                        throw afe("tAttrVolatile not supported in FormulaParser.ToFormulaString");
                    }
                    throw afe("NOW() registered with 1 arg instead of 0");
                }
                if (e.Message.StartsWith("too much stuff"))
                {
                    throw afe("DATEDIF() not registered");
                }
                // some other unexpected error
                throw e;
            }
            Assert.AreEqual("DATEDIF(NOW(),NOW(),\"d\")", formula);
        }
        [Test]
        public void TestDdb()
        {

            String formula = GetCellFormula(1);
            if ("externalflag(1,1,1,1,1)".Equals(formula))
            {
                throw afe("DDB() not registered");
            }
            Assert.AreEqual("DDB(1,1,1,1,1)", formula);
        }
        [Test]
        public void TestAtan()
        {

            String formula = GetCellFormula(2);
            if (formula.Equals("ARCTAN(1)"))
            {
                throw afe("func ix 18 registered as ARCTAN() instead of ATAN()");
            }
            Assert.AreEqual("ATAN(1)", formula);
        }
        [Test]
        public void TestUsdollar()
        {

            String formula = GetCellFormula(3);
            if (formula.Equals("YEN(1)"))
            {
                throw afe("func ix 204 registered as YEN() instead of USDOLLAR()");
            }
            Assert.AreEqual("USDOLLAR(1)", formula);
        }
        [Test]
        public void TestDBCS()
        {

            String formula;
            try
            {
                formula = GetCellFormula(4);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.StartsWith("too much stuff"))
                {
                    throw afe("DBCS() not registered");
                }
                // some other unexpected error
                throw e;
            }
            catch (Exception)
            {
                throw new AssertionException("found err- DBCS() registered with -1 args");
            }
            if (formula.Equals("JIS(\"abc\")"))
            {
                throw afe("func ix 215 registered as JIS() instead of DBCS()");
            }
            Assert.AreEqual("DBCS(\"abc\")", formula);
        }
        [Test]
        public void TestIsnontext()
        {
            String formula;
            try
            {
                formula = GetCellFormula(5);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.StartsWith("too much stuff"))
                {
                    throw afe("ISNONTEXT() registered with wrong index");
                }
                // some other unexpected error
                throw e;
            }
            Assert.AreEqual("ISNONTEXT(\"abc\")", formula);
        }
        [Test]
        public void TestDproduct()
        {

            String formula = GetCellFormula(6);
            Assert.AreEqual("DPRODUCT(C1:E5,\"HarvestYield\",G1:H2)", formula);
        }

        private String GetCellFormula(int rowIx)
        {
            ISheet sheet;
            try
            {
                sheet = GetSheet();
            }
            catch (RecordFormatException e)
            {
                if (e.Message.Equals("Invalid built-in function index (189)"))
                {
                    throw afe("DPRODUCT() registered with wrong index");
                }
                // some other unexpected error
                throw e;
            }
            String result = sheet.GetRow(rowIx).GetCell(0).CellFormula;
#if !HIDE_UNREACHABLE_CODE
            if (false)
            {
                System.Console.Error.WriteLine(result);
            }
#endif
            return result;
        }
        private static AssertionException afe(String msg)
        {
            return new AssertionException(msg);
        }
    }

}