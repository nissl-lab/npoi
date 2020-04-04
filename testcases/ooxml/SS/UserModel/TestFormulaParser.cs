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
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;

namespace TestCases.SS.UserModel
{
    [TestFixture]
    public class TestFormulaParser
    {
        [Test]
        public void TestHSSFFailsForOver65536()
        {
            IFormulaParsingWorkbook workbook = HSSFEvaluationWorkbook.Create(new HSSFWorkbook());
            try
            {
                FormulaParser.Parse("Sheet1!1:65537", workbook, FormulaType.Cell, 0);
                Assert.Fail("Expected exception");
            }
            catch (FormulaParseException expected)
            {
            }
        }
        [Test]
        public void TestHSSFPassCase()
        {
            IFormulaParsingWorkbook workbook = HSSFEvaluationWorkbook.Create(new HSSFWorkbook());
            FormulaParser.Parse("Sheet1!1:65536", workbook, FormulaType.Cell, 0);
        }
        [Test]
        public void TestXSSFWorksForOver65536()
        {
            IFormulaParsingWorkbook workbook = XSSFEvaluationWorkbook.Create(new XSSFWorkbook());
            FormulaParser.Parse("Sheet1!1:65537", workbook, FormulaType.Cell, 0);
        }
        [Test]
        public void TestXSSFFailCase()
        {
            IFormulaParsingWorkbook workbook = XSSFEvaluationWorkbook.Create(new XSSFWorkbook());
            try
            {
                FormulaParser.Parse("Sheet1!1:1048577", workbook, FormulaType.Cell, 0); // one more than max rows.
                Assert.Fail("Expected exception");
            }
            catch (FormulaParseException)
            {
            }
        }

        [Test]
        // copied from org.apache.poi.hssf.model.TestFormulaParser
        public void TestMacroFunction()
        {
            // testNames.xlsm contains a VB function called 'myFunc'
            String testFile = "testNames.xlsm";
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook(testFile);
            try
            {
                XSSFEvaluationWorkbook workbook = XSSFEvaluationWorkbook.Create(wb);

                //Expected ptg stack: [NamePtg(myFunc), StringPtg(arg), (additional operands would go here...), FunctionPtg(myFunc)]
                Ptg[] ptg = FormulaParser.Parse("myFunc(\"arg\")", workbook, FormulaType.Cell, -1);
                Assert.AreEqual(3, ptg.Length);

                // the name gets encoded as the first operand on the stack
                NameXPxg tname = (NameXPxg)ptg[0];
                Assert.AreEqual("myFunc", tname.ToFormulaString());

                // the function's arguments are pushed onto the stack from left-to-right as OperandPtgs
                StringPtg arg = (StringPtg)ptg[1];
                Assert.AreEqual("arg", arg.Value);

                // The external FunctionPtg is the last Ptg added to the stack
                // During formula evaluation, this Ptg pops off the the appropriate number of
                // arguments (getNumberOfOperands()) and pushes the result on the stack 
                AbstractFunctionPtg tfunc = (AbstractFunctionPtg)ptg[2];
                Assert.IsTrue(tfunc.IsExternalFunction);

                // confirm formula parsing is case-insensitive
                FormulaParser.Parse("mYfUnC(\"arg\")", workbook, FormulaType.Cell, -1);

                // confirm formula parsing doesn't care about argument count or type
                // this should only throw an error when evaluating the formula.
                FormulaParser.Parse("myFunc()", workbook, FormulaType.Cell, -1);
                FormulaParser.Parse("myFunc(\"arg\", 0, TRUE)", workbook, FormulaType.Cell, -1);

                // A completely unknown formula name (not saved in workbook) should still be parseable and renderable
                // but will throw an NotImplementedFunctionException or return a #NAME? error value if evaluated.
                FormulaParser.Parse("yourFunc(\"arg\")", workbook, FormulaType.Cell, -1);

                // Make sure workbook can be written and read
                XSSFTestDataSamples.WriteOutAndReadBack(wb).Close();

                // Manually check to make sure file isn't corrupted
                // TODO: develop a process for occasionally manually reviewing workbooks
                // to verify workbooks are not corrupted
                /*
                FileInfo fileIn = XSSFTestDataSamples.GetSampleFile(testFile);
                FileInfo reSavedFile = new FileInfo(fileIn.FullName.Replace(".xlsm", "-saved.xlsm"));
                FileStream fos = new FileStream(reSavedFile.FullName, FileMode.Create, FileAccess.ReadWrite);
                wb.Write(fos);
                fos.Close();
                */
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestParserErrors()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("testNames.xlsm");
            try {
                XSSFEvaluationWorkbook workbook = XSSFEvaluationWorkbook.Create(wb);

                parseExpectedException("(");
                parseExpectedException(")");
                parseExpectedException("+");
                parseExpectedException("42+");
                parseExpectedException("IF()");
                parseExpectedException("IF("); //no closing paren
                parseExpectedException("myFunc(", workbook); //no closing paren
            } finally {
                wb.Close();
            }
        }

        private static void parseExpectedException(String formula)
        {
            parseExpectedException(formula, null);
        }

        /** confirm formula has invalid syntax and parsing the formula results in FormulaParseException
         * @param formula
         * @param wb
         */
        private static void parseExpectedException(String formula, IFormulaParsingWorkbook wb)
        {
            try
            {
                FormulaParser.Parse(formula, wb, FormulaType.Cell, -1);
                Assert.Fail("Expected FormulaParseException: " + formula);
            }
            catch (FormulaParseException e)
            {
                // expected during successful test
                Assert.IsNotNull(e.Message);
            }
        }
    }
}
