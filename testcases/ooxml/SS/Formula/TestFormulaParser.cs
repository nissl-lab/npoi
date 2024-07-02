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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.Formula
{


    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;

    /// <summary>
    /// Test <see cref="FormulaParser"/>'s handling of row numbers at the edge of the
    /// HSSF/XSSF ranges.
    /// </summary>
    /// @author David North
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
            catch(FormulaParseException)
            {
                // expected here
            }
        }

        private static void CheckHSSFFormula(String formula)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            IFormulaParsingWorkbook workbook = HSSFEvaluationWorkbook.Create(wb);
            FormulaParser.Parse(formula, workbook, FormulaType.Cell, 0);
            IOUtils.CloseQuietly(wb);
        }
        private static void CheckXSSFFormula(String formula)
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            IFormulaParsingWorkbook workbook = XSSFEvaluationWorkbook.Create(wb);
            FormulaParser.Parse(formula, workbook, FormulaType.Cell, 0);
            IOUtils.CloseQuietly(wb);
        }
        private static void checkFormula(String formula)
        {
            CheckHSSFFormula(formula);
            CheckXSSFFormula(formula);
        }

        [Test]
        public void TestHSSFPassCase()
        {
            CheckHSSFFormula("Sheet1!1:65536");
        }

        [Test]
        public void TestXSSFWorksForOver65536()
        {
            CheckXSSFFormula("Sheet1!1:65537");
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
            catch(FormulaParseException)
            {
                // expected here
            }
        }

        // copied from NPOI.HSSF.Model.TestFormulaParser
        [Test]
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

                // the name Gets encoded as the first operand on the stack
                NameXPxg tname = (NameXPxg) ptg[0];
                Assert.AreEqual("myFunc", tname.ToFormulaString());

                // the function's arguments are pushed onto the stack from left-to-right as OperandPtgs
                StringPtg arg = (StringPtg) ptg[1];
                Assert.AreEqual("arg", arg.Value);

                // The external FunctionPtg is the last Ptg added to the stack
                // During formula evaluation, this Ptg pops off the the appropriate number of
                // arguments (getNumberOfOperands()) and pushes the result on the stack 
                AbstractFunctionPtg tfunc = (AbstractFunctionPtg) ptg[2];
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
                final File fileIn = XSSFTestDataSamples.GetSampleFile(testFile);
                final File reSavedFile = new File(fileIn.ParentFile, fileIn.Name.replace(".xlsm", "-saved.xlsm"));
                final FileOutputStream fos = new FileOutputStream(reSavedFile);
                wb.write(fos);
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
            try
            {
                XSSFEvaluationWorkbook workbook = XSSFEvaluationWorkbook.Create(wb);

                ParseExpectedException("(");
                ParseExpectedException(")");
                ParseExpectedException("+");
                ParseExpectedException("42+");
                ParseExpectedException("IF()");
                ParseExpectedException("IF("); //no closing paren
                ParseExpectedException("myFunc(", workbook); //no closing paren
            }
            finally
            {
                wb.Close();
            }
        }

        private static void ParseExpectedException(String formula)
        {
            ParseExpectedException(formula, null);
        }

        /// <summary>
        /// confirm formula has invalid syntax and parsing the formula results in FormulaParseException
        /// </summary>
        private static void ParseExpectedException(String formula, IFormulaParsingWorkbook wb)
        {
            try
            {
                FormulaParser.Parse(formula, wb, FormulaType.Cell, -1);
                Assert.Fail("Expected FormulaParseException: " + formula);
            }
            catch(FormulaParseException e)
            {
                // expected during successful test
                Assert.IsNotNull(e.Message);
            }
        }

        // trivial case for bug 60219: FormulaParser can't parse external references when sheet name is quoted
        [Test]
        public void TestParseExternalReferencesWithUnquotedSheetName()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpwb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs = FormulaParser.Parse("[1]Sheet1!A1", fpwb, FormulaType.Cell, -1);
            // NPOI.SS.Formula.PTG.Ref3DPxg [ [workbook=1] sheet=Sheet 1 ! A1]
            Assert.AreEqual(1, ptgs.Length, "Ptgs length");
            Assert.IsTrue(ptgs[0] is Ref3DPxg, "Ptg class");
            Ref3DPxg pxg = (Ref3DPxg) ptgs[0];
            Assert.AreEqual(1, pxg.ExternalWorkbookNumber, "External workbook number");
            Assert.AreEqual("Sheet1", pxg.SheetName, "Sheet name");
            Assert.AreEqual(0, pxg.Row, "Row");
            Assert.AreEqual(0, pxg.Column, "Column");
            wb.Close();
        }

        // bug 60219: FormulaParser can't parse external references when sheet name is quoted
        [Test]
        public void TestParseExternalReferencesWithQuotedSheetName()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpwb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs = FormulaParser.Parse("'[1]Sheet 1'!A1", fpwb, FormulaType.Cell, -1);
            // NPOI.SS.Formula.PTG.Ref3DPxg [ [workbook=1] sheet=Sheet 1 ! A1]
            Assert.AreEqual(1, ptgs.Length, "Ptgs length");
            Assert.IsTrue(ptgs[0] is Ref3DPxg, "Ptg class");
            Ref3DPxg pxg = (Ref3DPxg) ptgs[0];
            Assert.AreEqual(1, pxg.ExternalWorkbookNumber, "External workbook number");
            Assert.AreEqual("Sheet 1", pxg.SheetName, "Sheet name");
            Assert.AreEqual(0, pxg.Row, "Row");
            Assert.AreEqual(0, pxg.Column, "Column");
            wb.Close();
        }

        // bug 60260
        [Test]
        public void TestUnicodeSheetName()
        {
            checkFormula("'Sheet\u30FB1'!A1:A6");
        }
    }
}

