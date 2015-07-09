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

using NUnit.Framework;
using System;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
namespace NPOI.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFFormulaParser
    {

        private static Ptg[] Parse(XSSFEvaluationWorkbook fpb, String fmla)
        {
            return FormulaParser.Parse(fmla, fpb, FormulaType.Cell, -1);
        }

        [Test]
        public void TestParse()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;

            ptgs = Parse(fpb, "ABC10");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg);

            ptgs = Parse(fpb, "A500000");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg);

            ptgs = Parse(fpb, "ABC500000");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg);

            //highest allowed rows and column (XFD and 0x100000)
            ptgs = Parse(fpb, "XFD1048576");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg);


            //column greater than XFD
            try
            {
                ptgs = Parse(fpb, "XFE10");
                Assert.Fail("expected exception");
            }
            catch (FormulaParseException e)
            {
                Assert.AreEqual("Specified named range 'XFE10' does not exist in the current workbook.", e.Message);
            }

            //row greater than 0x100000
            try
            {
                ptgs = Parse(fpb, "XFD1048577");
                Assert.Fail("expected exception");
            }
            catch (FormulaParseException e)
            {
                Assert.AreEqual("Specified named range 'XFD1048577' does not exist in the current workbook.", e.Message);
            }
        }
        [Test]
        public void TestBuiltInFormulas()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;

            ptgs = Parse(fpb, "LOG10");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue((ptgs[0] is RefPtg));

            ptgs = Parse(fpb, "LOG10(100)");
            Assert.AreEqual(2, ptgs.Length);
            Assert.IsTrue(ptgs[0] is IntPtg);
            Assert.IsTrue(ptgs[1] is FuncPtg);
        }
        [Test]
        public void FormaulReferncesSameWorkbook()
        {
            // Use a test file with "other workbook" style references
            //  to itself
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56737.xlsx");
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;

            // Reference to a named range in our own workbook, as if it
            // were defined in a different workbook
            ptgs = Parse(fpb, "[0]!NR_Global_B2");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(typeof(NameXPxg), ptgs[0].GetType());
            Assert.AreEqual(0, ((NameXPxg)ptgs[0]).ExternalWorkbookNumber);
            Assert.AreEqual(null, ((NameXPxg)ptgs[0]).SheetName);
            Assert.AreEqual("NR_Global_B2", ((NameXPxg)ptgs[0]).NameName);
            Assert.AreEqual("[0]!NR_Global_B2", ((NameXPxg)ptgs[0]).ToFormulaString());
        }

        [Test]
        public void FormulaReferencesOtherSheets()
        {
            // Use a test file with the named ranges in place
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56737.xlsx");
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;

            // Reference to a single cell in a different sheet
            ptgs = Parse(fpb, "Uses!A1");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(typeof(Ref3DPxg), ptgs[0].GetType());
            Assert.AreEqual(-1, ((Ref3DPxg)ptgs[0]).ExternalWorkbookNumber);
            Assert.AreEqual("A1", ((Ref3DPxg)ptgs[0]).Format2DRefAsString());
            Assert.AreEqual("Uses!A1", ((Ref3DPxg)ptgs[0]).ToFormulaString());

            // Reference to a single cell in a different sheet, which needs quoting
            ptgs = Parse(fpb, "'Testing 47100'!A1");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(typeof(Ref3DPxg), ptgs[0].GetType());
            Assert.AreEqual(-1, ((Ref3DPxg)ptgs[0]).ExternalWorkbookNumber);
            Assert.AreEqual("Testing 47100", ((Ref3DPxg)ptgs[0]).SheetName);
            Assert.AreEqual("A1", ((Ref3DPxg)ptgs[0]).Format2DRefAsString());
            Assert.AreEqual("'Testing 47100'!A1", ((Ref3DPxg)ptgs[0]).ToFormulaString());
        

            // Reference to a sheet scoped named range from another sheet
            ptgs = Parse(fpb, "Defines!NR_To_A1");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(typeof(NameXPxg), ptgs[0].GetType());
            Assert.AreEqual(-1, ((NameXPxg)ptgs[0]).ExternalWorkbookNumber);
            Assert.AreEqual("Defines", ((NameXPxg)ptgs[0]).SheetName);
            Assert.AreEqual("NR_To_A1", ((NameXPxg)ptgs[0]).NameName);
            Assert.AreEqual("Defines!NR_To_A1", ((NameXPxg)ptgs[0]).ToFormulaString());

            // Reference to a workbook scoped named range
            ptgs = Parse(fpb, "NR_Global_B2");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(typeof(NamePtg), ptgs[0].GetType());
            Assert.AreEqual("NR_Global_B2", ((NamePtg)ptgs[0]).ToFormulaString(fpb));

        }

        [Test]
        public void FormulaReferencesOtherWorkbook()
        {
            // Use a test file with the external linked table in place
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ref-56737.xlsx");
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;

            // Reference to a single cell in a different workbook
            ptgs = Parse(fpb, "[1]Uses!$A$1");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(typeof(Ref3DPxg), ptgs[0].GetType());
            Assert.AreEqual(1, ((Ref3DPxg)ptgs[0]).ExternalWorkbookNumber);
            Assert.AreEqual("Uses", ((Ref3DPxg)ptgs[0]).SheetName);
            Assert.AreEqual("$A$1", ((Ref3DPxg)ptgs[0]).Format2DRefAsString());
            Assert.AreEqual("[1]Uses!$A$1", ((Ref3DPxg)ptgs[0]).ToFormulaString());

            // Reference to a sheet-scoped named range in a different workbook
            ptgs = Parse(fpb, "[1]Defines!NR_To_A1");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(typeof(NameXPxg), ptgs[0].GetType());
            Assert.AreEqual(1, ((NameXPxg)ptgs[0]).ExternalWorkbookNumber);
            Assert.AreEqual("Defines", ((NameXPxg)ptgs[0]).SheetName);
            Assert.AreEqual("NR_To_A1", ((NameXPxg)ptgs[0]).NameName);
            Assert.AreEqual("[1]Defines!NR_To_A1", ((NameXPxg)ptgs[0]).ToFormulaString());

            // Reference to a global named range in a different workbook
            ptgs = Parse(fpb, "[1]!NR_Global_B2");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(typeof(NameXPxg), ptgs[0].GetType());
            Assert.AreEqual(1, ((NameXPxg)ptgs[0]).ExternalWorkbookNumber);
            Assert.AreEqual(null, ((NameXPxg)ptgs[0]).SheetName);
            Assert.AreEqual("NR_Global_B2", ((NameXPxg)ptgs[0]).NameName);
            Assert.AreEqual("[1]!NR_Global_B2", ((NameXPxg)ptgs[0]).ToFormulaString());

        }

    }
}

