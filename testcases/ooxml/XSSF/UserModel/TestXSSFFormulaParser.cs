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
using NPOI.SS.UserModel;
using TestCases.HSSF;
using NPOI.HSSF.UserModel;
namespace NPOI.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFFormulaParser
    {

        private static Ptg[] Parse(IFormulaParsingWorkbook fpb, String fmla)
        {
            return FormulaParser.Parse(fmla, fpb, FormulaType.Cell, -1);
        }

        [Test]
        public void BasicParse()
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

            // Formula referencing one cell
            ptgs = Parse(fpb, "ISEVEN(A1)");
            Assert.AreEqual(3, ptgs.Length);
            Assert.AreEqual(typeof(NameXPxg), ptgs[0].GetType());
            Assert.AreEqual(typeof(RefPtg), ptgs[1].GetType());
            Assert.AreEqual(typeof(FuncVarPtg), ptgs[2].GetType());
            Assert.AreEqual("ISEVEN", ptgs[0].ToFormulaString());
            Assert.AreEqual("A1", ptgs[1].ToFormulaString());
            Assert.AreEqual("#external#", ptgs[2].ToFormulaString());

            // Formula referencing an area
            ptgs = Parse(fpb, "SUM(A1:B3)");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(typeof(AreaPtg), ptgs[0].GetType());
            Assert.AreEqual(typeof(AttrPtg), ptgs[1].GetType());
            Assert.AreEqual("A1:B3", ptgs[0].ToFormulaString());
            Assert.AreEqual("SUM", ptgs[1].ToFormulaString());

            // Formula referencing one cell in a different sheet
            ptgs = Parse(fpb, "SUM(Sheet1!A1)");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(typeof(Ref3DPxg), ptgs[0].GetType());
            Assert.AreEqual(typeof(AttrPtg), ptgs[1].GetType());
            Assert.AreEqual("Sheet1!A1", ptgs[0].ToFormulaString());
            Assert.AreEqual("SUM", ptgs[1].ToFormulaString());

            // Formula referencing an area in a different sheet
            ptgs = Parse(fpb, "SUM(Sheet1!A1:B3)");
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(typeof(Area3DPxg), ptgs[0].GetType());
            Assert.AreEqual(typeof(AttrPtg), ptgs[1].GetType());
            Assert.AreEqual("Sheet1!A1:B3", ptgs[0].ToFormulaString());
            Assert.AreEqual("SUM", ptgs[1].ToFormulaString());

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

        /**
         * A handful of functions (such as SUM, COUNTA, MIN) support
         *  multi-sheet references (eg Sheet1:Sheet3!A1 = Cell A1 from
         *  Sheets 1 through Sheet 3) and multi-sheet area references 
         *  (eg Sheet1:Sheet3!A1:B2 = Cells A1 through B2 from Sheets
         *   1 through Sheet 3).
         * This test, based on common test files for HSSF and XSSF, checks
         *  that we can read and parse these kinds of references 
         * (but not evaluate - that's elsewhere in the test suite)
         */
        [Test]
        public void MultiSheetReferencesHSSFandXSSF()
        {
            IWorkbook[] wbs = new IWorkbook[] {
                HSSFTestDataSamples.OpenSampleWorkbook("55906-MultiSheetRefs.xls"),
                XSSFTestDataSamples.OpenSampleWorkbook("55906-MultiSheetRefs.xlsx")
        };
            foreach (IWorkbook wb in wbs)
            {
                ISheet s1 = wb.GetSheetAt(0);
                Ptg[] ptgs;

                // Check the contents
                ICell sumF = s1.GetRow(2).GetCell(0);
                Assert.IsNotNull(sumF);
                Assert.AreEqual("SUM(Sheet1:Sheet3!A1)", sumF.CellFormula);

                ICell avgF = s1.GetRow(2).GetCell(1);
                Assert.IsNotNull(avgF);
                Assert.AreEqual("AVERAGE(Sheet1:Sheet3!A1)", avgF.CellFormula);

                ICell countAF = s1.GetRow(2).GetCell(2);
                Assert.IsNotNull(countAF);
                Assert.AreEqual("COUNTA(Sheet1:Sheet3!C1)", countAF.CellFormula);

                ICell maxF = s1.GetRow(4).GetCell(1);
                Assert.IsNotNull(maxF);
                Assert.AreEqual("MAX(Sheet1:Sheet3!A$1)", maxF.CellFormula);

                ICell sumFA = s1.GetRow(2).GetCell(7);
                Assert.IsNotNull(sumFA);
                Assert.AreEqual("SUM(Sheet1:Sheet3!A1:B2)", sumFA.CellFormula);

                ICell avgFA = s1.GetRow(2).GetCell(8);
                Assert.IsNotNull(avgFA);
                Assert.AreEqual("AVERAGE(Sheet1:Sheet3!A1:B2)", avgFA.CellFormula);

                ICell maxFA = s1.GetRow(4).GetCell(8);
                Assert.IsNotNull(maxFA);
                Assert.AreEqual("MAX(Sheet1:Sheet3!A$1:B$2)", maxFA.CellFormula);

                ICell countFA = s1.GetRow(5).GetCell(8);
                Assert.IsNotNull(countFA);
                Assert.AreEqual("COUNT(Sheet1:Sheet3!$A$1:$B$2)", countFA.CellFormula);
            

                // Create a formula Parser
                IFormulaParsingWorkbook fpb = null;
                if (wb is HSSFWorkbook)
                    fpb = HSSFEvaluationWorkbook.Create((HSSFWorkbook)wb);
                else
                    fpb = XSSFEvaluationWorkbook.Create((XSSFWorkbook)wb);


                // Check things parse as expected:

                // SUM to one cell over 3 workbooks, relative reference
                ptgs = Parse(fpb, "SUM(Sheet1:Sheet3!A1)");
                Assert.AreEqual(2, ptgs.Length);
                if (wb is HSSFWorkbook)
                {
                    Assert.AreEqual(typeof(Ref3DPtg), ptgs[0].GetType());
                }
                else
                {
                    Assert.AreEqual(typeof(Ref3DPxg), ptgs[0].GetType());
                }
                Assert.AreEqual("Sheet1:Sheet3!A1", ToFormulaString(ptgs[0], fpb));
                Assert.AreEqual(typeof(AttrPtg), ptgs[1].GetType());
                Assert.AreEqual("SUM", ToFormulaString(ptgs[1], fpb));


                // MAX to one cell over 3 workbooks, absolute row reference
                ptgs = Parse(fpb, "MAX(Sheet1:Sheet3!A$1)");
                Assert.AreEqual(2, ptgs.Length);
                if (wb is HSSFWorkbook)
                {
                    Assert.AreEqual(typeof(Ref3DPtg), ptgs[0].GetType());
                }
                else
                {
                    Assert.AreEqual(typeof(Ref3DPxg), ptgs[0].GetType());
                }
                Assert.AreEqual("Sheet1:Sheet3!A$1", ToFormulaString(ptgs[0], fpb));
                Assert.AreEqual(typeof(FuncVarPtg), ptgs[1].GetType());
                Assert.AreEqual( "MAX", ToFormulaString(ptgs[1], fpb));


                // MIN to one cell over 3 workbooks, absolute reference
                ptgs = Parse(fpb, "MIN(Sheet1:Sheet3!$A$1)");
                Assert.AreEqual(2, ptgs.Length);
                if (wb is HSSFWorkbook)
                {
                    Assert.AreEqual(typeof(Ref3DPtg), ptgs[0].GetType());
                }
                else
                {
                    Assert.AreEqual(typeof(Ref3DPxg), ptgs[0].GetType());
                }
                Assert.AreEqual("Sheet1:Sheet3!$A$1", ToFormulaString(ptgs[0], fpb));
                Assert.AreEqual(typeof(FuncVarPtg), ptgs[1].GetType());
                Assert.AreEqual("MIN", ToFormulaString(ptgs[1], fpb));

                // SUM to a range of cells over 3 workbooks
                ptgs = Parse(fpb, "SUM(Sheet1:Sheet3!A1:B2)");
                Assert.AreEqual(2, ptgs.Length);
                if (wb is HSSFWorkbook)
                {
                    Assert.AreEqual(typeof(Area3DPtg), ptgs[0].GetType());
                }
                else
                {
                    Assert.AreEqual(typeof(Area3DPxg), ptgs[0].GetType());
                }
                Assert.AreEqual(ToFormulaString(ptgs[0], fpb), "Sheet1:Sheet3!A1:B2");
                Assert.AreEqual(typeof(AttrPtg), ptgs[1].GetType());
                Assert.AreEqual("SUM", ToFormulaString(ptgs[1], fpb));


                // MIN to a range of cells over 3 workbooks, absolute reference
                ptgs = Parse(fpb, "MIN(Sheet1:Sheet3!$A$1:$B$2)");
                Assert.AreEqual(2, ptgs.Length);
                if (wb is HSSFWorkbook)
                {
                    Assert.AreEqual(typeof(Area3DPtg), ptgs[0].GetType());
                }
                else
                {
                    Assert.AreEqual(typeof(Area3DPxg), ptgs[0].GetType());
                }
                Assert.AreEqual(ToFormulaString(ptgs[0], fpb), "Sheet1:Sheet3!$A$1:$B$2");
                Assert.AreEqual(typeof(FuncVarPtg), ptgs[1].GetType());
                Assert.AreEqual("MIN", ToFormulaString(ptgs[1], fpb));


                // Check we can round-trip - try to Set a new one to a new single cell
                ICell newF = s1.GetRow(0).CreateCell(10, CellType.Formula);
                newF.CellFormula = (/*setter*/"SUM(Sheet2:Sheet3!A1)");
                Assert.AreEqual("SUM(Sheet2:Sheet3!A1)", newF.CellFormula);

                // Check we can round-trip - try to Set a new one to a cell range
                newF = s1.GetRow(0).CreateCell(11, CellType.Formula);
                newF.CellFormula = (/*setter*/"MIN(Sheet1:Sheet2!A1:B2)");
                Assert.AreEqual("MIN(Sheet1:Sheet2!A1:B2)", newF.CellFormula);

            }
        }
        private static String ToFormulaString(Ptg ptg, IFormulaParsingWorkbook wb)
        {
            if (ptg is WorkbookDependentFormula)
            {
                return ((WorkbookDependentFormula)ptg).ToFormulaString((IFormulaRenderingWorkbook)wb);
            }
            return ptg.ToFormulaString();
        }


    }
}

