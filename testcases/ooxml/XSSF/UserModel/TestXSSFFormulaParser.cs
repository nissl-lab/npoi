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
using NPOI.Util;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFFormulaParser
    {

        private static Ptg[] Parse(IFormulaParsingWorkbook fpb, String fmla)
        {
            return FormulaParser.Parse(fmla, fpb, FormulaType.Cell, -1, -1);
        }

        private static Ptg[] Parse(IFormulaParsingWorkbook fpb, String fmla, int rowIndex)
        {
            return FormulaParser.Parse(fmla, fpb, FormulaType.Cell, -1, rowIndex);
        }

        [Test]
        public void BasicParse()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;

            ptgs = Parse(fpb, "ABC10");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));

            ptgs = Parse(fpb, "A500000");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));

            ptgs = Parse(fpb, "ABC500000");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));

            //highest allowed rows and column (XFD and 0x100000)
            ptgs = Parse(fpb, "XFD1048576");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));


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

            wb.Close();
        }
        [Test]
        public void TestBuiltInFormulas()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;

            ptgs = Parse(fpb, "LOG10");
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));

            ptgs = Parse(fpb, "LOG10(100)");
            Assert.AreEqual(2, ptgs.Length);
            Assert.IsTrue(ptgs[0] is IntPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is FuncPtg, "Had " + Arrays.ToString(ptgs));

            wb.Close();
        }
        [Test]
        public void FormulaReferencesSameWorkbook()
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

            wb.Close();
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

            wb.Close();
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

            wb.Close();
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

                wb.Close();
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

        [Test]
        public void Test58648Single()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;
            ptgs = Parse(fpb, "(ABC10 )");
            Assert.AreEqual(2, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));

            wb.Close();
        }
        [Test]
        public void Test58648Basic()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;
            // verify whitespaces in different places
            ptgs = Parse(fpb, "(ABC10)");
            Assert.AreEqual(2, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is RefPtg);
            Assert.IsTrue(ptgs[1] is ParenthesisPtg);
            ptgs = Parse(fpb, "( ABC10)");
            Assert.AreEqual(2, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, "(ABC10 )");
            Assert.AreEqual(2, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, "((ABC10))");
            Assert.AreEqual(3, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, "((ABC10) )");
            Assert.AreEqual(3, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, "( (ABC10))");
            Assert.AreEqual(3, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is RefPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is ParenthesisPtg, "Had " + Arrays.ToString(ptgs));

            wb.Close();
        }
        [Test]
        public void Test58648FormulaParsing()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("58648.xlsx");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet xsheet = wb.GetSheetAt(i);
                foreach (IRow row in xsheet)
                {
                    foreach (ICell cell in row)
                    {
                        if (cell.CellType == CellType.Formula)
                        {
                            try
                            {
                                evaluator.EvaluateFormulaCell(cell);
                            }
                            catch (Exception e)
                            {
                                CellReference cellRef = new CellReference(cell.RowIndex, cell.ColumnIndex);
                                throw new RuntimeException("error at: " + cellRef.ToString(), e);
                            }
                        }
                    }
                }
            }
            ISheet sheet = wb.GetSheet("my-sheet");
            ICell cell1 = sheet.GetRow(1).GetCell(4);
            Assert.AreEqual(5d, cell1.NumericCellValue, 0d);

            wb.Close();
        }
        [Test]
        public void TestWhitespaceInFormula()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;
            // verify whitespaces in different places
            ptgs = Parse(fpb, "INTERCEPT(A2:A5, B2:B5)");
            Assert.AreEqual(3, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is FuncPtg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, " INTERCEPT ( \t \r A2 : \nA5 , B2 : B5 ) \t");
            Assert.AreEqual(3, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is FuncPtg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, "(VLOOKUP(\"item1\", A2:B3, 2, FALSE) - VLOOKUP(\"item2\", A2:B3, 2, FALSE) )");
            Assert.AreEqual(12, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is StringPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is IntPtg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, "A1:B1 B1:B2");
            Assert.AreEqual(4, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is MemAreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[3] is IntersectionPtg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, "A1:B1    B1:B2");
            Assert.AreEqual(4, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is MemAreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[3] is IntersectionPtg, "Had " + Arrays.ToString(ptgs));

            wb.Close();
        }
        [Test]
        public void TestWhitespaceInComplexFormula()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;
            // verify whitespaces in different places
            ptgs = Parse(fpb, "SUM(A1:INDEX(1:1048576,MAX(IFERROR(MATCH(99^99,B:B,1),0),IFERROR(MATCH(\"zzzz\",B:B,1),0)),MAX(IFERROR(MATCH(99^99,1:1,1),0),IFERROR(MATCH(\"zzzz\",1:1,1),0))))");
            Assert.AreEqual(40, ptgs.Length, "Had: " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is MemFuncPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is RefPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[3] is NameXPxg, "Had " + Arrays.ToString(ptgs));
            ptgs = Parse(fpb, "SUM ( A1 : INDEX( 1 : 1048576 , MAX( IFERROR ( MATCH ( 99 ^ 99 , B : B , 1 ) , 0 ) , IFERROR ( MATCH ( \"zzzz\" , B:B , 1 ) , 0 ) ) , MAX ( IFERROR ( MATCH ( 99 ^ 99 , 1 : 1 , 1 ) , 0 ) , IFERROR ( MATCH ( \"zzzz\" , 1 : 1 , 1 )   , 0 )   )   )   )");
            Assert.AreEqual(40, ptgs.Length, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[0] is MemFuncPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[1] is RefPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[2] is AreaPtg, "Had " + Arrays.ToString(ptgs));
            Assert.IsTrue(ptgs[3] is NameXPxg, "Had " + Arrays.ToString(ptgs));

            wb.Close();
        }

        [Test]
        public void ParseStructuredReferences()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs;
            /*
            The following cases are tested (copied from FormulaParser.parseStructuredReference)
               1 Table1[col]
               2 Table1[[#Totals],[col]]
               3 Table1[#Totals]
               4 Table1[#All]
               5 Table1[#Data]
               6 Table1[#Headers]
               7 Table1[#Totals]
               8 Table1[#This Row]
               9 Table1[[#All],[col]]
              10 Table1[[#Headers],[col]]
              11 Table1[[#Totals],[col]]
              12 Table1[[#All],[col1]:[col2]]
              13 Table1[[#Data],[col1]:[col2]]
              14 Table1[[#Headers],[col1]:[col2]]
              15 Table1[[#Totals],[col1]:[col2]]
              16 Table1[[#Headers],[#Data],[col2]]
              17 Table1[[#This Row], [col1]]
              18 Table1[ [col1]:[col2] ]
            */
            String tbl = "\\_Prime.1";
            String noTotalsRowReason = ": Tables without a Totals row should return #REF! on [#Totals]";
            ////// Case 1: Evaluate Table1[col] with apostrophe-escaped #-signs ////////
            ptgs = Parse(fpb, "SUM(" + tbl + "[calc='#*'#])");
            Assert.AreEqual(2, ptgs.Length);
            // Area3DPxg [sheet=Table ! A2:A7]
            Assert.IsTrue(ptgs[0] is Area3DPxg);
            Area3DPxg ptg0 = (Area3DPxg)ptgs[0];
            Assert.AreEqual("Table", ptg0.SheetName);
            Assert.AreEqual("A2:A7", ptg0.Format2DRefAsString());
            // Note: structured references are evaluated and resolved to regular 3D area references.
            Assert.AreEqual("Table!A2:A7", ptg0.ToFormulaString());
            // AttrPtg [sum ]
            Assert.IsTrue(ptgs[1] is AttrPtg);
            AttrPtg ptg1 = (AttrPtg)ptgs[1];
            Assert.IsTrue(ptg1.IsSum);

            ////// Case 1: Evaluate "Table1[col]" ////////
            ptgs = Parse(fpb, tbl + "[Name]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!B2:B7", ptgs[0].ToFormulaString(), "Table1[col]");

            ////// Case 2: Evaluate "Table1[[#Totals],[col]]" ////////
            ptgs = Parse(fpb, tbl + "[[#Totals],[col]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(ErrPtg.REF_INVALID, ptgs[0], "Table1[[#Totals],[col]]" + noTotalsRowReason);

            ////// Case 3: Evaluate "Table1[#Totals]" ////////
            ptgs = Parse(fpb, tbl + "[#Totals]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(ErrPtg.REF_INVALID, ptgs[0], "Table1[#Totals]" + noTotalsRowReason);

            ////// Case 4: Evaluate "Table1[#All]" ////////
            ptgs = Parse(fpb, tbl + "[#All]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!A1:C7", ptgs[0].ToFormulaString(), "Table1[#All]");

            ////// Case 5: Evaluate "Table1[#Data]" (excludes Header and Data rows) ////////
            ptgs = Parse(fpb, tbl + "[#Data]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!A2:C7", ptgs[0].ToFormulaString(), "Table1[#Data]");

            ////// Case 6: Evaluate "Table1[#Headers]" ////////
            ptgs = Parse(fpb, tbl + "[#Headers]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!A1:C1", ptgs[0].ToFormulaString(), "Table1[#Headers]");

            ////// Case 7: Evaluate "Table1[#Totals]" ////////
            ptgs = Parse(fpb, tbl + "[#Totals]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(ErrPtg.REF_INVALID, ptgs[0], "Table1[#Totals]" + noTotalsRowReason);

            ////// Case 8: Evaluate "Table1[#This Row]" ////////
            ptgs = Parse(fpb, tbl + "[#This Row]", 2);
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!A3:C3", ptgs[0].ToFormulaString(), "Table1[#This Row]");

            ////// Evaluate "Table1[@]" (equivalent to "Table1[#This Row]") ////////
            ptgs = Parse(fpb, tbl + "[@]", 2);
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!A3:C3", ptgs[0].ToFormulaString());
            
            ////// Evaluate "Table1[#This Row]" when rowIndex is outside Table ////////
            ptgs = Parse(fpb, tbl + "[#This Row]", 10);
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(ErrPtg.VALUE_INVALID, ptgs[0], "Table1[#This Row]");
            
            ////// Evaluate "Table1[@]" when rowIndex is outside Table ////////
            ptgs = Parse(fpb, tbl + "[@]", 10);
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(ErrPtg.VALUE_INVALID, ptgs[0], "Table1[@]");
            ////// Evaluate "Table1[[#Data],[col]]" ////////
            ptgs = Parse(fpb, tbl + "[[#Data], [Number]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!C2:C7", ptgs[0].ToFormulaString(), "Table1[[#Data],[col]]");
            
            ////// Case 9: Evaluate "Table1[[#All],[col]]" ////////
            ptgs = Parse(fpb, tbl + "[[#All], [Number]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!C1:C7", ptgs[0].ToFormulaString(), "Table1[[#All],[col]]");
            
            ////// Case 10: Evaluate "Table1[[#Headers],[col]]" ////////
            ptgs = Parse(fpb, tbl + "[[#Headers], [Number]]");
            Assert.AreEqual(1, ptgs.Length);
            // also acceptable: Table1!B1
            Assert.AreEqual("Table!C1:C1", ptgs[0].ToFormulaString(), "Table1[[#Headers],[col]]");
            ////// Case 11: Evaluate "Table1[[#Totals],[col]]" ////////
            ptgs = Parse(fpb, tbl + "[[#Totals],[Name]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(ErrPtg.REF_INVALID, ptgs[0], "Table1[[#Totals],[col]]" + noTotalsRowReason);
            
            ////// Case 12: Evaluate "Table1[[#All],[col1]:[col2]]" ////////
            ptgs = Parse(fpb, tbl + "[[#All], [Name]:[Number]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!B1:C7", ptgs[0].ToFormulaString(), "Table1[[#All],[col1]:[col2]]");
            
            ////// Case 13: Evaluate "Table1[[#Data],[col]:[col2]]" ////////
            ptgs = Parse(fpb, tbl + "[[#Data], [Name]:[Number]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!B2:C7", ptgs[0].ToFormulaString(), "Table1[[#Data],[col]:[col2]]");
            
            ////// Case 14: Evaluate "Table1[[#Headers],[col1]:[col2]]" ////////
            ptgs = Parse(fpb, tbl + "[[#Headers], [Name]:[Number]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!B1:C1", ptgs[0].ToFormulaString(), "Table1[[#Headers],[col1]:[col2]]");
            
            ////// Case 15: Evaluate "Table1[[#Totals],[col]:[col2]]" ////////
            ptgs = Parse(fpb, tbl + "[[#Totals], [Name]:[Number]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual(ErrPtg.REF_INVALID, ptgs[0], "Table1[[#Totals],[col]:[col2]]" + noTotalsRowReason);
            
            ////// Case 16: Evaluate "Table1[[#Headers],[#Data],[col]]" ////////
            ptgs = Parse(fpb, tbl + "[[#Headers],[#Data],[Number]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!C1:C7", ptgs[0].ToFormulaString(), "Table1[[#Headers],[#Data],[col]]");
            
            ////// Case 17: Evaluate "Table1[[#This Row], [col1]]" ////////
            ptgs = Parse(fpb, tbl + "[[#This Row], [Number]]", 2);
            Assert.AreEqual(1, ptgs.Length);
            // also acceptable: Table!C3
            Assert.AreEqual("Table!C3:C3", ptgs[0].ToFormulaString(), "Table1[[#This Row], [col1]]");
            
            ////// Case 18: Evaluate "Table1[[col]:[col2]]" ////////
            ptgs = Parse(fpb, tbl + "[[Name]:[Number]]");
            Assert.AreEqual(1, ptgs.Length);
            Assert.AreEqual("Table!B2:C7", ptgs[0].ToFormulaString(), "Table1[[col]:[col2]]");
            wb.Close();
        }

    }
}

