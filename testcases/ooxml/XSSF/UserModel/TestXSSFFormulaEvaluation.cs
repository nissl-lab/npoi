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

using TestCases.SS.UserModel;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using TestCases.HSSF;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;

namespace TestCases.XSSF.UserModel
{

    [TestFixture]
    public class TestXSSFFormulaEvaluation : BaseTestFormulaEvaluator
    {

        public TestXSSFFormulaEvaluation()
            : base(XSSFITestDataProvider.instance)
        {

        }

        [Test]
        public void TestSharedFormulas()
        {
            BaseTestSharedFormulas("shared_formulas.xlsx");
        }
        [Test]
        public void TestSharedFormulas_EvaluateInCell()
        {
            XSSFWorkbook wb = (XSSFWorkbook)_testDataProvider.OpenSampleWorkbook("49872.xlsx");
            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            ISheet sheet = wb.GetSheetAt(0);

            double result = 3.0;

            // B3 is a master shared formula, C3 and D3 don't have the formula written in their f element.
            // Instead, the attribute si for a particular cell is used to figure what the formula expression
            // should be based on the cell's relative location to the master formula, e.g.
            // B3:        <f t="shared" ref="B3:D3" si="0">B1+B2</f>
            // C3 and D3: <f t="shared" si="0"/>

            // Get B3 and Evaluate it in the cell
            ICell b3 = sheet.GetRow(2).GetCell(1);
            Assert.AreEqual(result, Evaluator.EvaluateInCell(b3).NumericCellValue, 0);

            //at this point the master formula is gone, but we are still able to Evaluate dependent cells
            ICell c3 = sheet.GetRow(2).GetCell(2);
            Assert.AreEqual(result, Evaluator.EvaluateInCell(c3).NumericCellValue, 0);

            ICell d3 = sheet.GetRow(2).GetCell(3);
            Assert.AreEqual(result, Evaluator.EvaluateInCell(d3).NumericCellValue, 0);

            wb.Close();
        }

        /**
         * Evaluation of cell references with column indexes greater than 255. See bugzilla 50096
         */
        [Test]
        public void TestEvaluateColumnGreaterThan255()
        {
            XSSFWorkbook wb = (XSSFWorkbook)_testDataProvider.OpenSampleWorkbook("50096.xlsx");
            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            /**
             *  The first row simply Contains the numbers 1 - 300.
             *  The second row simply refers to the cell value above in the first row by a simple formula.
             */
            for (int i = 245; i < 265; i++)
            {
                ICell cell_noformula = wb.GetSheetAt(0).GetRow(0).GetCell(i);
                ICell cell_formula = wb.GetSheetAt(0).GetRow(1).GetCell(i);

                CellReference ref_noformula = new CellReference(cell_noformula.RowIndex, cell_noformula.ColumnIndex);
                CellReference ref_formula = new CellReference(cell_noformula.RowIndex, cell_noformula.ColumnIndex);
                String fmla = cell_formula.CellFormula;
                // assure that the formula refers to the cell above.
                // the check below is 'deep' and involves conversion of the shared formula:
                // in the sample file a shared formula in GN1 is spanned in the range GN2:IY2,
                Assert.AreEqual(ref_noformula.FormatAsString(), fmla);

                CellValue cv_noformula = Evaluator.Evaluate(cell_noformula);
                CellValue cv_formula = Evaluator.Evaluate(cell_formula);
                Assert.AreEqual(cv_noformula.NumberValue, cv_formula.NumberValue, 0, "Wrong Evaluation result in " + ref_formula.FormatAsString());
            }

            wb.Close();
        }

        /**
     * Related to bugs #56737 and #56752 - XSSF workbooks which have
     *  formulas that refer to cells and named ranges in multiple other
     *  workbooks, both HSSF and XSSF ones
     */
        [Test]
        public void TestReferencesToOtherWorkbooks()
        {
            XSSFWorkbook wb = (XSSFWorkbook)_testDataProvider.OpenSampleWorkbook("ref2-56737.xlsx");
            XSSFFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;
            XSSFSheet s = wb.GetSheetAt(0) as XSSFSheet;

            // References to a .xlsx file
            IRow rXSLX = s.GetRow(2);
            ICell cXSLX_cell = rXSLX.GetCell(4);
            ICell cXSLX_sNR = rXSLX.GetCell(6);
            ICell cXSLX_gNR = rXSLX.GetCell(8);
            Assert.AreEqual("[1]Uses!$A$1", cXSLX_cell.CellFormula);
            Assert.AreEqual("[1]Defines!NR_To_A1", cXSLX_sNR.CellFormula);
            Assert.AreEqual("[1]!NR_Global_B2", cXSLX_gNR.CellFormula);

            Assert.AreEqual("Hello!", cXSLX_cell.StringCellValue);
            Assert.AreEqual("Test A1", cXSLX_sNR.StringCellValue);
            Assert.AreEqual(142.0, cXSLX_gNR.NumericCellValue, 0);

            // References to a .xls file
            IRow rXSL = s.GetRow(4);
            ICell cXSL_cell = rXSL.GetCell(4);
            ICell cXSL_sNR = rXSL.GetCell(6);
            ICell cXSL_gNR = rXSL.GetCell(8);
            Assert.AreEqual("[2]Uses!$C$1", cXSL_cell.CellFormula);
            Assert.AreEqual("[2]Defines!NR_To_A1", cXSL_sNR.CellFormula);
            Assert.AreEqual("[2]!NR_Global_B2", cXSL_gNR.CellFormula);

            Assert.AreEqual("Hello!", cXSL_cell.StringCellValue);
            Assert.AreEqual("Test A1", cXSL_sNR.StringCellValue);
            Assert.AreEqual(142.0, cXSL_gNR.NumericCellValue, 0);

            // Try to Evaluate without references, won't work
            // (At least, not unit we fix bug #56752 that is1)
            try
            {
                evaluator.Evaluate(cXSL_cell);
                Assert.Fail("Without a fix for #56752, shouldn't be able to Evaluate a " +
                     "reference to a non-provided linked workbook");
            }
            catch (Exception)
            {
            }

            // Setup the environment
            Dictionary<String, IFormulaEvaluator> evaluators = new Dictionary<String, IFormulaEvaluator>();
            evaluators.Add("ref2-56737.xlsx", evaluator);
            evaluators.Add("56737.xlsx",
                    _testDataProvider.OpenSampleWorkbook("56737.xlsx").GetCreationHelper().CreateFormulaEvaluator());
            evaluators.Add("56737.xls",
                    HSSFTestDataSamples.OpenSampleWorkbook("56737.xls").GetCreationHelper().CreateFormulaEvaluator());
            evaluator.SetupReferencedWorkbooks(evaluators);

            // Try Evaluating all of them, ensure we don't blow up
            foreach (IRow r in s)
            {
                foreach (ICell c in r)
                {
                    // TODO Fix and enable
                    evaluator.Evaluate(c);
                }
            }

            // And evaluate the other way too
            evaluator.EvaluateAll();

            // Static evaluator won't work, as no references passed in
            try
            {
                XSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
                Assert.Fail("Static method lacks references, shouldn't work");
            }
            catch (Exception)
            {
                // expected here
            }


            // Evaluate specific cells and check results
            Assert.AreEqual("\"Hello!\"", evaluator.Evaluate(cXSLX_cell).FormatAsString());
            Assert.AreEqual("\"Test A1\"", evaluator.Evaluate(cXSLX_sNR).FormatAsString());
            //Assert.AreEqual("142.0", evaluator.Evaluate(cXSLX_gNR).FormatAsString());
            Assert.AreEqual("142", evaluator.Evaluate(cXSLX_gNR).FormatAsString());

            Assert.AreEqual("\"Hello!\"", evaluator.Evaluate(cXSL_cell).FormatAsString());
            Assert.AreEqual("\"Test A1\"", evaluator.Evaluate(cXSL_sNR).FormatAsString());
            //Assert.AreEqual("142.0", evaluator.Evaluate(cXSL_gNR).FormatAsString());
            Assert.AreEqual("142", evaluator.Evaluate(cXSL_gNR).FormatAsString());

            // Add another formula referencing these workbooks
            ICell cXSL_cell2 = rXSL.CreateCell(40);
            cXSL_cell2.CellFormula = (/*setter*/"[56737.xls]Uses!$C$1");
            // TODO Shouldn't it become [2] like the others?
            Assert.AreEqual("[56737.xls]Uses!$C$1", cXSL_cell2.CellFormula);
            Assert.AreEqual("\"Hello!\"", evaluator.Evaluate(cXSL_cell2).FormatAsString());


            // Now add a formula that refers to yet another (different) workbook
            // Won't work without the workbook being linked
            ICell cXSLX_nw_cell = rXSLX.CreateCell(42);
            try
            {
                cXSLX_nw_cell.CellFormula = (/*setter*/"[alt.xlsx]Sheet1!$A$1");
                Assert.Fail("New workbook not linked, shouldn't be able to Add");
            }
            catch (Exception) { }

            // Link and re-try
            IWorkbook alt = new XSSFWorkbook();
            try
            {
                alt.CreateSheet().CreateRow(0).CreateCell(0).SetCellValue("In another workbook");
                // TODO Implement the rest of this, see bug #57184
                /*
                            wb.linkExternalWorkbook("alt.xlsx", alt);

                            cXSLX_nw_cell.setCellFormula("[alt.xlsx]Sheet1!$A$1");
                            // Check it - TODO Is this correct? Or should it become [3]Sheet1!$A$1 ?
                            Assert.AreEqual("[alt.xlsx]Sheet1!$A$1", cXSLX_nw_cell.getCellFormula());

                            // Evaluate it, without a link to that workbook
                            try {
                                evaluator.evaluate(cXSLX_nw_cell);
                                fail("No cached value and no link to workbook, shouldn't evaluate");
                            } catch(Exception e) {}

                            // Add a link, check it does
                            evaluators.put("alt.xlsx", alt.getCreationHelper().createFormulaEvaluator());
                            evaluator.setupReferencedWorkbooks(evaluators);

                            evaluator.evaluate(cXSLX_nw_cell);
                            Assert.AreEqual("In another workbook", cXSLX_nw_cell.getStringCellValue());
                */
            }
            finally
            {
                alt.Close();
            }
            wb.Close();
        }

        /**
         * If a formula references cells or named ranges in another workbook,
         *  but that isn't available at Evaluation time, the cached values
         *  should be used instead
         * TODO Add the support then add a unit test
         * See bug #56752
         */
        [Test]
        public void TestCachedReferencesToOtherWorkbooks()
        {
            // TODO
        }
        /**
         * A handful of functions (such as SUM, COUNTA, MIN) support
         *  multi-sheet references (eg Sheet1:Sheet3!A1 = Cell A1 from
         *  Sheets 1 through Sheet 3).
         * This test, based on common test files for HSSF and XSSF, Checks
         *  that we can correctly Evaluate these
         */
        [Test]
        public void TestMultiSheetReferencesHSSFandXSSF()
        {
            IWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("55906-MultiSheetRefs.xls");
            IWorkbook wb2 = XSSFTestDataSamples.OpenSampleWorkbook("55906-MultiSheetRefs.xlsx");

            foreach (IWorkbook wb in new IWorkbook[] { wb1, wb2 })
            {
                IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
                ISheet s1 = wb.GetSheetAt(0);


                // Simple SUM over numbers
                ICell sumF = s1.GetRow(2).GetCell(0);
                Assert.IsNotNull(sumF);
                Assert.AreEqual("SUM(Sheet1:Sheet3!A1)", sumF.CellFormula);
                Assert.AreEqual("66", Evaluator.Evaluate(sumF).FormatAsString(), "Failed for " + wb.GetType());


                // Various Stats formulas on numbers
                ICell avgF = s1.GetRow(2).GetCell(1);
                Assert.IsNotNull(avgF);
                Assert.AreEqual("AVERAGE(Sheet1:Sheet3!A1)", avgF.CellFormula);
                Assert.AreEqual("22", Evaluator.Evaluate(avgF).FormatAsString());

                ICell minF = s1.GetRow(3).GetCell(1);
                Assert.IsNotNull(minF);
                Assert.AreEqual("MIN(Sheet1:Sheet3!A$1)", minF.CellFormula);
                Assert.AreEqual("11", Evaluator.Evaluate(minF).FormatAsString());

                ICell maxF = s1.GetRow(4).GetCell(1);
                Assert.IsNotNull(maxF);
                Assert.AreEqual("MAX(Sheet1:Sheet3!A$1)", maxF.CellFormula);
                Assert.AreEqual("33", Evaluator.Evaluate(maxF).FormatAsString());

                ICell countF = s1.GetRow(5).GetCell(1);
                Assert.IsNotNull(countF);
                Assert.AreEqual("COUNT(Sheet1:Sheet3!A$1)", countF.CellFormula);
                Assert.AreEqual("3", Evaluator.Evaluate(countF).FormatAsString());


                // Various CountAs on Strings
                ICell countA_1F = s1.GetRow(2).GetCell(2);
                Assert.IsNotNull(countA_1F);
                Assert.AreEqual("COUNTA(Sheet1:Sheet3!C1)", countA_1F.CellFormula);
                Assert.AreEqual("3", Evaluator.Evaluate(countA_1F).FormatAsString());

                ICell countA_2F = s1.GetRow(2).GetCell(3);
                Assert.IsNotNull(countA_2F);
                Assert.AreEqual("COUNTA(Sheet1:Sheet3!D1)", countA_2F.CellFormula);
                Assert.AreEqual("0", Evaluator.Evaluate(countA_2F).FormatAsString());

                ICell countA_3F = s1.GetRow(2).GetCell(4);
                Assert.IsNotNull(countA_3F);
                Assert.AreEqual("COUNTA(Sheet1:Sheet3!E1)", countA_3F.CellFormula);
                Assert.AreEqual("3", Evaluator.Evaluate(countA_3F).FormatAsString());
            }

            wb2.Close();
            wb1.Close();
        }
        /**
         * A handful of functions (such as SUM, COUNTA, MIN) support
         *  multi-sheet areas (eg Sheet1:Sheet3!A1:B2 = Cell A1 to Cell B2,
         *  from Sheets 1 through Sheet 3).
         * This test, based on common test files for HSSF and XSSF, checks
         *  that we can correctly evaluate these
         */
        [Test]
        public void TestMultiSheetAreasHSSFandXSSF()
        {
            IWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("55906-MultiSheetRefs.xls");
            IWorkbook wb2 = XSSFTestDataSamples.OpenSampleWorkbook("55906-MultiSheetRefs.xlsx");

            foreach (IWorkbook wb in new IWorkbook[] { wb1, wb2 })
            {
                IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
                ISheet s1 = wb.GetSheetAt(0);


                // SUM over a range
                ICell sumFA = s1.GetRow(2).GetCell(7);
                Assert.IsNotNull(sumFA);
                Assert.AreEqual("SUM(Sheet1:Sheet3!A1:B2)", sumFA.CellFormula);
                Assert.AreEqual("110", Evaluator.Evaluate(sumFA).FormatAsString(), "Failed for " + wb.GetType());


                // Various Stats formulas on ranges of numbers
                ICell avgFA = s1.GetRow(2).GetCell(8);
                Assert.IsNotNull(avgFA);
                Assert.AreEqual("AVERAGE(Sheet1:Sheet3!A1:B2)", avgFA.CellFormula);
                Assert.AreEqual("27.5", Evaluator.Evaluate(avgFA).FormatAsString());

                ICell minFA = s1.GetRow(3).GetCell(8);
                Assert.IsNotNull(minFA);
                Assert.AreEqual("MIN(Sheet1:Sheet3!A$1:B$2)", minFA.CellFormula);
                Assert.AreEqual("11", Evaluator.Evaluate(minFA).FormatAsString());

                ICell maxFA = s1.GetRow(4).GetCell(8);
                Assert.IsNotNull(maxFA);
                Assert.AreEqual("MAX(Sheet1:Sheet3!A$1:B$2)", maxFA.CellFormula);
                Assert.AreEqual("44", Evaluator.Evaluate(maxFA).FormatAsString());

                ICell countFA = s1.GetRow(5).GetCell(8);
                Assert.IsNotNull(countFA);
                Assert.AreEqual("COUNT(Sheet1:Sheet3!$A$1:$B$2)", countFA.CellFormula);
                Assert.AreEqual("4", Evaluator.Evaluate(countFA).FormatAsString());

            }

            wb2.Close();
            wb1.Close();
        }

        [Test]
        public void TestMultisheetFormulaEval()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet1 = wb.CreateSheet("Sheet1") as XSSFSheet;
                XSSFSheet sheet2 = wb.CreateSheet("Sheet2") as XSSFSheet;
                XSSFSheet sheet3 = wb.CreateSheet("Sheet3") as XSSFSheet;

                // sheet1 A1
                XSSFCell cell = sheet1.CreateRow(0).CreateCell(0) as XSSFCell;
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(1.0);

                // sheet2 A1
                cell = sheet2.CreateRow(0).CreateCell(0) as XSSFCell;
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(1.0);

                // sheet2 B1
                cell = sheet2.GetRow(0).CreateCell(1) as XSSFCell;
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(1.0);

                // sheet3 A1
                cell = sheet3.CreateRow(0).CreateCell(0) as XSSFCell;
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(1.0);

                // sheet1 A2 formulae
                cell = sheet1.CreateRow(1).CreateCell(0) as XSSFCell;
                cell.SetCellType(CellType.Formula);
                cell.CellFormula = (/*setter*/"SUM(Sheet1:Sheet3!A1)");

                // sheet1 A3 formulae
                cell = sheet1.CreateRow(2).CreateCell(0) as XSSFCell;
                cell.SetCellType(CellType.Formula);
                cell.CellFormula = (/*setter*/"SUM(Sheet1:Sheet3!A1:B1)");

                wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

                cell = sheet1.GetRow(1).GetCell(0) as XSSFCell;
                Assert.AreEqual(3.0, cell.NumericCellValue, 0);

                cell = sheet1.GetRow(2).GetCell(0) as XSSFCell;
                Assert.AreEqual(4.0, cell.NumericCellValue, 0);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet("test") as XSSFSheet;
                XSSFRow row = sheet.CreateRow(0) as XSSFRow;
                XSSFRow row2 = sheet.CreateRow(1) as XSSFRow;
                XSSFCell cellA2 = row2.CreateCell(0, CellType.Formula) as XSSFCell;
                XSSFCell cellB1 = row.CreateCell(1, CellType.Numeric) as XSSFCell;
                cellB1.SetCellValue(10);
                XSSFFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;
                cellA2.SetCellFormula("IF(B1=0,\"\",((ROW()-ROW(A$1))*12))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                Assert.AreEqual("12", Evaluate.FormatAsString());

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW()-ROW(A$1))*12),\"\")");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);

                Assert.AreEqual(Evaluate.ToString(), EvaluateN.ToString());
                Assert.AreEqual("12", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843a()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet("test") as XSSFSheet;
                XSSFRow row = sheet.CreateRow(0) as XSSFRow;
                XSSFRow row2 = sheet.CreateRow(1) as XSSFRow;
                XSSFCell cellA2 = row2.CreateCell(0, CellType.Formula) as XSSFCell;
                XSSFCell cellB1 = row.CreateCell(1, CellType.Numeric) as XSSFCell;
                cellB1.SetCellValue(10);
                XSSFFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;
                cellA2.SetCellFormula("IF(B1=0,\"\",((ROW(A$1))))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                Assert.AreEqual("1", Evaluate.FormatAsString());

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW(A$1))),\"\")");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);

                Assert.AreEqual(Evaluate.ToString(), EvaluateN.ToString());
                Assert.AreEqual("1", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843b()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet("test") as XSSFSheet;
                XSSFRow row = sheet.CreateRow(0) as XSSFRow;
                XSSFRow row2 = sheet.CreateRow(1) as XSSFRow;
                XSSFCell cellA2 = row2.CreateCell(0, CellType.Formula) as XSSFCell;
                XSSFCell cellB1 = row.CreateCell(1, CellType.Numeric) as XSSFCell;
                cellB1.SetCellValue(10);
                XSSFFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;

                cellA2.SetCellFormula("IF(B1=0,\"\",((ROW())))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                Assert.AreEqual("2", Evaluate.FormatAsString());

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW())),\"\")");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);

                Assert.AreEqual(Evaluate.ToString(), EvaluateN.ToString());
                Assert.AreEqual("2", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843c()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet("test") as XSSFSheet;
                XSSFRow row = sheet.CreateRow(0) as XSSFRow;
                XSSFRow row2 = sheet.CreateRow(1) as XSSFRow;
                XSSFCell cellA2 = row2.CreateCell(0, CellType.Formula) as XSSFCell;
                XSSFCell cellB1 = row.CreateCell(1, CellType.Numeric) as XSSFCell;
                cellB1.SetCellValue(10);
                XSSFFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW())))");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);
                Assert.AreEqual("2", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843d()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet("test") as XSSFSheet;
                XSSFRow row = sheet.CreateRow(0) as XSSFRow;
                XSSFRow row2 = sheet.CreateRow(1) as XSSFRow;
                XSSFCell cellA2 = row2.CreateCell(0, CellType.Formula) as XSSFCell;
                XSSFCell cellB1 = row.CreateCell(1, CellType.Numeric) as XSSFCell;
                cellB1.SetCellValue(10);
                XSSFFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW())),\"\")");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);
                Assert.AreEqual("2", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843e()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet("test") as XSSFSheet;
                XSSFRow row = sheet.CreateRow(0) as XSSFRow;
                XSSFRow row2 = sheet.CreateRow(1) as XSSFRow;
                XSSFCell cellA2 = row2.CreateCell(0, CellType.Formula) as XSSFCell;
                XSSFCell cellB1 = row.CreateCell(1, CellType.Numeric) as XSSFCell;
                cellB1.SetCellValue(10);
                XSSFFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;

                cellA2.SetCellFormula("IF(B1=0,\"\",((ROW())))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                Assert.AreEqual("2", Evaluate.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }


        [Test]
        public void TestBug55843f()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet("test") as XSSFSheet;
                XSSFRow row = sheet.CreateRow(0) as XSSFRow;
                XSSFRow row2 = sheet.CreateRow(1) as XSSFRow;
                XSSFCell cellA2 = row2.CreateCell(0, CellType.Formula) as XSSFCell;
                XSSFCell cellB1 = row.CreateCell(1, CellType.Numeric) as XSSFCell;
                cellB1.SetCellValue(10);
                XSSFFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;

                cellA2.SetCellFormula("IF(B1=0,\"\",IF(B1=10,3,4))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                Assert.AreEqual("3", Evaluate.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }
        [Test]
        public void TestBug56655()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            setCellFormula(sheet, 0, 0, "#VALUE!");
            setCellFormula(sheet, 0, 1, "SUMIFS(A:A,A:A,#VALUE!)");

            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            Assert.AreEqual(CellType.Error, getCell(sheet, 0, 0).CachedFormulaResultType);
            Assert.AreEqual(FormulaError.VALUE.Code, getCell(sheet, 0, 0).ErrorCellValue);
            Assert.AreEqual(CellType.Error, getCell(sheet, 0, 1).CachedFormulaResultType);
            Assert.AreEqual(FormulaError.VALUE.Code, getCell(sheet, 0, 1).ErrorCellValue);

            wb.Close();
        }
        [Test]
        public void TestBug56655a()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();

            setCellFormula(sheet, 0, 0, "B1*C1");
            sheet.GetRow(0).CreateCell(1).SetCellValue("A");
            setCellFormula(sheet, 1, 0, "B1*C1");
            sheet.GetRow(1).CreateCell(1).SetCellValue("A");
            setCellFormula(sheet, 0, 3, "SUMIFS(A:A,A:A,A2)");

            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            Assert.AreEqual(CellType.Error, getCell(sheet, 0, 0).CachedFormulaResultType);
            Assert.AreEqual(FormulaError.VALUE.Code, getCell(sheet, 0, 0).ErrorCellValue);
            Assert.AreEqual(CellType.Error, getCell(sheet, 1, 0).CachedFormulaResultType);
            Assert.AreEqual(FormulaError.VALUE.Code, getCell(sheet, 1, 0).ErrorCellValue);
            Assert.AreEqual(CellType.Error, getCell(sheet, 0, 3).CachedFormulaResultType);
            Assert.AreEqual(FormulaError.VALUE.Code, getCell(sheet, 0, 3).ErrorCellValue);

            wb.Close();
        }

        // bug 57721
        [Test]
        public void structuredReferences()
        {
            verifyAllFormulasInWorkbookCanBeEvaluated("evaluate_formula_with_structured_table_references.xlsx");
        }

        // bug 57840
        [Ignore("Takes over a minute to evaluate all formulas in this large workbook. Run this test when profiling for formula evaluation speed.")]
        [Test]
        public void TestLotsOfFormulasWithStructuredReferencesToCalculatedTableColumns()
        {
            verifyAllFormulasInWorkbookCanBeEvaluated("StructuredRefs-lots-with-lookups.xlsx");
        }
        // FIXME: use junit4 parameterization
        private static void verifyAllFormulasInWorkbookCanBeEvaluated(String sampleWorkbook)
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook(sampleWorkbook);
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
            wb.Close();
        }


        /**
        * @param row 0-based
        * @param column 0-based
        */
        private void setCellFormula(ISheet sheet, int row, int column, String formula)
        {
            IRow r = sheet.GetRow(row);
            if (r == null)
            {
                r = sheet.CreateRow(row);
            }
            ICell cell = r.GetCell(column);
            if (cell == null)
            {
                cell = r.CreateCell(column);
            }
            cell.SetCellType(CellType.Formula);
            cell.CellFormula = (formula);
        }

        /**
         * @param rowNo 0-based
         * @param column 0-based
         */
        private ICell getCell(ISheet sheet, int rowNo, int column)
        {
            return sheet.GetRow(rowNo).GetCell(column);
        }

        [Test]
        public void Test59736()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("59736.xlsx");
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            ICell cell = wb.GetSheetAt(0).GetRow(0).GetCell(0);
            Assert.AreEqual(1, cell.NumericCellValue, 0.001);
            cell = wb.GetSheetAt(0).GetRow(1).GetCell(0);
            CellValue value = evaluator.Evaluate(cell);
            Assert.AreEqual(1, value.NumberValue, 0.001);
            cell = wb.GetSheetAt(0).GetRow(2).GetCell(0);
            value = evaluator.Evaluate(cell);
            Assert.AreEqual(1, value.NumberValue, 0.001);
        }

        [Test]
        public void EvaluateInCellReturnsSameDataType()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            wb.CreateSheet().CreateRow(0).CreateCell(0);
            XSSFFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;
            XSSFCell cell = wb.GetSheetAt(0).GetRow(0).GetCell(0) as XSSFCell;
            XSSFCell same = evaluator.EvaluateInCell(cell) as XSSFCell;
            //assertSame(cell, same);
            Assert.AreSame(cell, same);
            wb.Close();
        }

        [Test]
        public void TestNPOIIssue_1057()
        {
            BaseTestNPOIIssue_1057("xparams.xlsx", "xinstall.xlsx");
        }
    }

}
