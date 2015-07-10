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
namespace NPOI.XSSF.UserModel
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
            Assert.AreEqual(result, Evaluator.EvaluateInCell(b3).NumericCellValue);

            //at this point the master formula is gone, but we are still able to Evaluate dependent cells
            ICell c3 = sheet.GetRow(2).GetCell(2);
            Assert.AreEqual(result, Evaluator.EvaluateInCell(c3).NumericCellValue);

            ICell d3 = sheet.GetRow(2).GetCell(3);
            Assert.AreEqual(result, Evaluator.EvaluateInCell(d3).NumericCellValue);
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
                Assert.AreEqual(cv_noformula.NumberValue, cv_formula.NumberValue, "Wrong Evaluation result in " + ref_formula.FormatAsString());
            }

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
            Assert.AreEqual(142.0, cXSLX_gNR.NumericCellValue);

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
            Assert.AreEqual(142.0, cXSL_gNR.NumericCellValue);

            // Try to Evaluate without references, won't work
            // (At least, not unit we fix bug #56752 that is1)
            try
            {
                evaluator.Evaluate(cXSL_cell);
                Assert.Fail("Without a fix for #56752, shouldn't be able to Evaluate a " +
                     "reference to a non-provided linked workbook");
            }
            catch (Exception e) { }

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

            Assert.AreEqual("\"Hello!\"", evaluator.Evaluate(cXSLX_cell).FormatAsString());
            // TODO Fix XSSF reference Evaluations to work
            //        Assert.AreEqual("\"Test A1\"", Evaluator.Evaluate(cXSLX_sNR).FormatAsString());
            //        Assert.AreEqual("142.0",   Evaluator.Evaluate(cXSLX_gNR).FormatAsString());

            Assert.AreEqual("\"Hello!\"", evaluator.Evaluate(cXSL_cell).FormatAsString());
            Assert.AreEqual("\"Test A1\"", evaluator.Evaluate(cXSL_sNR).FormatAsString());
            //Assert.AreEqual("142.0", evaluator.Evaluate(cXSL_gNR).FormatAsString());
            Assert.AreEqual("142", evaluator.Evaluate(cXSL_gNR).FormatAsString());

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

    }

}
