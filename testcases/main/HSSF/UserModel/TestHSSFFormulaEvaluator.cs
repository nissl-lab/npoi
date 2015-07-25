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

namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    using NUnit.Framework;
    using TestCases.HSSF;
    using NPOI.SS.Formula;
    using TestCases.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.HSSF.Record;
    using TestCases.SS.UserModel;
    /**
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestHSSFFormulaEvaluator : BaseTestFormulaEvaluator
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }
        public TestHSSFFormulaEvaluator()
            : base(HSSFITestDataProvider.Instance)
        {
        }
        /**
         * Test that the HSSFFormulaEvaluator can Evaluate simple named ranges
         *  (single cells and rectangular areas)
         */
        [Test]
        public void TestEvaluateSimple()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("TestNames.xls");
            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);
            ICell cell = sheet.GetRow(8).GetCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            NPOI.SS.UserModel.CellValue cv = fe.Evaluate(cell);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.Numeric, cv.CellType);
            Assert.AreEqual(3.72, cv.NumberValue, 0.0);
        }
        

        private static void setValue(NPOI.SS.UserModel.ISheet sheet, int rowIndex, int colIndex, double value)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }
            row.CreateCell(colIndex).SetCellValue(value);
        }

        

        /**
         * When evaluating defined names, POI has to decide whether it is capable.  Currently
         * (May2009) POI only supports simple cell and area refs.<br/>
         * The sample spreadsheet (bugzilla attachment 23508) had a name flagged as 'complex'
         * which contained a simple area ref.  It is not clear what the 'complex' flag is used
         * for but POI should look elsewhere to decide whether it can evaluate the name.
         */
        [Test]
        public void TestDefinedNameWithComplexFlag_bug47048()
        {
            // Mock up a spreadsheet to match the critical details of the sample
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Input");
            IName definedName = wb.CreateName();
            definedName.NameName = ("Is_Multicar_Vehicle");
            definedName.RefersToFormula = ("Input!$B$17:$G$17");

            // Set up some data and the formula
            IRow row17 = sheet.CreateRow(16);
            row17.CreateCell(0).SetCellValue(25.0);
            row17.CreateCell(1).SetCellValue(1.33);
            row17.CreateCell(2).SetCellValue(4.0);

            IRow row = sheet.CreateRow(0);
            ICell cellA1 = row.CreateCell(0);
            cellA1.CellFormula = ("SUM(Is_Multicar_Vehicle)");

            // Set the complex flag - POI doesn't usually manipulate this flag
            NameRecord nameRec = TestHSSFName.GetNameRecord(definedName);
            nameRec.OptionFlag = (short)0x10; // 0x10 -> complex

            HSSFFormulaEvaluator hsf = new HSSFFormulaEvaluator(wb);
            CellValue value;
            try
            {
                value = hsf.Evaluate(cellA1);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Don't now how to evalate name 'Is_Multicar_Vehicle'"))
                {
                    throw new AssertionException("Identified bug 47048a");
                }
                throw e;
            }

            Assert.AreEqual(CellType.Numeric, value.CellType);
            Assert.AreEqual(5.33, value.NumberValue, 0.0);
        }
        private class EvalCountListener : EvaluationListener
        {
            private int _evalCount;
            public EvalCountListener()
            {
                _evalCount = 0;
            }
            public override void OnStartEvaluate(IEvaluationCell cell, ICacheEntry entry)
            {
                _evalCount++;
            }
            public int EvalCount
            {
                get
                {
                    return _evalCount;
                }
            }
        }

        /**
         * The HSSFFormula evaluator performance benefits greatly from caching of intermediate cell values
         */
        [Test]
        public void TestShortCircuitIfEvaluation()
        {

            // Set up a simple IF() formula that has measurable evaluation cost for its operands.
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cellA1 = row.CreateCell(0);
            cellA1.CellFormula = "if(B1,C1,D1+E1+F1)";
            // populate cells B1..F1 with simple formulas instead of plain values so we can use
            // EvaluationListener to check which parts of the first formula get evaluated
            for (int i = 1; i < 6; i++)
            {
                // formulas are just literal constants "1".."5"
                row.CreateCell(i).CellFormula = i.ToString();
            }

            EvalCountListener evalListener = new EvalCountListener();
            WorkbookEvaluator evaluator = WorkbookEvaluatorTestHelper.CreateEvaluator(wb, evalListener);
            ValueEval ve = evaluator.Evaluate(HSSFEvaluationTestHelper.WrapCell(cellA1));
            int evalCount = evalListener.EvalCount;
            if (evalCount == 6)
            {
                // Without short-circuit-if evaluation, evaluating cell 'A1' takes 3 extra evaluations (for D1,E1,F1)
                throw new AssertionException("Identifed bug 48195 - Formula evaluator should short-circuit IF() calculations.");
            }
            Assert.AreEqual(3, evalCount);
            Assert.AreEqual(2.0, ((NumberEval)ve).NumberValue, 0D);
        }
        /**
     * Ensures that we can handle NameXPtgs in the formulas
     *  we Parse.
     */
        [Test]
        public void TestXRefs()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("XRefCalc.xls");
            IWorkbook wbData = HSSFTestDataSamples.OpenSampleWorkbook("XRefCalcData.xls");
            ICell cell;

            // VLookup on a name in another file
            cell = wb.GetSheetAt(0).GetRow(1).GetCell(2);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual(CellType.Numeric, cell.CachedFormulaResultType);
            Assert.AreEqual(12.30, cell.NumericCellValue, 0.0001);
            // WARNING - this is wrong!
            // The file name should be Showing, but bug #45970 is fixed
            //  we seem to loose it
            Assert.AreEqual("VLOOKUP(PART,COSTS,2,FALSE)", cell.CellFormula);


            // Simple reference to a name in another file
            cell = wb.GetSheetAt(0).GetRow(1).GetCell(4);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual(CellType.Numeric, cell.CachedFormulaResultType);
            Assert.AreEqual(36.90, cell.NumericCellValue, 0.0001);
            // WARNING - this is wrong!
            // The file name should be Showing, but bug #45970 is fixed
            //  we seem to loose it
            Assert.AreEqual("Cost*Markup_Cost", cell.CellFormula);


            // Evaluate the cells
            HSSFFormulaEvaluator eval = new HSSFFormulaEvaluator(wb);
            HSSFFormulaEvaluator.SetupEnvironment(
                  new String[] { "XRefCalc.xls", "XRefCalcData.xls" },
                  new HSSFFormulaEvaluator[] {
                  eval,
                  new HSSFFormulaEvaluator(wbData)
            }
            );
            eval.EvaluateFormulaCell(
                  wb.GetSheetAt(0).GetRow(1).GetCell(2)
            );
            eval.EvaluateFormulaCell(
                  wb.GetSheetAt(0).GetRow(1).GetCell(4)
            );


            // Re-check VLOOKUP one
            cell = wb.GetSheetAt(0).GetRow(1).GetCell(2);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual(CellType.Numeric, cell.CachedFormulaResultType);
            Assert.AreEqual(12.30, cell.NumericCellValue, 0.0001);

            // Re-check ref one
            cell = wb.GetSheetAt(0).GetRow(1).GetCell(4);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual(CellType.Numeric, cell.CachedFormulaResultType);
            Assert.AreEqual(36.90, cell.NumericCellValue, 0.0001);

            // Add a formula that refers to one of the existing external workbooks
            cell = wb.GetSheetAt(0).GetRow(1).CreateCell(40);
            cell.CellFormula = (/*setter*/"Cost*[XRefCalcData.xls]MarkupSheet!$B$1");

            // Check is was stored correctly
            Assert.AreEqual("Cost*[XRefCalcData.xls]MarkupSheet!$B$1", cell.CellFormula);

            // Check it Evaluates correctly
            eval.EvaluateFormulaCell(cell);
            Assert.AreEqual(24.60 * 1.8, cell.NumericCellValue);


            // Try to add a formula for a new external workbook, won't be allowed to start
            try
            {
                cell = wb.GetSheetAt(0).GetRow(1).CreateCell(42);
                cell.CellFormula = (/*setter*/"[alt.xls]Sheet0!$A$1");
                Assert.Fail("New workbook not linked, shouldn't be able to Add");
            }
            catch (Exception) { }

            // Link our new workbook
            HSSFWorkbook alt = new HSSFWorkbook();
            alt.CreateSheet().CreateRow(0).CreateCell(0).SetCellValue("In another workbook");
            wb.LinkExternalWorkbook("alt.xls", alt);

            // Now add a formula that refers to our new workbook
            cell.CellFormula = (/*setter*/"[alt.xls]Sheet0!$A$1");
            Assert.AreEqual("[alt.xls]Sheet0!$A$1", cell.CellFormula);

            // Evaluate it, without a link to that workbook
            try
            {
                eval.Evaluate(cell);
                Assert.Fail("No cached value and no link to workbook, shouldn't Evaluate");
            }
            catch (Exception) { }

            // Add a link, check it does
            HSSFFormulaEvaluator.SetupEnvironment(
                    new String[] { "XRefCalc.xls", "XRefCalcData.xls", "alt.xls" },
                    new HSSFFormulaEvaluator[] {
                    eval,
                    new HSSFFormulaEvaluator(wbData),
                    new HSSFFormulaEvaluator(alt)
              }
            );
            eval.EvaluateFormulaCell(cell);
            Assert.AreEqual("In another workbook", cell.StringCellValue);


            // Save and re-load
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb as HSSFWorkbook);
            eval = new HSSFFormulaEvaluator(wb);
            HSSFFormulaEvaluator.SetupEnvironment(
                    new String[] { "XRefCalc.xls", "XRefCalcData.xls", "alt.xls" },
                    new HSSFFormulaEvaluator[] {
                    eval,
                    new HSSFFormulaEvaluator(wbData),
                    new HSSFFormulaEvaluator(alt)
              }
            );

            // Check the one referring to the previously existing workbook behaves
            cell = wb.GetSheetAt(0).GetRow(1).GetCell(40);
            Assert.AreEqual("Cost*[XRefCalcData.xls]MarkupSheet!$B$1", cell.CellFormula);
            eval.EvaluateFormulaCell(cell);
            Assert.AreEqual(24.60 * 1.8, cell.NumericCellValue);

            // Now check the newly Added reference
            cell = wb.GetSheetAt(0).GetRow(1).GetCell(42);
            Assert.AreEqual("[alt.xls]Sheet0!$A$1", cell.CellFormula);
            eval.EvaluateFormulaCell(cell);
            Assert.AreEqual("In another workbook", cell.StringCellValue);

        }
        [Test]
        public void TestSharedFormulas()
        {
            BaseTestSharedFormulas("shared_formulas.xls");
        }
    }
}