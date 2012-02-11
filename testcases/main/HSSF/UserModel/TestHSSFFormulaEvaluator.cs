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

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TestCases.HSSF;
    using NPOI.SS.Formula;
    using TestCases.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.HSSF.Record;
    /**
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestHSSFFormulaEvaluator
    {
        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [TestInitialize()]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        /**
         * Test that the HSSFFormulaEvaluator can Evaluate simple named ranges
         *  (single cells and rectangular areas)
         */
        [TestMethod]
        public void TestEvaluateSimple()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("TestNames.xls");
            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);
            ICell cell = sheet.GetRow(8).GetCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            NPOI.SS.UserModel.CellValue cv = fe.Evaluate(cell);
            Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, cv.CellType);
            Assert.AreEqual(3.72, cv.NumberValue, 0.0);
        }
        [TestMethod]
        public void TestFullColumnRefs()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            cell0.CellFormula = ("sum(D:D)");
            ICell cell1 = row.CreateCell(1);
            cell1.CellFormula = ("sum(D:E)");

            // some values in column D
            setValue(sheet, 1, 3, 5.0);
            setValue(sheet, 2, 3, 6.0);
            setValue(sheet, 5, 3, 7.0);
            setValue(sheet, 50, 3, 8.0);

            // some values in column E
            setValue(sheet, 1, 4, 9.0);
            setValue(sheet, 2, 4, 10.0);
            setValue(sheet, 30000, 4, 11.0);

            // some other values 
            setValue(sheet, 1, 2, 100.0);
            setValue(sheet, 2, 5, 100.0);
            setValue(sheet, 3, 6, 100.0);


            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            Assert.AreEqual(26.0, fe.Evaluate(cell0).NumberValue, 0.0);
            Assert.AreEqual(56.0, fe.Evaluate(cell1).NumberValue, 0.0);
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
         * {@link HSSFFormulaEvaluator#Evaluate(Cell)} should behave the same whether the cell
         * is <c>null</c> or blank.
         */
        [TestMethod]
        public void TestEvaluateBlank()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            Assert.IsNull(fe.Evaluate(null));
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Sheet1");
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            Assert.IsNull(fe.Evaluate(cell));
        }
        /**
 * Test for bug due to attempt to convert a cached formula error result to a boolean
 */
        [TestMethod]
        [Ignore] // Identified bug 46479a https://issues.apache.org/bugzilla/show_bug.cgi?id=46479 fixed in POI in two commits https://svn.apache.org/viewvc?view=revision&revision=731715 and https://svn.apache.org/viewvc?view=revision&revision=886951
        public void TestUpdateCachedFormulaResultFromErrorToNumber_bug46479()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cellA1 = row.CreateCell(0);
            ICell cellB1 = row.CreateCell(1);
            cellB1.CellFormula = "A1+1";
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            cellA1.SetCellErrorValue((byte)HSSFErrorConstants.ERROR_NAME);
            fe.EvaluateFormulaCell(cellB1);

            cellA1.SetCellValue(2.5);
            fe.NotifyUpdateCell(cellA1);
            try
            {
                fe.EvaluateInCell(cellB1);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals("Cannot get a numeric value from a error formula cell"))
                {
                    throw new AssertFailedException("Identified bug 46479a");
                }
            }
            Assert.AreEqual(3.5, cellB1.NumericCellValue, 0.0);
        }

        /**
         * When evaluating defined names, POI has to decide whether it is capable.  Currently
         * (May2009) POI only supports simple cell and area refs.<br/>
         * The sample spreadsheet (bugzilla attachment 23508) had a name flagged as 'complex'
         * which contained a simple area ref.  It is not clear what the 'complex' flag is used
         * for but POI should look elsewhere to decide whether it can evaluate the name.
         */
        [TestMethod]
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
                    throw new AssertFailedException("Identified bug 47048a");
                }
                throw e;
            }

            Assert.AreEqual(CellType.NUMERIC, value.CellType);
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
        [TestMethod]
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
                throw new AssertFailedException("Identifed bug 48195 - Formula evaluator should short-circuit IF() calculations.");
            }
            Assert.AreEqual(3, evalCount);
            Assert.AreEqual(2.0, ((NumberEval)ve).NumberValue, 0D);
        }
    }
}