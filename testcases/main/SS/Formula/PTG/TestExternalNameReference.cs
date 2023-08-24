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

namespace TestCases.SS.Formula.PTG
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using TestCases.HSSF;
    /**
     * Tests for proper calculation of named ranges from external workbooks.
     * 
     * 
     * @author Stephen Wolke (smwolke at geistig.com)
     */
    [TestFixture]
    public class TestExternalNameReference
    {
        // not used: double MARKUP_COST = 1.9d;
        double MARKUP_COST_1 = 1.8d;
        double MARKUP_COST_2 = 1.5d;
        double PART_COST = 12.3d;
        double NEW_QUANT = 7.0d;
        double NEW_PART_COST = 15.3d;

        /// <summary>
        ///  Some of the tests are depending on the american culture.
        /// </summary>
        [SetUp]
        public void InitializeCultere()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
        }

        /**
         * Tests <c>NameXPtg for external cell reference by name</c> and logic in Workbook below that   
         */
        [Test]
        public void TestReadCalcSheet()
        {
            try
            {
                HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("XRefCalc.xls");
                Assert.AreEqual("Sheet1!$A$2", wb.GetName("QUANT").RefersToFormula);
                Assert.AreEqual("Sheet1!$B$2", wb.GetName("PART").RefersToFormula);
                Assert.AreEqual("x123", wb.GetSheet("Sheet1").GetRow(1).GetCell(1).StringCellValue);
                Assert.AreEqual("Sheet1!$C$2", wb.GetName("UNITCOST").RefersToFormula);
                CellReference cellRef = new CellReference(wb.GetName("UNITCOST").RefersToFormula);
                ICell cell = wb.GetSheet(cellRef.SheetName).GetRow(cellRef.Row).GetCell((int)cellRef.Col);
                Assert.AreEqual("VLOOKUP(PART,COSTS,2,FALSE)", cell.CellFormula);
                Assert.AreEqual("Sheet1!$D$2", wb.GetName("COST").RefersToFormula);
                cellRef = new CellReference(wb.GetName("COST").RefersToFormula);
                cell = wb.GetSheet(cellRef.SheetName).GetRow(cellRef.Row).GetCell((int)cellRef.Col);
                Assert.AreEqual("UNITCOST*Quant", cell.CellFormula);
                Assert.AreEqual("Sheet1!$E$2", wb.GetName("TOTALCOST").RefersToFormula);
                cellRef = new CellReference(wb.GetName("TOTALCOST").RefersToFormula);
                cell = wb.GetSheet(cellRef.SheetName).GetRow(cellRef.Row).GetCell((int)cellRef.Col);
                Assert.AreEqual("Cost*Markup_Cost", cell.CellFormula);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }
        [Test]
        public void TestReadReferencedSheet()
        {
            try
            {
                HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("XRefCalcData.xls");
                Assert.AreEqual("CostSheet!$A$2:$B$3", wb.GetName("COSTS").RefersToFormula);
                Assert.AreEqual("x123", wb.GetSheet("CostSheet").GetRow(1).GetCell(0).StringCellValue);
                Assert.AreEqual(PART_COST, wb.GetSheet("CostSheet").GetRow(1).GetCell(1).NumericCellValue);
                Assert.AreEqual("MarkupSheet!$B$1", wb.GetName("Markup_Cost").RefersToFormula);
                Assert.AreEqual(MARKUP_COST_1, wb.GetSheet("MarkupSheet").GetRow(0).GetCell(1).NumericCellValue);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }
        [Test]
        public void TestEvaluate()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("XRefCalc.xls");
            HSSFWorkbook wb2 = HSSFTestDataSamples.OpenSampleWorkbook("XRefCalcData.xls");
            CellReference cellRef = new CellReference(wb.GetName("QUANT").RefersToFormula);
            ICell cell = wb.GetSheet(cellRef.SheetName).GetRow(cellRef.Row).GetCell((int)cellRef.Col);
            cell.SetCellValue(NEW_QUANT);
            cell = wb2.GetSheet("CostSheet").GetRow(1).GetCell(1);
            cell.SetCellValue(NEW_PART_COST);
            HSSFFormulaEvaluator Evaluator = new HSSFFormulaEvaluator(wb);
            HSSFFormulaEvaluator EvaluatorCost = new HSSFFormulaEvaluator(wb2);
            String[] bookNames = { "XRefCalc.xls", "XRefCalcData.xls" };
            HSSFFormulaEvaluator[] Evaluators = { Evaluator, EvaluatorCost, };
            HSSFFormulaEvaluator.SetupEnvironment(bookNames, Evaluators);
            cellRef = new CellReference(wb.GetName("UNITCOST").RefersToFormula);
            ICell uccell = wb.GetSheet(cellRef.SheetName).GetRow(cellRef.Row).GetCell((int)cellRef.Col);
            cellRef = new CellReference(wb.GetName("COST").RefersToFormula);
            ICell ccell = wb.GetSheet(cellRef.SheetName).GetRow(cellRef.Row).GetCell((int)cellRef.Col);
            cellRef = new CellReference(wb.GetName("TOTALCOST").RefersToFormula);
            ICell tccell = wb.GetSheet(cellRef.SheetName).GetRow(cellRef.Row).GetCell((int)cellRef.Col);
            Evaluator.EvaluateFormulaCell(uccell);
            Evaluator.EvaluateFormulaCell(ccell);
            Evaluator.EvaluateFormulaCell(tccell);
            Assert.AreEqual(NEW_PART_COST, uccell.NumericCellValue);
            double ctotal = decimal.ToDouble((decimal)NEW_PART_COST * (decimal)NEW_QUANT);
            Assert.AreEqual(ctotal, ccell.NumericCellValue);
            Assert.AreEqual(NEW_PART_COST * NEW_QUANT * MARKUP_COST_2, tccell.NumericCellValue);
        }
    }

}