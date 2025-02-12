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

namespace TestCases.SS.UserModel
{
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;

    /**
     * Common superclass for Testing implementatiosn of{@link FormulaEvaluator}
     *
     * @author Yegor Kozlov
     */
    public abstract class BaseTestFormulaEvaluator
    {

        protected ITestDataProvider _testDataProvider;
        //public BaseTestFormulaEvaluator()
        //    : this(TestCases.HSSF.HSSFITestDataProvider.Instance)
        //{ }
        /**
         * @param TestDataProvider an object that provides Test data in  /  specific way
         */
        protected BaseTestFormulaEvaluator(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }
        [Test]
        public void TestSimpleArithmetic()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);

            ICell c1 = r.CreateCell(0);
            c1.CellFormula = (/*setter*/"1+5");
            Assert.AreEqual(0.0, c1.NumericCellValue, 0.0);

            ICell c2 = r.CreateCell(1);
            c2.CellFormula = (/*setter*/"10/2");
            Assert.AreEqual(0.0, c2.NumericCellValue, 0.0);

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            fe.EvaluateFormulaCell(c1);
            fe.EvaluateFormulaCell(c2);

            Assert.AreEqual(6.0, c1.NumericCellValue, 0.0001);
            Assert.AreEqual(5.0, c2.NumericCellValue, 0.0001);

            wb.Close();
        }
        [Test]
        public void TestSumCount()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet s = wb.CreateSheet();
            IRow r = s.CreateRow(0);
            r.CreateCell(0).SetCellValue(2.5);
            r.CreateCell(1).SetCellValue(1.1);
            r.CreateCell(2).SetCellValue(3.2);
            r.CreateCell(4).SetCellValue(10.7);

            r = s.CreateRow(1);

            ICell c1 = r.CreateCell(0);
            c1.CellFormula = (/*setter*/"SUM(A1:B1)");
            Assert.AreEqual(0.0, c1.NumericCellValue, 0.0);

            ICell c2 = r.CreateCell(1);
            c2.CellFormula = (/*setter*/"SUM(A1:E1)");
            Assert.AreEqual(0.0, c2.NumericCellValue, 0.0);

            ICell c3 = r.CreateCell(2);
            c3.CellFormula = (/*setter*/"COUNT(A1:A1)");
            Assert.AreEqual(0.0, c3.NumericCellValue, 0.0);

            ICell c4 = r.CreateCell(3);
            c4.CellFormula = (/*setter*/"COUNTA(A1:E1)");
            Assert.AreEqual(0.0, c4.NumericCellValue, 0.0);


            // Evaluate and Test
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            fe.EvaluateFormulaCell(c1);
            fe.EvaluateFormulaCell(c2);
            fe.EvaluateFormulaCell(c3);
            fe.EvaluateFormulaCell(c4);

            Assert.AreEqual(3.6, c1.NumericCellValue, 0.0001);
            Assert.AreEqual(17.5, c2.NumericCellValue, 0.0001);
            Assert.AreEqual(1, c3.NumericCellValue, 0.0001);
            Assert.AreEqual(4, c4.NumericCellValue, 0.0001);

            wb.Close();
        }

        public void BaseTestSharedFormulas(String sampleFile)
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(sampleFile);

            ISheet sheet = wb.GetSheetAt(0);

            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            ICell cell;

            cell = sheet.GetRow(1).GetCell(0);
            Assert.AreEqual("B2", cell.CellFormula);
            Assert.AreEqual("ProductionOrderConfirmation", Evaluator.Evaluate(cell).StringValue);

            cell = sheet.GetRow(2).GetCell(0);
            Assert.AreEqual("B3", cell.CellFormula);
            Assert.AreEqual("RequiredAcceptanceDate", Evaluator.Evaluate(cell).StringValue);

            cell = sheet.GetRow(3).GetCell(0);
            Assert.AreEqual("B4", cell.CellFormula);
            Assert.AreEqual("Header", Evaluator.Evaluate(cell).StringValue);

            cell = sheet.GetRow(4).GetCell(0);
            Assert.AreEqual("B5", cell.CellFormula);
            Assert.AreEqual("UniqueDocumentNumberID", Evaluator.Evaluate(cell).StringValue);

            wb.Close();
        }

        /**
         * Test creation / Evaluation of formulas with sheet-level names
         */
        [Test]
        public void TestSheetLevelFormulas()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            IRow row;
            ISheet sh1 = wb.CreateSheet("Sheet1");
            IName nm1 = wb.CreateName();
            nm1.NameName = (/*setter*/"sales_1");
            nm1.SheetIndex = (/*setter*/0);
            nm1.RefersToFormula = (/*setter*/"Sheet1!$A$1");
            row = sh1.CreateRow(0);
            row.CreateCell(0).SetCellValue(3);
            row.CreateCell(1).CellFormula = (/*setter*/"sales_1");
            row.CreateCell(2).CellFormula = (/*setter*/"sales_1*2");

            ISheet sh2 = wb.CreateSheet("Sheet2");
            IName nm2 = wb.CreateName();
            nm2.NameName = (/*setter*/"sales_1");
            nm2.SheetIndex = (/*setter*/1);
            nm2.RefersToFormula = (/*setter*/"Sheet2!$A$1");

            row = sh2.CreateRow(0);
            row.CreateCell(0).SetCellValue(5);
            row.CreateCell(1).CellFormula = (/*setter*/"sales_1");
            row.CreateCell(2).CellFormula = (/*setter*/"sales_1*3");

            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual(3.0, Evaluator.Evaluate(sh1.GetRow(0).GetCell(1)).NumberValue, 0.0);
            Assert.AreEqual(6.0, Evaluator.Evaluate(sh1.GetRow(0).GetCell(2)).NumberValue, 0.0);

            Assert.AreEqual(5.0, Evaluator.Evaluate(sh2.GetRow(0).GetCell(1)).NumberValue, 0.0);
            Assert.AreEqual(15.0, Evaluator.Evaluate(sh2.GetRow(0).GetCell(2)).NumberValue, 0.0);

            wb.Close();
        }
        [Test]
        public void TestFullColumnRefs()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cell0 = row.CreateCell(0);
            cell0.CellFormula = (/*setter*/"sum(D:D)");
            ICell cell1 = row.CreateCell(1);
            cell1.CellFormula = (/*setter*/"sum(D:E)");

            // some values in column D
            SetValue(sheet, 1, 3, 5.0);
            SetValue(sheet, 2, 3, 6.0);
            SetValue(sheet, 5, 3, 7.0);
            SetValue(sheet, 50, 3, 8.0);

            // some values in column E
            SetValue(sheet, 1, 4, 9.0);
            SetValue(sheet, 2, 4, 10.0);
            SetValue(sheet, 30000, 4, 11.0);

            // some other values
            SetValue(sheet, 1, 2, 100.0);
            SetValue(sheet, 2, 5, 100.0);
            SetValue(sheet, 3, 6, 100.0);


            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual(26.0, fe.Evaluate(cell0).NumberValue, 0.0);
            Assert.AreEqual(56.0, fe.Evaluate(cell1).NumberValue, 0.0);

            wb.Close();
        }
        [Test]
        public void TestRepeatedEvaluation()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow r = sheet.CreateRow(0);
            ICell c = r.CreateCell(0, CellType.Formula);

            // Create a value and check it
            c.CellFormula = (/*setter*/"Date(2011,10,6)");
            CellValue cellValue = fe.Evaluate(c);
            Assert.AreEqual(40822.0, cellValue.NumberValue, 0.0);
            cellValue = fe.Evaluate(c);
            Assert.AreEqual(40822.0, cellValue.NumberValue, 0.0);

            // Change it
            c.CellFormula = (/*setter*/"Date(2011,10,4)");

            // Evaluate it, no change as the formula Evaluator
            //  won't know to clear the cache
            cellValue = fe.Evaluate(c);
            Assert.AreEqual(40822.0, cellValue.NumberValue, 0.0);

            // Manually flush for this cell, and check
            fe.NotifySetFormula(c);
            cellValue = fe.Evaluate(c);
            Assert.AreEqual(40820.0, cellValue.NumberValue, 0.0);

            // Change again, without Notifying
            c.CellFormula = (/*setter*/"Date(2010,10,4)");
            cellValue = fe.Evaluate(c);
            Assert.AreEqual(40820.0, cellValue.NumberValue, 0.0);

            // Now manually clear all, will see the new value
            fe.ClearAllCachedResultValues();
            cellValue = fe.Evaluate(c);
            Assert.AreEqual(40455.0, cellValue.NumberValue, 0.0);

            wb.Close();
        }

        private static void SetValue(ISheet sheet, int rowIndex, int colIndex, double value)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }
            row.CreateCell(colIndex).SetCellValue(value);
        }

        /**
         * {@link FormulaEvaluator#Evaluate(NPOI.SS.UserModel.Cell)} should behave the same whether the cell
         * is <code>null</code> or blank.
         */
        [Test]
        public void TestEvaluateBlank()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.IsNull(fe.Evaluate(null));
            ISheet sheet = wb.CreateSheet("Sheet1");
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            Assert.IsNull(fe.Evaluate(cell));

            wb.Close();
        }

        /**
         * Test for bug due to attempt to convert a cached formula error result to a bool
         */
        [Test]
        public virtual void TestUpdateCachedFormulaResultFromErrorToNumber_bug46479()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cellA1 = row.CreateCell(0);
            ICell cellB1 = row.CreateCell(1);
            cellB1.SetCellFormula("A1+1");
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            cellA1.SetCellErrorValue(FormulaError.NAME.Code);
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
                    Assert.Fail("Identified bug 46479a");
                }
            }
            Assert.AreEqual(3.5, cellB1.NumericCellValue, 0.0);

            wb.Close();
        }

        [Test]
        public void TestRounding_bug51339()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            ICell cellA1 = row.CreateCell(0);
            cellA1.SetCellValue(2162.615d);
            ICell cellB1 = row.CreateCell(1);
            cellB1.CellFormula = (/*setter*/"round(a1,2)");
            ICell cellC1 = row.CreateCell(2);
            cellC1.CellFormula = (/*setter*/"roundup(a1,2)");
            ICell cellD1 = row.CreateCell(3);
            cellD1.CellFormula = (/*setter*/"rounddown(a1,2)");
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            Assert.AreEqual(2162.62, fe.EvaluateInCell(cellB1).NumericCellValue, 0.0);
            Assert.AreEqual(2162.62, fe.EvaluateInCell(cellC1).NumericCellValue, 0.0);
            Assert.AreEqual(2162.61, fe.EvaluateInCell(cellD1).NumericCellValue, 0.0);

            wb.Close();
        }

        [Test]
        public void EvaluateInCellReturnsSameCell()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet().CreateRow(0).CreateCell(0);
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            ICell cell = wb.GetSheetAt(0).GetRow(0).GetCell(0);
            ICell same = evaluator.EvaluateInCell(cell);
            //assertSame(cell, same);
            Assert.AreSame(cell, same);
            wb.Close();
        }

        public void BaseTestNPOIIssue_1057(string paramsFile, string installFile)
        {
            DataFormatter formatter = new DataFormatter();
            var paramswb = _testDataProvider.OpenSampleWorkbook(paramsFile);
            var installwb = _testDataProvider.OpenSampleWorkbook(installFile);
            
            var paramsevaluator = paramswb.GetCreationHelper().CreateFormulaEvaluator();
            var installevaluator = installwb.GetCreationHelper().CreateFormulaEvaluator();

            var refs = new Dictionary<string, IFormulaEvaluator>();
            refs.Add(paramsFile, paramsevaluator);
            refs.Add(installFile, installevaluator);
            installevaluator.SetupReferencedWorkbooks(refs);

            var suffix = installFile[0].ToString().ToUpper();
            var sheet = installwb.GetSheetAt(0);
            var row = sheet.GetRow(0);
            var cell = row.GetCell(1);

            
            String cellValue = formatter.FormatCellValue(cell, installevaluator);
            Assert.AreEqual("Referenced value in sheet 1" + suffix, cellValue);

            row = sheet.GetRow(1);
            cell = row.GetCell(1);
            cellValue = formatter.FormatCellValue(cell, installevaluator);
            Assert.AreEqual("Referenced value in sheet 2" + suffix, cellValue);
        }
    }
}