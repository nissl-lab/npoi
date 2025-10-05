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
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using SixLabors.Fonts;
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
            ClassicAssert.AreEqual(0.0, c1.NumericCellValue, 0.0);

            ICell c2 = r.CreateCell(1);
            c2.CellFormula = (/*setter*/"10/2");
            ClassicAssert.AreEqual(0.0, c2.NumericCellValue, 0.0);

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            fe.EvaluateFormulaCell(c1);
            fe.EvaluateFormulaCell(c2);

            ClassicAssert.AreEqual(6.0, c1.NumericCellValue, 0.0001);
            ClassicAssert.AreEqual(5.0, c2.NumericCellValue, 0.0001);

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
            ClassicAssert.AreEqual(0.0, c1.NumericCellValue, 0.0);

            ICell c2 = r.CreateCell(1);
            c2.CellFormula = (/*setter*/"SUM(A1:E1)");
            ClassicAssert.AreEqual(0.0, c2.NumericCellValue, 0.0);

            ICell c3 = r.CreateCell(2);
            c3.CellFormula = (/*setter*/"COUNT(A1:A1)");
            ClassicAssert.AreEqual(0.0, c3.NumericCellValue, 0.0);

            ICell c4 = r.CreateCell(3);
            c4.CellFormula = (/*setter*/"COUNTA(A1:E1)");
            ClassicAssert.AreEqual(0.0, c4.NumericCellValue, 0.0);


            // Evaluate and Test
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            fe.EvaluateFormulaCell(c1);
            fe.EvaluateFormulaCell(c2);
            fe.EvaluateFormulaCell(c3);
            fe.EvaluateFormulaCell(c4);

            ClassicAssert.AreEqual(3.6, c1.NumericCellValue, 0.0001);
            ClassicAssert.AreEqual(17.5, c2.NumericCellValue, 0.0001);
            ClassicAssert.AreEqual(1, c3.NumericCellValue, 0.0001);
            ClassicAssert.AreEqual(4, c4.NumericCellValue, 0.0001);

            wb.Close();
        }

        public void BaseTestSharedFormulas(String sampleFile)
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook(sampleFile);

            ISheet sheet = wb.GetSheetAt(0);

            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            ICell cell;

            cell = sheet.GetRow(1).GetCell(0);
            ClassicAssert.AreEqual("B2", cell.CellFormula);
            ClassicAssert.AreEqual("ProductionOrderConfirmation", Evaluator.Evaluate(cell).StringValue);

            cell = sheet.GetRow(2).GetCell(0);
            ClassicAssert.AreEqual("B3", cell.CellFormula);
            ClassicAssert.AreEqual("RequiredAcceptanceDate", Evaluator.Evaluate(cell).StringValue);

            cell = sheet.GetRow(3).GetCell(0);
            ClassicAssert.AreEqual("B4", cell.CellFormula);
            ClassicAssert.AreEqual("Header", Evaluator.Evaluate(cell).StringValue);

            cell = sheet.GetRow(4).GetCell(0);
            ClassicAssert.AreEqual("B5", cell.CellFormula);
            ClassicAssert.AreEqual("UniqueDocumentNumberID", Evaluator.Evaluate(cell).StringValue);

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
            ClassicAssert.AreEqual(3.0, Evaluator.Evaluate(sh1.GetRow(0).GetCell(1)).NumberValue, 0.0);
            ClassicAssert.AreEqual(6.0, Evaluator.Evaluate(sh1.GetRow(0).GetCell(2)).NumberValue, 0.0);

            ClassicAssert.AreEqual(5.0, Evaluator.Evaluate(sh2.GetRow(0).GetCell(1)).NumberValue, 0.0);
            ClassicAssert.AreEqual(15.0, Evaluator.Evaluate(sh2.GetRow(0).GetCell(2)).NumberValue, 0.0);

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
            ClassicAssert.AreEqual(26.0, fe.Evaluate(cell0).NumberValue, 0.0);
            ClassicAssert.AreEqual(56.0, fe.Evaluate(cell1).NumberValue, 0.0);

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
            ClassicAssert.AreEqual(40822.0, cellValue.NumberValue, 0.0);
            cellValue = fe.Evaluate(c);
            ClassicAssert.AreEqual(40822.0, cellValue.NumberValue, 0.0);

            // Change it
            c.CellFormula = (/*setter*/"Date(2011,10,4)");

            // Evaluate it, no change as the formula Evaluator
            //  won't know to clear the cache
            cellValue = fe.Evaluate(c);
            ClassicAssert.AreEqual(40822.0, cellValue.NumberValue, 0.0);

            // Manually flush for this cell, and check
            fe.NotifySetFormula(c);
            cellValue = fe.Evaluate(c);
            ClassicAssert.AreEqual(40820.0, cellValue.NumberValue, 0.0);

            // Change again, without Notifying
            c.CellFormula = (/*setter*/"Date(2010,10,4)");
            cellValue = fe.Evaluate(c);
            ClassicAssert.AreEqual(40820.0, cellValue.NumberValue, 0.0);

            // Now manually clear all, will see the new value
            fe.ClearAllCachedResultValues();
            cellValue = fe.Evaluate(c);
            ClassicAssert.AreEqual(40455.0, cellValue.NumberValue, 0.0);

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
            ClassicAssert.IsNull(fe.Evaluate(null));
            ISheet sheet = wb.CreateSheet("Sheet1");
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            ClassicAssert.IsNull(fe.Evaluate(cell));

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
            ClassicAssert.AreEqual(CellType.Error, fe.EvaluateFormulaCell(cellB1));
            ClassicAssert.AreEqual(CellType.Formula, cellB1.CellType);
            fe.EvaluateFormulaCell(cellB1);

            cellA1.SetCellValue(2.5);
            fe.NotifyUpdateCell(cellA1);
            try
            {
                fe.EvaluateInCell(cellB1);
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals("Cannot get a numeric value from a error formula cell",
                    StringComparison.OrdinalIgnoreCase))
                {
                    Assert.Fail("Identified bug 46479a");
                }
            }
            ClassicAssert.AreEqual(3.5, cellB1.NumericCellValue, 0.0);

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

            ClassicAssert.AreEqual(2162.62, fe.EvaluateInCell(cellB1).NumericCellValue, 0.0);
            ClassicAssert.AreEqual(2162.62, fe.EvaluateInCell(cellC1).NumericCellValue, 0.0);
            ClassicAssert.AreEqual(2162.61, fe.EvaluateInCell(cellD1).NumericCellValue, 0.0);

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
            
            ClassicAssert.AreSame(cell, same);
            wb.Close();
        }

        [Test]
        public void TestBug61148()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            cell.SetCellFormula("1+2");

            ClassicAssert.AreEqual(0, (int)cell.NumericCellValue);
            ClassicAssert.AreEqual("1+2", cell.ToString());

            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            eval.EvaluateInCell(cell);

            ClassicAssert.AreEqual("3", cell.ToString());
        }

        [Test]
        public void TestMultisheetFormulaEval()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet1 = wb.CreateSheet("Sheet1");
                ISheet sheet2 = wb.CreateSheet("Sheet2");
                ISheet sheet3 = wb.CreateSheet("Sheet3");

                // sheet1 A1
                ICell cell = sheet1.CreateRow(0).CreateCell(0);
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(1.0);

                // sheet2 A1
                cell = sheet2.CreateRow(0).CreateCell(0);
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(1.0);

                // sheet2 B1
                cell = sheet2.GetRow(0).CreateCell(1);
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(1.0);

                // sheet3 A1
                cell = sheet3.CreateRow(0).CreateCell(0);
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(1.0);

                // sheet1 A2 formulae
                cell = sheet1.CreateRow(1).CreateCell(0);
                cell.SetCellType(CellType.Formula);
                cell.CellFormula = (/*setter*/"SUM(Sheet1:Sheet3!A1)");

                // sheet1 A3 formulae
                cell = sheet1.CreateRow(2).CreateCell(0);
                cell.SetCellType(CellType.Formula);
                cell.CellFormula = (/*setter*/"SUM(Sheet1:Sheet3!A1:B1)");

                wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

                cell = sheet1.GetRow(1).GetCell(0);
                ClassicAssert.AreEqual(3.0, cell.NumericCellValue, 0);

                cell = sheet1.GetRow(2).GetCell(0);
                ClassicAssert.AreEqual(4.0, cell.NumericCellValue, 0);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet("test");
                IRow row = sheet.CreateRow(0);
                IRow row2 = sheet.CreateRow(1);
                ICell cellA2 = row2.CreateCell(0, CellType.Formula);
                ICell cellB1 = row.CreateCell(1, CellType.Numeric);
                cellB1.SetCellValue(10);
                IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
                cellA2.SetCellFormula("IF(B1=0,\"\",((ROW()-ROW(A$1))*12))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                ClassicAssert.AreEqual("12", Evaluate.FormatAsString());

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW()-ROW(A$1))*12),\"\")");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);

                ClassicAssert.AreEqual(Evaluate.ToString(), EvaluateN.ToString());
                ClassicAssert.AreEqual("12", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843a()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet("test");
                IRow row = sheet.CreateRow(0);
                IRow row2 = sheet.CreateRow(1);
                ICell cellA2 = row2.CreateCell(0, CellType.Formula);
                ICell cellB1 = row.CreateCell(1, CellType.Numeric);
                cellB1.SetCellValue(10);
                IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
                cellA2.SetCellFormula("IF(B1=0,\"\",((ROW(A$1))))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                ClassicAssert.AreEqual("1", Evaluate.FormatAsString());

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW(A$1))),\"\")");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);

                ClassicAssert.AreEqual(Evaluate.ToString(), EvaluateN.ToString());
                ClassicAssert.AreEqual("1", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843b()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet("test");
                IRow row = sheet.CreateRow(0);
                IRow row2 = sheet.CreateRow(1);
                ICell cellA2 = row2.CreateCell(0, CellType.Formula);
                ICell cellB1 = row.CreateCell(1, CellType.Numeric);
                cellB1.SetCellValue(10);
                IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

                cellA2.SetCellFormula("IF(B1=0,\"\",((ROW())))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                ClassicAssert.AreEqual("2", Evaluate.FormatAsString());

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW())),\"\")");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);

                ClassicAssert.AreEqual(Evaluate.ToString(), EvaluateN.ToString());
                ClassicAssert.AreEqual("2", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843c()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet("test");
                IRow row = sheet.CreateRow(0);
                IRow row2 = sheet.CreateRow(1);
                ICell cellA2 = row2.CreateCell(0, CellType.Formula);
                ICell cellB1 = row.CreateCell(1, CellType.Numeric);
                cellB1.SetCellValue(10);
                IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW())))");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);
                ClassicAssert.AreEqual("2", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843d()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet("test");
                IRow row = sheet.CreateRow(0);
                IRow row2 = sheet.CreateRow(1);
                ICell cellA2 = row2.CreateCell(0, CellType.Formula);
                ICell cellB1 = row.CreateCell(1, CellType.Numeric);
                cellB1.SetCellValue(10);
                IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

                cellA2.CellFormula = (/*setter*/"IF(NOT(B1=0),((ROW())),\"\")");
                CellValue EvaluateN = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(EvaluateN);
                ClassicAssert.AreEqual("2", EvaluateN.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug55843e()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet("test");
                IRow row = sheet.CreateRow(0);
                IRow row2 = sheet.CreateRow(1);
                ICell cellA2 = row2.CreateCell(0, CellType.Formula);
                ICell cellB1 = row.CreateCell(1, CellType.Numeric);
                cellB1.SetCellValue(10);
                IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

                cellA2.SetCellFormula("IF(B1=0,\"\",((ROW())))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                ClassicAssert.AreEqual("2", Evaluate.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }


        [Test]
        public void TestBug55843f()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet("test");
                IRow row = sheet.CreateRow(0);
                IRow row2 = sheet.CreateRow(1);
                ICell cellA2 = row2.CreateCell(0, CellType.Formula);
                ICell cellB1 = row.CreateCell(1, CellType.Numeric);
                cellB1.SetCellValue(10);
                IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

                cellA2.SetCellFormula("IF(B1=0,\"\",IF(B1=10,3,4))");
                CellValue Evaluate = formulaEvaluator.Evaluate(cellA2);
                System.Console.WriteLine(Evaluate);
                ClassicAssert.AreEqual("3", Evaluate.FormatAsString());
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug56655()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();

                SetCellFormula(sheet, 0, 0, "#VALUE!");
                SetCellFormula(sheet, 0, 1, "SUMIFS(A:A,A:A,#VALUE!)");

                wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

                ClassicAssert.AreEqual(CellType.Error, GetCell(sheet, 0, 0).CachedFormulaResultType);
                ClassicAssert.AreEqual(FormulaError.VALUE.Code, GetCell(sheet, 0, 0).ErrorCellValue);
                ClassicAssert.AreEqual(CellType.Error, GetCell(sheet, 0, 1).CachedFormulaResultType);
                ClassicAssert.AreEqual(FormulaError.VALUE.Code, GetCell(sheet, 0, 1).ErrorCellValue);
            }
            catch(Exception)
            {
                wb.Close();
            }
        }

        [Test]
        public void TestBug56655a()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();

                SetCellFormula(sheet, 0, 0, "B1*C1");
                sheet.GetRow(0).CreateCell(1).SetCellValue("A");
                SetCellFormula(sheet, 1, 0, "B1*C1");
                sheet.GetRow(1).CreateCell(1).SetCellValue("A");
                SetCellFormula(sheet, 0, 3, "SUMIFS(A:A,A:A,A2)");

                wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

                ClassicAssert.AreEqual(CellType.Error, GetCell(sheet, 0, 0).CachedFormulaResultType);
                ClassicAssert.AreEqual(FormulaError.VALUE.Code, GetCell(sheet, 0, 0).ErrorCellValue);
                ClassicAssert.AreEqual(CellType.Error, GetCell(sheet, 1, 0).CachedFormulaResultType);
                ClassicAssert.AreEqual(FormulaError.VALUE.Code, GetCell(sheet, 1, 0).ErrorCellValue);
                ClassicAssert.AreEqual(CellType.Error, GetCell(sheet, 0, 3).CachedFormulaResultType);
                ClassicAssert.AreEqual(FormulaError.VALUE.Code, GetCell(sheet, 0, 3).ErrorCellValue);
            }
            catch(Exception)
            {
                wb.Close();
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="row">0-based</param>
        /// <param name="column">0-based</param>
        private void SetCellFormula(ISheet sheet, int row, int column, string formula)
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
            cell.SetCellFormula(formula);
        }

        /// <summary>
        /// </summary>
        /// <param name="rowNo">0-based</param>
        /// <param name="column">0-based</param>
        private ICell GetCell(ISheet sheet, int rowNo, int column)
        {
            return sheet.GetRow(rowNo).GetCell(column);
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
            ClassicAssert.AreEqual("Referenced value in sheet 1" + suffix, cellValue);

            row = sheet.GetRow(1);
            cell = row.GetCell(1);
            cellValue = formatter.FormatCellValue(cell, installevaluator);
            ClassicAssert.AreEqual("Referenced value in sheet 2" + suffix, cellValue);
        }
    }
}