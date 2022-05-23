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
    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using TestCases.SS;
    using NPOI.Util;

    /**
     * Common superclass for Testing usermodel API for array formulas.<br/>
     * Formula Evaluation is not Tested here.
     *
     * @author Yegor Kozlov
     * @author Josh Micich
     */
    public abstract class BaseTestSheetUpdateArrayFormulas
    {
        protected ITestDataProvider _testDataProvider;
        //public BaseTestSheetUpdateArrayFormulas()
        //{
        //    _testDataProvider = TestCases.HSSF.HSSFITestDataProvider.Instance;
        //}
        protected BaseTestSheetUpdateArrayFormulas(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }
        [Test]
        public void TestAutoCreateOtherCells()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow row1 = sheet.CreateRow(0);
            ICell cellA1 = row1.CreateCell(0);
            ICell cellB1 = row1.CreateCell(1);
            String formula = "42";
            sheet.SetArrayFormula(formula, CellRangeAddress.ValueOf("A1:B2"));

            Assert.AreEqual(formula, cellA1.CellFormula);
            Assert.AreEqual(formula, cellB1.CellFormula);
            IRow row2 = sheet.GetRow(1);
            Assert.IsNotNull(row2);
            Assert.AreEqual(formula, row2.GetCell(0).CellFormula);
            Assert.AreEqual(formula, row2.GetCell(1).CellFormula);

            workbook.Close();
        }
        /**
         *  Set Single-cell array formula
         */
        [Test]
        public void TestSetArrayFormula_SingleCell()
        {
            ICell[] cells;

            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            Assert.IsFalse(cell.IsPartOfArrayFormulaGroup);
            try
            {
                CellRangeAddress c= cell.ArrayFormulaRange;
                Assert.Fail("expected exception");
            }
            catch (InvalidOperationException e)
            {
                Assert.AreEqual("Cell A1 is not part of an array formula.", e.Message);
            }

            // row 3 does not yet exist
            Assert.IsNull(sheet.GetRow(2));
            CellRangeAddress range = new CellRangeAddress(2, 2, 2, 2);
            cells = sheet.SetArrayFormula("SUM(C11:C12*D11:D12)", range).FlattenedCells;
            Assert.AreEqual(1, cells.Length);
            // sheet.SetArrayFormula Creates rows and cells for the designated range
            Assert.IsNotNull(sheet.GetRow(2));
            cell = sheet.GetRow(2).GetCell(2);
            Assert.IsNotNull(cell);

            Assert.IsTrue(cell.IsPartOfArrayFormulaGroup);
            //retrieve the range and check it is the same
            Assert.AreEqual(range.FormatAsString(), cell.ArrayFormulaRange.FormatAsString());
            //check the formula
            Assert.AreEqual("SUM(C11:C12*D11:D12)", cell.CellFormula);
            workbook.Close();
        }

        /**
         * Set multi-cell array formula
         */
        [Test]
        public void TestSetArrayFormula_multiCell()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            // multi-cell formula
            // rows 3-5 don't exist yet
            Assert.IsNull(sheet.GetRow(3));
            Assert.IsNull(sheet.GetRow(4));
            Assert.IsNull(sheet.GetRow(5));

            CellRangeAddress range = CellRangeAddress.ValueOf("C4:C6");
            ICell[] cells = sheet.SetArrayFormula("SUM(A1:A3*B1:B3)", range).FlattenedCells;
            Assert.AreEqual(3, cells.Length);

            // sheet.SetArrayFormula Creates rows and cells for the designated range
            Assert.AreSame(cells[0], sheet.GetRow(3).GetCell(2));
            Assert.AreSame(cells[1], sheet.GetRow(4).GetCell(2));
            Assert.AreSame(cells[2], sheet.GetRow(5).GetCell(2));

            foreach (ICell acell in cells)
            {
                Assert.IsTrue(acell.IsPartOfArrayFormulaGroup);
                Assert.AreEqual(CellType.Formula, acell.CellType);
                Assert.AreEqual("SUM(A1:A3*B1:B3)", acell.CellFormula);
                //retrieve the range and check it is the same
                Assert.AreEqual(range.FormatAsString(), acell.ArrayFormulaRange.FormatAsString());
            }
            workbook.Close();
        }

        /**
         * Passing an incorrect formula to sheet.SetArrayFormula
         *  should throw FormulaParseException
         */
        [Test]
        public void TestSetArrayFormula_incorrectFormula()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            try
            {
                sheet.SetArrayFormula("incorrect-formula(C11_C12*D11_D12)",
                        new CellRangeAddress(10, 10, 10, 10));
                Assert.Fail("expected exception");
            }
            catch (FormulaParseException)
            {
                //expected exception
            }
            workbook.Close();
        }

        /**
         * Calls of cell.GetArrayFormulaRange and sheet.RemoveArrayFormula
         * on a not-array-formula cell throw InvalidOperationException
         */
        [Test]
        public void TestArrayFormulas_illegalCalls()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            ICell cell = sheet.CreateRow(0).CreateCell(0);
            Assert.IsFalse(cell.IsPartOfArrayFormulaGroup);
            try
            {
                CellRangeAddress c= cell.ArrayFormulaRange;
                Assert.Fail("expected exception");
            }
            catch (InvalidOperationException e)
            {
                Assert.AreEqual("Cell A1 is not part of an array formula.", e.Message);
            }

            try
            {
                sheet.RemoveArrayFormula(cell);
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Cell A1 is not part of an array formula.", e.Message);
            }

            workbook.Close();
        }

        /**
         * create and remove array formulas
         */
        [Test]
        public void TestRemoveArrayFormula()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            CellRangeAddress range = new CellRangeAddress(3, 5, 2, 2);
            Assert.AreEqual("C4:C6", range.FormatAsString());
            ICellRange<ICell> cr = sheet.SetArrayFormula("SUM(A1:A3*B1:B3)", range);
            Assert.AreEqual(3, cr.Size);

            // remove the formula cells in C4:C6
            ICellRange<ICell> dcells = sheet.RemoveArrayFormula(cr.TopLeftCell);
            // RemoveArrayFormula should return the same cells as SetArrayFormula
            Assert.IsTrue(Arrays.Equals(cr.FlattenedCells, dcells.FlattenedCells));

            foreach (ICell acell in cr)
            {
                Assert.IsFalse(acell.IsPartOfArrayFormulaGroup);
                Assert.AreEqual(CellType.Blank, acell.CellType);
            }

            // cells C4:C6 are not included in array formula,
            // invocation of sheet.RemoveArrayFormula on any of them throws ArgumentException
            foreach (ICell acell in cr)
            {
                try
                {
                    sheet.RemoveArrayFormula(acell);
                    Assert.Fail("expected exception");
                }
                catch (ArgumentException e)
                {
                    String ref1 = new CellReference(acell).FormatAsString();
                    Assert.AreEqual("Cell " + ref1 + " is not part of an array formula.", e.Message);
                }
            }

            workbook.Close();
        }

        /**
         * Test that when Reading a workbook from input stream, array formulas are recognized
         */
        [Test]
        public void TestReadArrayFormula()
        {
            ICell[] cells;

            IWorkbook workbook1 = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = workbook1.CreateSheet();
            cells = sheet1.SetArrayFormula("SUM(A1:A3*B1:B3)", CellRangeAddress.ValueOf("C4:C6")).FlattenedCells;
            Assert.AreEqual(3, cells.Length);

            cells = sheet1.SetArrayFormula("MAX(A1:A3*B1:B3)", CellRangeAddress.ValueOf("A4:A6")).FlattenedCells;
            Assert.AreEqual(3, cells.Length);

            ISheet sheet2 = workbook1.CreateSheet();
            cells = sheet2.SetArrayFormula("MIN(A1:A3*B1:B3)", CellRangeAddress.ValueOf("D2:D4")).FlattenedCells;
            Assert.AreEqual(3, cells.Length);

            IWorkbook workbook2 = _testDataProvider.WriteOutAndReadBack(workbook1);
            workbook1.Close();

            sheet1 = workbook2.GetSheetAt(0);
            for (int rownum = 3; rownum <= 5; rownum++)
            {
                ICell cell1 = sheet1.GetRow(rownum).GetCell(2);
                Assert.IsTrue(cell1.IsPartOfArrayFormulaGroup);

                ICell cell2 = sheet1.GetRow(rownum).GetCell(0);
                Assert.IsTrue(cell2.IsPartOfArrayFormulaGroup);
            }

            sheet2 = workbook2.GetSheetAt(1);
            for (int rownum = 1; rownum <= 3; rownum++)
            {
                ICell cell1 = sheet2.GetRow(rownum).GetCell(3);
                Assert.IsTrue(cell1.IsPartOfArrayFormulaGroup);
            }
            workbook2.Close();
        }

        /**
         * Test that we can Set pre-calculated formula result for array formulas
         */
        [Test]
        public void TestModifyArrayCells_setFormulaResult()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            //Single-cell array formula
            ICellRange<ICell> srange =
                    sheet.SetArrayFormula("SUM(A4:A6,B4:B6)", CellRangeAddress.ValueOf("B5"));
            ICell scell = srange.TopLeftCell;
            Assert.AreEqual(CellType.Formula, scell.CellType);
            Assert.AreEqual(0.0, scell.NumericCellValue, 0);
            scell.SetCellValue(1.1);
            Assert.AreEqual(1.1, scell.NumericCellValue, 0);

            //multi-cell array formula
            ICellRange<ICell> mrange =
                    sheet.SetArrayFormula("A1:A3*B1:B3", CellRangeAddress.ValueOf("C1:C3"));
            foreach (ICell mcell in mrange)
            {
                Assert.AreEqual(CellType.Formula, mcell.CellType);
                Assert.AreEqual(0.0, mcell.NumericCellValue, 0);
                double fmlaResult = 1.2;
                mcell.SetCellValue(fmlaResult);
                Assert.AreEqual(fmlaResult, mcell.NumericCellValue, 0);
            }
            workbook.Close();
        }
        [Test]
        public void TestModifyArrayCells_setCellType()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            // Single-cell array formulas behave just like normal cells -
            // changing cell type Removes the array formula and associated cached result
            ICellRange<ICell> srange =
                    sheet.SetArrayFormula("SUM(A4:A6,B4:B6)", CellRangeAddress.ValueOf("B5"));
            ICell scell = srange.TopLeftCell;
            Assert.AreEqual(CellType.Formula, scell.CellType);
            Assert.AreEqual(0.0, scell.NumericCellValue, 0);
            scell.SetCellType(CellType.String);
            Assert.AreEqual(CellType.String, scell.CellType);
            scell.SetCellValue("string cell");
            Assert.AreEqual("string cell", scell.StringCellValue);

            //once you create a multi-cell array formula, you cannot change the type of its cells
            ICellRange<ICell> mrange =
                    sheet.SetArrayFormula("A1:A3*B1:B3", CellRangeAddress.ValueOf("C1:C3"));
            foreach (ICell mcell in mrange)
            {
                try
                {
                    Assert.AreEqual(CellType.Formula, mcell.CellType);
                    mcell.SetCellType(CellType.Numeric);
                    Assert.Fail("expected exception");
                }
                catch (InvalidOperationException e)
                {
                    CellReference ref1 = new CellReference(mcell);
                    String msg = "Cell " + ref1.FormatAsString() + " is part of a multi-cell array formula. You cannot change part of an array.";
                    Assert.AreEqual(msg, e.Message);
                }
                // a failed invocation of Cell.SetCellType leaves the cell
                // in the state that it was in prior to the invocation
                Assert.AreEqual(CellType.Formula, mcell.CellType);
                Assert.IsTrue(mcell.IsPartOfArrayFormulaGroup);
            }
            workbook.Close();
        }
        [Test]
        public void TestModifyArrayCells_setCellFormula()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            ICellRange<ICell> srange =
                    sheet.SetArrayFormula("SUM(A4:A6,B4:B6)", CellRangeAddress.ValueOf("B5"));
            ICell scell = srange.TopLeftCell;
            Assert.AreEqual("SUM(A4:A6,B4:B6)", scell.CellFormula);
            Assert.AreEqual(CellType.Formula, scell.CellType);
            Assert.IsTrue(scell.IsPartOfArrayFormulaGroup);
            scell.CellFormula = (/*setter*/"SUM(A4,A6)");
            //we are now a normal formula cell
            Assert.AreEqual("SUM(A4,A6)", scell.CellFormula);
            Assert.IsFalse(scell.IsPartOfArrayFormulaGroup);
            Assert.AreEqual(CellType.Formula, scell.CellType);
            //check that Setting formula result works
            Assert.AreEqual(0.0, scell.NumericCellValue, 0);
            scell.SetCellValue(33.0);
            Assert.AreEqual(33.0, scell.NumericCellValue, 0);

            //multi-cell array formula
            ICellRange<ICell> mrange =
                    sheet.SetArrayFormula("A1:A3*B1:B3", CellRangeAddress.ValueOf("C1:C3"));
            foreach (ICell mcell in mrange)
            {
                //we cannot Set individual formulas for cells included in an array formula
                try
                {
                    Assert.AreEqual("A1:A3*B1:B3", mcell.CellFormula);
                    mcell.CellFormula = (/*setter*/"A1+A2");
                    Assert.Fail("expected exception");
                }
                catch (InvalidOperationException e)
                {
                    CellReference ref1 = new CellReference(mcell);
                    String msg = "Cell " + ref1.FormatAsString() + " is part of a multi-cell array formula. You cannot change part of an array.";
                    Assert.AreEqual(msg, e.Message);
                }
                // a failed invocation of Cell.SetCellFormula leaves the cell
                // in the state that it was in prior to the invocation
                Assert.AreEqual("A1:A3*B1:B3", mcell.CellFormula);
                Assert.IsTrue(mcell.IsPartOfArrayFormulaGroup);
            }

            workbook.Close();
        }
        [Test]
        public void TestModifyArrayCells_RemoveCell()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            //Single-cell array formulas behave just like normal cells
            CellRangeAddress cra = CellRangeAddress.ValueOf("B5");
            ICellRange<ICell> srange =
                    sheet.SetArrayFormula("SUM(A4:A6,B4:B6)", cra);
            ICell scell = srange.TopLeftCell;

            IRow srow = sheet.GetRow(cra.FirstRow);
            Assert.AreSame(srow, scell.Row);
            srow.RemoveCell(scell);
            Assert.IsNull(srow.GetCell(cra.FirstColumn));

            //re-create the Removed cell
            scell = srow.CreateCell(cra.FirstColumn);
            Assert.AreEqual(CellType.Blank, scell.CellType);
            Assert.IsFalse(scell.IsPartOfArrayFormulaGroup);

            //we cannot remove cells included in a multi-cell array formula
            ICellRange<ICell> mrange =
                    sheet.SetArrayFormula("A1:A3*B1:B3", CellRangeAddress.ValueOf("C1:C3"));
            foreach (ICell mcell in mrange)
            {
                int columnIndex = mcell.ColumnIndex;
                IRow mrow = mcell.Row;
                try
                {
                    mrow.RemoveCell(mcell);
                    Assert.Fail("expected exception");
                }
                catch (InvalidOperationException e)
                {
                    CellReference ref1 = new CellReference(mcell);
                    String msg = "Cell " + ref1.FormatAsString() + " is part of a multi-cell array formula. You cannot change part of an array.";
                    Assert.AreEqual(msg, e.Message);
                }
                // a failed invocation of Row.RemoveCell leaves the row
                // in the state that it was in prior to the invocation
                Assert.AreSame(mcell, mrow.GetCell(columnIndex));
                Assert.IsTrue(mcell.IsPartOfArrayFormulaGroup);
                Assert.AreEqual(CellType.Formula, mcell.CellType);
            }

            workbook.Close();
        }
        [Test]
        public void TestModifyArrayCells_RemoveRow()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            //Single-cell array formulas behave just like normal cells
            CellRangeAddress cra = CellRangeAddress.ValueOf("B5");
            ICellRange<ICell> srange =
                    sheet.SetArrayFormula("SUM(A4:A6,B4:B6)", cra);
            ICell scell = srange.TopLeftCell;
            Assert.AreEqual(CellType.Formula, scell.CellType);

            IRow srow = scell.Row;
            Assert.AreSame(srow, sheet.GetRow(cra.FirstRow));
            sheet.RemoveRow(srow);
            Assert.IsNull(sheet.GetRow(cra.FirstRow));

            //re-create the Removed row and cell
            scell = sheet.CreateRow(cra.FirstRow).CreateCell(cra.FirstColumn);
            Assert.AreEqual(CellType.Blank, scell.CellType);
            Assert.IsFalse(scell.IsPartOfArrayFormulaGroup);

            //we cannot remove rows with cells included in a multi-cell array formula
            ICellRange<ICell> mrange =
                    sheet.SetArrayFormula("A1:A3*B1:B3", CellRangeAddress.ValueOf("C1:C3"));
            foreach (ICell mcell in mrange)
            {
                int columnIndex = mcell.ColumnIndex;
                IRow mrow = mcell.Row;
                try
                {
                    sheet.RemoveRow(mrow);
                    Assert.Fail("expected exception");
                }
                catch (InvalidOperationException)
                {
                    // String msg = "Row[rownum=" + mrow.RowNum + "] Contains cell(s) included in a multi-cell array formula. You cannot change part of an array.";
                    //Assert.AreEqual(msg, e.Message);
                }
                // a failed invocation of Row.RemoveCell leaves the row
                // in the state that it was in prior to the invocation
                Assert.AreSame(mrow, sheet.GetRow(mrow.RowNum));
                Assert.AreSame(mcell, mrow.GetCell(columnIndex));
                Assert.IsTrue(mcell.IsPartOfArrayFormulaGroup);
                Assert.AreEqual(CellType.Formula, mcell.CellType);
            }

            workbook.Close();
        }
        [Test]
        public void TestModifyArrayCells_mergeCellsSingle()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            Assert.AreEqual(0, sheet.NumMergedRegions);

            //Single-cell array formulas behave just like normal cells
            ICellRange<ICell> srange =
                    sheet.SetArrayFormula("SUM(A4:A6,B4:B6)", CellRangeAddress.ValueOf("B5"));
            ICell scell = srange.TopLeftCell;
            sheet.AddMergedRegion(CellRangeAddress.ValueOf("B5:C6"));
            //we are still an array formula
            Assert.AreEqual(CellType.Formula, scell.CellType);
            Assert.IsTrue(scell.IsPartOfArrayFormulaGroup);
            Assert.AreEqual(1, sheet.NumMergedRegions);
            workbook.Close();
        }

        [Test]
        public void testModifyArrayCells_mergeCellsMulti()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();
            int expectedNumMergedRegions = 0;
            Assert.AreEqual(expectedNumMergedRegions, sheet.NumMergedRegions);
            // we cannot merge cells included in an array formula
            sheet.SetArrayFormula("A1:A4*B1:B4", CellRangeAddress.ValueOf("C2:F5"));
            foreach (String ref1 in Arrays.AsList(
                    "C2:F5", // identity
                    "D3:E4", "B1:G6", // contains
                    "B1:C2", "F1:G2", "F5:G6", "B5:C6", // 1x1 corner intersection
                    "B1:C6", "B1:G2", "F1:G6", "B5:G6", // 1-row/1-column intersection
                    "B1:D3", "E1:G3", "E4:G6", "B4:D6", // 2x2 corner intersection
                    "B1:D6", "B1:G3", "E1:G6", "B4:G6"  // 2-row/2-column intersection
            )) {
                CellRangeAddress cra = CellRangeAddress.ValueOf(ref1);
                try
                {
                    sheet.AddMergedRegion(cra);
                    Assert.Fail("expected exception with ref " + ref1);
                }
                catch (InvalidOperationException e)
                {
                    String msg = "The range " + cra.FormatAsString() + " intersects with a multi-cell array formula. You cannot merge cells of an array.";
                    Assert.AreEqual(msg, e.Message);
                }
            }
            //the number of merged regions remains the same
            Assert.AreEqual(expectedNumMergedRegions, sheet.NumMergedRegions);

            // we can merge non-intersecting cells
            foreach (String ref1 in Arrays.AsList(
                    "C1:F1", //above
                    "G2:G5", //right
                    "C6:F6",  //bottom
                    "B2:B5", "H7:J9")) {
                CellRangeAddress cra = CellRangeAddress.ValueOf(ref1);
                try
                {
                    sheet.AddMergedRegion(cra);
                    expectedNumMergedRegions++;
                    Assert.AreEqual(expectedNumMergedRegions, sheet.NumMergedRegions);
                }
                catch (InvalidOperationException e)
                {
                    Assert.Fail("did not expect exception with ref: " + ref1 +"\n" + e.Message);
                }
            }

            workbook.Close();
        }

        [Test]
        public void TestModifyArrayCells_ShiftRows()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            //Single-cell array formulas behave just like normal cells - we can change the cell type
            ICellRange<ICell> srange =
                    sheet.SetArrayFormula("SUM(A4:A6,B4:B6)", CellRangeAddress.ValueOf("B5"));
            ICell scell = srange.TopLeftCell;
            Assert.AreEqual("SUM(A4:A6,B4:B6)", scell.CellFormula);
            sheet.ShiftRows(0, 0, 1);
            sheet.ShiftRows(0, 1, 1);

            //we cannot Set individual formulas for cells included in an array formula
            ICellRange<ICell> mrange =
                    sheet.SetArrayFormula("A1:A3*B1:B3", CellRangeAddress.ValueOf("C1:C3"));

            try
            {
                sheet.ShiftRows(0, 0, 1);
                Assert.Fail("expected exception");
            }
            catch (InvalidOperationException e)
            {
                String msg = "Row[rownum=0] contains cell(s) included in a multi-cell array formula. You cannot change part of an array.";
                Assert.AreEqual(msg, e.Message);
            }
            /*
             TODO: enable Shifting the whole array

            sheet.ShiftRows(0, 2, 1);
            //the array C1:C3 is now C2:C4
            CellRangeAddress cra = CellRangeAddress.ValueOf("C2:C4");
            foreach(Cell mcell in mrange){
                //TODO define Equals and hashcode for CellRangeAddress
                Assert.AreEqual(cra.FormatAsString(), mcell.ArrayFormulaRange.formatAsString());
                Assert.AreEqual("A2:A4*B2:B4", mcell.CellFormula);
                Assert.IsTrue(mcell.IsPartOfArrayFormulaGroup);
                Assert.AreEqual(CellType.Formula, mcell.CellType);
            }

            */

            workbook.Close();
        }

        [Ignore("also ignored in poi")]
        [Test]
        public void ShouldNotBeAbleToCreateArrayFormulaOnPreexistingMergedRegion()
        {
            /*
             *  m  = merged region
             *  f  = array formula
             *  fm = cell belongs to a merged region and an array formula (illegal, that's what this tests for)
             *  
             *   A  B  C
             * 1    f  f
             * 2    fm fm
             * 3    f  f
             */
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            ISheet sheet = workbook.CreateSheet();

            CellRangeAddress mergedRegion = CellRangeAddress.ValueOf("B2:C2");
            sheet.AddMergedRegion(mergedRegion);
            CellRangeAddress arrayFormula = CellRangeAddress.ValueOf("C1:C3");
            Assume.That(mergedRegion.Intersects(arrayFormula));
            Assume.That(arrayFormula.Intersects(mergedRegion));
            try
            {
                sheet.SetArrayFormula("SUM(A1:A3)",  arrayFormula);
                Assert.Fail("expected exception: should not be able to create an array formula that intersects with a merged region");
            }
            catch (InvalidOperationException)
            {
                // expected
            }

            workbook.Close();
        }
    }

}