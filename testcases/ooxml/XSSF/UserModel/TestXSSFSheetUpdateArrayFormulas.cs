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
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using TestCases.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;

namespace TestCases.XSSF.UserModel
{
    /**
     * Test array formulas in XSSF
     *
     * @author Yegor Kozlov
     * @author Josh Micich
     */
    [TestFixture]
    public class TestXSSFSheetUpdateArrayFormulas : BaseTestSheetUpdateArrayFormulas
    {

        public TestXSSFSheetUpdateArrayFormulas():base(XSSFITestDataProvider.instance)
        {
            
        }

        // Test methods common with HSSF are in superclass
        // Local methods here Test XSSF-specific details of updating array formulas
        [Test]
        public void TestXSSFSetArrayFormula_SingleCell()
        {
            ICellRange<ICell> cells;

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();

            // 1. Single-cell array formula
            String formula1 = "123";
            CellRangeAddress range = CellRangeAddress.ValueOf("C3:C3");
            cells = sheet.SetArrayFormula(formula1, range);
            Assert.AreEqual(1, cells.Size);

            // check GetFirstCell...
            XSSFCell firstCell = (XSSFCell)cells.TopLeftCell;
            Assert.AreSame(firstCell, sheet.GetFirstCellInArrayFormula(firstCell));
            //retrieve the range and check it is the same
            Assert.AreEqual(range.FormatAsString(), firstCell.ArrayFormulaRange.FormatAsString());
            ConfirmArrayFormulaCell(firstCell, "C3", formula1, "C3");

            workbook.Close();
        }
        [Test]
        public void TestXSSFSetArrayFormula_multiCell()
        {
            ICellRange<ICell> cells;

            String formula2 = "456";
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();

            CellRangeAddress range = CellRangeAddress.ValueOf("C4:C6");
            cells = sheet.SetArrayFormula(formula2, range);
            Assert.AreEqual(3, cells.Size);

            // sheet.SetArrayFormula Creates rows and cells for the designated range
            /*
             * From the spec:
             * For a multi-cell formula, the c elements for all cells except the top-left
             * cell in that range shall not have an f element;
             */
            // Check that each cell exists and that the formula text is Set correctly on the first cell
            XSSFCell firstCell = (XSSFCell)cells.TopLeftCell;
            ConfirmArrayFormulaCell(firstCell, "C4", formula2, "C4:C6");
            ConfirmArrayFormulaCell(cells.GetCell(1, 0), "C5");
            ConfirmArrayFormulaCell(cells.GetCell(2, 0), "C6");

            Assert.AreSame(firstCell, sheet.GetFirstCellInArrayFormula(firstCell));
            workbook.Close();
        }

        private static void ConfirmArrayFormulaCell(ICell c, String cellRef)
        {
            ConfirmArrayFormulaCell(c, cellRef, null, null);
        }
        private static void ConfirmArrayFormulaCell(ICell c, String cellRef, String formulaText, String arrayRangeRef)
        {
            if (c == null)
            {
                throw new AssertionException("Cell should not be null.");
            }
            CT_Cell ctCell = ((XSSFCell)c).GetCTCell();
            Assert.AreEqual(cellRef, ctCell.r);
            if (formulaText == null)
            {
                Assert.IsFalse(ctCell.IsSetF());
                Assert.IsNull(ctCell.f);
            }
            else
            {
                CT_CellFormula f = ctCell.f;
                Assert.AreEqual(arrayRangeRef, f.@ref);
                Assert.AreEqual(formulaText, f.Value);
                Assert.AreEqual(ST_CellFormulaType.array, f.t);
            }
        }

    }


}