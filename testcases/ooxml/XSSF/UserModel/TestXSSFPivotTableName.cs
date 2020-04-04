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

namespace TestCases.XSSF.UserModel
{

    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;

    /**
     * Test pivot tables Created by named range
     */
    [TestFixture]
    public class TestXSSFPivotTableName : BaseTestXSSFPivotTable
    {


        [SetUp]
        public override void SetUp()
        {
            wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

            IRow row1 = sheet.CreateRow(0);
            // Create a cell and Put a value in it.
            ICell cell = row1.CreateCell(0);
            cell.SetCellValue("Names");
            ICell cell2 = row1.CreateCell(1);
            cell2.SetCellValue("#");
            ICell cell7 = row1.CreateCell(2);
            cell7.SetCellValue("Data");
            ICell cell10 = row1.CreateCell(3);
            cell10.SetCellValue("Value");

            IRow row2 = sheet.CreateRow(1);
            ICell cell3 = row2.CreateCell(0);
            cell3.SetCellValue("Jan");
            ICell cell4 = row2.CreateCell(1);
            cell4.SetCellValue(10);
            ICell cell8 = row2.CreateCell(2);
            cell8.SetCellValue("Apa");
            ICell cell11 = row1.CreateCell(3);
            cell11.SetCellValue(11.11);

            IRow row3 = sheet.CreateRow(2);
            ICell cell5 = row3.CreateCell(0);
            cell5.SetCellValue("Ben");
            ICell cell6 = row3.CreateCell(1);
            cell6.SetCellValue(9);
            ICell cell9 = row3.CreateCell(2);
            cell9.SetCellValue("Bepa");
            ICell cell12 = row1.CreateCell(3);
            cell12.SetCellValue(12.12);

            XSSFName namedRange = sheet.Workbook.CreateName() as XSSFName;
            namedRange.RefersToFormula = (/*setter*/sheet.SheetName + "!" + "A1:C2");
            pivotTable = sheet.CreatePivotTable(namedRange, new CellReference("H5"));

            XSSFSheet offsetSheet = wb.CreateSheet() as XSSFSheet;

            IRow tableRow_1 = offsetSheet.CreateRow(1);
            offsetOuterCell = tableRow_1.CreateCell(1);
            offsetOuterCell.SetCellValue(-1);
            ICell tableCell_1_1 = tableRow_1.CreateCell(2);
            tableCell_1_1.SetCellValue("Row #");
            ICell tableCell_1_2 = tableRow_1.CreateCell(3);
            tableCell_1_2.SetCellValue("Exponent");
            ICell tableCell_1_3 = tableRow_1.CreateCell(4);
            tableCell_1_3.SetCellValue("10^Exponent");

            IRow tableRow_2 = offsetSheet.CreateRow(2);
            ICell tableCell_2_1 = tableRow_2.CreateCell(2);
            tableCell_2_1.SetCellValue(0);
            ICell tableCell_2_2 = tableRow_2.CreateCell(3);
            tableCell_2_2.SetCellValue(0);
            ICell tableCell_2_3 = tableRow_2.CreateCell(4);
            tableCell_2_3.SetCellValue(1);

            IRow tableRow_3 = offsetSheet.CreateRow(3);
            ICell tableCell_3_1 = tableRow_3.CreateCell(2);
            tableCell_3_1.SetCellValue(1);
            ICell tableCell_3_2 = tableRow_3.CreateCell(3);
            tableCell_3_2.SetCellValue(1);
            ICell tableCell_3_3 = tableRow_3.CreateCell(4);
            tableCell_3_3.SetCellValue(10);

            IRow tableRow_4 = offsetSheet.CreateRow(4);
            ICell tableCell_4_1 = tableRow_4.CreateCell(2);
            tableCell_4_1.SetCellValue(2);
            ICell tableCell_4_2 = tableRow_4.CreateCell(3);
            tableCell_4_2.SetCellValue(2);
            ICell tableCell_4_3 = tableRow_4.CreateCell(4);
            tableCell_4_3.SetCellValue(100);

            namedRange = sheet.Workbook.CreateName() as XSSFName;
            namedRange.RefersToFormula = (/*setter*/"C2:E4");
            namedRange.SheetIndex = (/*setter*/sheet.Workbook.GetSheetIndex(sheet));
            offsetPivotTable = offsetSheet.CreatePivotTable(namedRange, new CellReference("C6"));
        }
    }

}