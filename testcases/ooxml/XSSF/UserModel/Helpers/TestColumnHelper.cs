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

using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
using NPOI.XSSF.Model;
namespace NPOI.XSSF.UserModel.Helpers
{
    /**
     * Tests for {@link ColumnHelper}
     *
     */
    [TestFixture]
    public class TestColumnHelper
    {
        [Test]
        public void TestCleanColumns()
        {
            CT_Worksheet worksheet = new CT_Worksheet();

            CT_Cols cols1 = worksheet.AddNewCols();
            CT_Col col1 = cols1.AddNewCol();
            col1.min = (1);
            col1.max = (1);
            col1.width = (88);
            col1.hidden = (true);
            CT_Col col2 = cols1.AddNewCol();
            col2.min = (2);
            col2.max = (3);
            CT_Cols cols2 = worksheet.AddNewCols();
            CT_Col col4 = cols2.AddNewCol();
            col4.min = (13);
            col4.max = (16384);

            // Test cleaning cols
            Assert.AreEqual(2, worksheet.sizeOfColsArray());
            int count = countColumns(worksheet);
            Assert.AreEqual(16375, count);
            // Clean columns and Test a clean worksheet
            ColumnHelper helper = new ColumnHelper(worksheet);
            Assert.AreEqual(1, worksheet.sizeOfColsArray());
            count = countColumns(worksheet);
            Assert.AreEqual(16375, count);
            // Remember - POI column 0 == OOXML column 1
            Assert.AreEqual(88.0, helper.GetColumn(0, false).width, 0.0);
            Assert.IsTrue(helper.GetColumn(0, false).hidden);
            Assert.AreEqual(0.0, helper.GetColumn(1, false).width, 0.0);
            Assert.IsFalse(helper.GetColumn(1, false).hidden);
        }
        [Test]
        public void TestSortColumns()
        {
            //CT_Worksheet worksheet = new CT_Worksheet();
            //ColumnHelper helper = new ColumnHelper(worksheet);

            CT_Cols cols1 = new CT_Cols();
            CT_Col col1 = cols1.AddNewCol();
            col1.min = (1);
            col1.max = (1);
            col1.width = (88);
            col1.hidden = (true);
            CT_Col col2 = cols1.AddNewCol();
            col2.min = (2);
            col2.max = (3);
            CT_Col col3 = cols1.AddNewCol();
            col3.min = (13);
            col3.max = (16750);
            Assert.AreEqual(3, cols1.sizeOfColArray());
            CT_Col col4 = cols1.AddNewCol();
            col4.min = (8);
            col4.max = (11);
            Assert.AreEqual(4, cols1.sizeOfColArray());
            CT_Col col5 = cols1.AddNewCol();
            col5.min = (4);
            col5.max = (5);
            Assert.AreEqual(5, cols1.sizeOfColArray());
            CT_Col col6 = cols1.AddNewCol();
            col6.min = (8);
            col6.max = (9);
            col6.hidden = (true);
            CT_Col col7 = cols1.AddNewCol();
            col7.min = (6);
            col7.max = (8);
            col7.width = (17.0);
            CT_Col col8 = cols1.AddNewCol();
            col8.min = (25);
            col8.max = (27);
            CT_Col col9 = cols1.AddNewCol();
            col9.min = (20);
            col9.max = (30);
            Assert.AreEqual(9, cols1.sizeOfColArray());
            Assert.AreEqual(20u, cols1.GetColArray(8).min);
            Assert.AreEqual(30u, cols1.GetColArray(8).max);
            ColumnHelper.SortColumns(cols1);
            Assert.AreEqual(9, cols1.sizeOfColArray());
            Assert.AreEqual(25u, cols1.GetColArray(8).min);
            Assert.AreEqual(27u, cols1.GetColArray(8).max);
        }
        [Test]
        public void TestCloneCol()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            ColumnHelper helper = new ColumnHelper(worksheet);

            CT_Cols cols = new CT_Cols();
            CT_Col col = new CT_Col();
            col.min = (2);
            col.max = (8);
            col.hidden = (true);
            col.width = (13.4);
            CT_Col newCol = helper.CloneCol(cols, col);
            Assert.AreEqual(2u, newCol.min);
            Assert.AreEqual(8u, newCol.max);
            Assert.IsTrue(newCol.hidden);
            Assert.AreEqual(13.4, newCol.width, 0.0);
        }
        [Test]
        public void TestAddCleanColIntoCols()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            ColumnHelper helper = new ColumnHelper(worksheet);

            CT_Cols cols1 = new CT_Cols();
            CT_Col col1 = cols1.AddNewCol();
            col1.min = (1);
            col1.max = (1);
            col1.width = (88);
            col1.hidden = (true);
            CT_Col col2 = cols1.AddNewCol();
            col2.min = (2);
            col2.max = (3);
            CT_Col col3 = cols1.AddNewCol();
            col3.min = (13);
            col3.max = (16750);
            Assert.AreEqual(3, cols1.sizeOfColArray());
            CT_Col col4 = cols1.AddNewCol();
            col4.min = (8);
            col4.max = (9);
            Assert.AreEqual(4, cols1.sizeOfColArray());

            CT_Col col5 = new CT_Col();
            col5.min = (4);
            col5.max = (5);
            helper.AddCleanColIntoCols(cols1, col5);
            Assert.AreEqual(5, cols1.sizeOfColArray());

            CT_Col col6 = new CT_Col();
            col6.min = (8);
            col6.max = (11);
            col6.hidden = (true);
            helper.AddCleanColIntoCols(cols1, col6);
            Assert.AreEqual(6, cols1.sizeOfColArray());

            CT_Col col7 = new CT_Col();
            col7.min = (6);
            col7.max = (8);
            col7.width = (17.0);
            helper.AddCleanColIntoCols(cols1, col7);
            Assert.AreEqual(8, cols1.sizeOfColArray());

            CT_Col col8 = new CT_Col();
            col8.min = (20);
            col8.max = (30);
            helper.AddCleanColIntoCols(cols1, col8);
            Assert.AreEqual(10, cols1.sizeOfColArray());

            CT_Col col9 = new CT_Col();
            col9.min = (25);
            col9.max = (27);
            helper.AddCleanColIntoCols(cols1, col9);

            // TODO - assert something interesting
            Assert.AreEqual(12, cols1.col.Count);
            Assert.AreEqual(1u, cols1.GetColArray(0).min);
            Assert.AreEqual(16750u, cols1.GetColArray(11).max);
        }
        [Test]
        public void TestColumn()
        {
            CT_Worksheet worksheet = new CT_Worksheet();

            CT_Cols cols1 = worksheet.AddNewCols();
            CT_Col col1 = cols1.AddNewCol();
            col1.min = (1);
            col1.max = (1);
            col1.width = (88);
            col1.hidden = (true);
            CT_Col col2 = cols1.AddNewCol();
            col2.min = (2);
            col2.max = (3);
            CT_Cols cols2 = worksheet.AddNewCols();
            CT_Col col4 = cols2.AddNewCol();
            col4.min = (3);
            col4.max = (6);

            // Remember - POI column 0 == OOXML column 1
            ColumnHelper helper = new ColumnHelper(worksheet);
            Assert.IsNotNull(helper.GetColumn(0, false));
            Assert.IsNotNull(helper.GetColumn(1, false));
            Assert.AreEqual(88.0, helper.GetColumn(0, false).width, 0.0);
            Assert.AreEqual(0.0, helper.GetColumn(1, false).width, 0.0);
            Assert.IsTrue(helper.GetColumn(0, false).hidden);
            Assert.IsFalse(helper.GetColumn(1, false).hidden);
            Assert.IsNull(helper.GetColumn(99, false));
            Assert.IsNotNull(helper.GetColumn(5, false));
        }
        [Test]
        public void TestSetColumnAttributes()
        {
            CT_Col col = new CT_Col();
            col.width = (12);
            col.hidden = (true);
            CT_Col newCol = new CT_Col();
            Assert.AreEqual(0.0, newCol.width, 0.0);
            Assert.IsFalse(newCol.hidden);
            ColumnHelper helper = new ColumnHelper(new CT_Worksheet());
            helper.SetColumnAttributes(col, newCol);
            Assert.AreEqual(12.0, newCol.width, 0.0);
            Assert.IsTrue(newCol.hidden);
        }
        [Test]
        public void TestGetOrCreateColumn()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Sheet 1");
            ColumnHelper columnHelper = sheet.GetColumnHelper();

            // Check POI 0 based, OOXML 1 based
            CT_Col col = columnHelper.GetOrCreateColumn1Based(3, false);
            Assert.IsNotNull(col);
            Assert.IsNull(columnHelper.GetColumn(1, false));
            Assert.IsNotNull(columnHelper.GetColumn(2, false));
            Assert.IsNotNull(columnHelper.GetColumn1Based(3, false));
            Assert.IsNull(columnHelper.GetColumn(3, false));

            CT_Col col2 = columnHelper.GetOrCreateColumn1Based(30, false);
            Assert.IsNotNull(col2);
            Assert.IsNull(columnHelper.GetColumn(28, false));
            Assert.IsNotNull(columnHelper.GetColumn(29, false));
            Assert.IsNotNull(columnHelper.GetColumn1Based(30, false));
            Assert.IsNull(columnHelper.GetColumn(30, false));
        }
        [Test]
        public void TestGetSetColDefaultStyle()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            CT_Worksheet ctWorksheet = sheet.GetCTWorksheet();
            ColumnHelper columnHelper = sheet.GetColumnHelper();

            // POI column 3, OOXML column 4
            CT_Col col = columnHelper.GetOrCreateColumn1Based(4, false);

            Assert.IsNotNull(col);
            Assert.IsNotNull(columnHelper.GetColumn(3, false));
            columnHelper.SetColDefaultStyle(3, 2);
            Assert.AreEqual(2, columnHelper.GetColDefaultStyle(3));
            Assert.AreEqual(-1, columnHelper.GetColDefaultStyle(4));
            StylesTable stylesTable = workbook.GetStylesSource();
            CT_Xf cellXf = new CT_Xf();
            cellXf.fontId = (0);
            cellXf.fillId = (0);
            cellXf.borderId = (0);
            cellXf.numFmtId = (0);
            cellXf.xfId = (0);
            stylesTable.PutCellXf(cellXf);
            CT_Col col_2 = ctWorksheet.GetColsArray(0).AddNewCol();
            col_2.min = (10);
            col_2.max = (12);
            col_2.style = (1);
            Assert.AreEqual(1, columnHelper.GetColDefaultStyle(11));
            XSSFCellStyle cellStyle = new XSSFCellStyle(0, 0, stylesTable, null);
            columnHelper.SetColDefaultStyle(11, cellStyle);
            Assert.AreEqual(0u, col_2.style);
            Assert.AreEqual(1, columnHelper.GetColDefaultStyle(10));
        }

        private static int countColumns(CT_Worksheet worksheet)
        {
            int count;
            count = 0;
            for (int i = 0; i < worksheet.sizeOfColsArray(); i++)
            {
                for (int y = 0; y < worksheet.GetColsArray(i).sizeOfColArray(); y++)
                {
                    for (long k = worksheet.GetColsArray(i).GetColArray(y).min; k <= worksheet
                            .GetColsArray(i).GetColArray(y).max; k++)
                    {
                        count++;
                    }
                }
            }
            return count;
        }
    }
}

