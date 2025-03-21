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
namespace TestCases
{
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    [TestFixture]
    public class TestSheetProtection
    {
        private XSSFSheet sheet;

        [SetUp]
        protected void SetUp()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("sheetProtection_not_protected.xlsx");
            sheet = (XSSFSheet)workbook.GetSheetAt(0);
        }

        [Test]
        public void TestShouldReadWorkbookProtection()
        {
            ClassicAssert.IsFalse(sheet.IsAutoFilterLocked);
            ClassicAssert.IsFalse(sheet.IsDeleteColumnsLocked);
            ClassicAssert.IsFalse(sheet.IsDeleteRowsLocked);
            ClassicAssert.IsFalse(sheet.IsFormatCellsLocked);
            ClassicAssert.IsFalse(sheet.IsFormatColumnsLocked);
            ClassicAssert.IsFalse(sheet.IsFormatRowsLocked);
            ClassicAssert.IsFalse(sheet.IsInsertColumnsLocked);
            ClassicAssert.IsFalse(sheet.IsInsertHyperlinksLocked);
            ClassicAssert.IsFalse(sheet.IsInsertRowsLocked);
            ClassicAssert.IsFalse(sheet.IsPivotTablesLocked);
            ClassicAssert.IsFalse(sheet.IsSortLocked);
            ClassicAssert.IsFalse(sheet.IsObjectsLocked);
            ClassicAssert.IsFalse(sheet.IsScenariosLocked);
            ClassicAssert.IsFalse(sheet.IsSelectLockedCellsLocked);
            ClassicAssert.IsFalse(sheet.IsSelectUnlockedCellsLocked);
            ClassicAssert.IsFalse(sheet.IsSheetLocked);

            sheet = (XSSFSheet)XSSFTestDataSamples.OpenSampleWorkbook("sheetProtection_allLocked.xlsx").GetSheetAt(0);

            ClassicAssert.IsTrue(sheet.IsAutoFilterLocked);
            ClassicAssert.IsTrue(sheet.IsDeleteColumnsLocked);
            ClassicAssert.IsTrue(sheet.IsDeleteRowsLocked);
            ClassicAssert.IsTrue(sheet.IsFormatCellsLocked);
            ClassicAssert.IsTrue(sheet.IsFormatColumnsLocked);
            ClassicAssert.IsTrue(sheet.IsFormatRowsLocked);
            ClassicAssert.IsTrue(sheet.IsInsertColumnsLocked);
            ClassicAssert.IsTrue(sheet.IsInsertHyperlinksLocked);
            ClassicAssert.IsTrue(sheet.IsInsertRowsLocked);
            ClassicAssert.IsTrue(sheet.IsPivotTablesLocked);
            ClassicAssert.IsTrue(sheet.IsSortLocked);
            ClassicAssert.IsTrue(sheet.IsObjectsLocked);
            ClassicAssert.IsTrue(sheet.IsScenariosLocked);
            ClassicAssert.IsTrue(sheet.IsSelectLockedCellsLocked);
            ClassicAssert.IsTrue(sheet.IsSelectUnlockedCellsLocked);
            ClassicAssert.IsTrue(sheet.IsSheetLocked);
        }

        [Test]
        public void TestWriteAutoFilter()
        {
            ClassicAssert.IsFalse(sheet.IsAutoFilterLocked);
            sheet.LockAutoFilter(true);
            ClassicAssert.IsFalse(sheet.IsAutoFilterLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsAutoFilterLocked);
            sheet.LockAutoFilter(false);
            ClassicAssert.IsFalse(sheet.IsAutoFilterLocked);
        }

        [Test]
        public void TestWriteDeleteColumns()
        {
            ClassicAssert.IsFalse(sheet.IsDeleteColumnsLocked);
            sheet.LockDeleteColumns(true);
            ClassicAssert.IsFalse(sheet.IsDeleteColumnsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsDeleteColumnsLocked);
            sheet.LockDeleteColumns(false);
            ClassicAssert.IsFalse(sheet.IsDeleteColumnsLocked);
        }

        [Test]
        public void TestWriteDeleteRows()
        {
            ClassicAssert.IsFalse(sheet.IsDeleteRowsLocked);
            sheet.LockDeleteRows(true);
            ClassicAssert.IsFalse(sheet.IsDeleteRowsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsDeleteRowsLocked);
            sheet.LockDeleteRows(false);
            ClassicAssert.IsFalse(sheet.IsDeleteRowsLocked);
        }

        [Test]
        public void TestWriteFormatCells()
        {
            ClassicAssert.IsFalse(sheet.IsFormatCellsLocked);
            sheet.LockFormatCells(true);
            ClassicAssert.IsFalse(sheet.IsFormatCellsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsFormatCellsLocked);
            sheet.LockFormatCells(false);
            ClassicAssert.IsFalse(sheet.IsFormatCellsLocked);
        }

        [Test]
        public void TestWriteFormatColumns()
        {
            ClassicAssert.IsFalse(sheet.IsFormatColumnsLocked);
            sheet.LockFormatColumns(true);
            ClassicAssert.IsFalse(sheet.IsFormatColumnsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsFormatColumnsLocked);
            sheet.LockFormatColumns(false);
            ClassicAssert.IsFalse(sheet.IsFormatColumnsLocked);
        }

        [Test]
        public void TestWriteFormatRows()
        {
            ClassicAssert.IsFalse(sheet.IsFormatRowsLocked);
            sheet.LockFormatRows(true);
            ClassicAssert.IsFalse(sheet.IsFormatRowsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsFormatRowsLocked);
            sheet.LockFormatRows(false);
            ClassicAssert.IsFalse(sheet.IsFormatRowsLocked);
        }

        [Test]
        public void TestWriteInsertColumns()
        {
            ClassicAssert.IsFalse(sheet.IsInsertColumnsLocked);
            sheet.LockInsertColumns(true);
            ClassicAssert.IsFalse(sheet.IsInsertColumnsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsInsertColumnsLocked);
            sheet.LockInsertColumns(false);
            ClassicAssert.IsFalse(sheet.IsInsertColumnsLocked);
        }

        [Test]
        public void TestWriteInsertHyperlinks()
        {
            ClassicAssert.IsFalse(sheet.IsInsertHyperlinksLocked);
            sheet.LockInsertHyperlinks(true);
            ClassicAssert.IsFalse(sheet.IsInsertHyperlinksLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsInsertHyperlinksLocked);
            sheet.LockInsertHyperlinks(false);
            ClassicAssert.IsFalse(sheet.IsInsertHyperlinksLocked);
        }

        [Test]
        public void TestWriteInsertRows()
        {
            ClassicAssert.IsFalse(sheet.IsInsertRowsLocked);
            sheet.LockInsertRows(true);
            ClassicAssert.IsFalse(sheet.IsInsertRowsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsInsertRowsLocked);
            sheet.LockInsertRows(false);
            ClassicAssert.IsFalse(sheet.IsInsertRowsLocked);
        }

        [Test]
        public void TestWritePivotTables()
        {
            ClassicAssert.IsFalse(sheet.IsPivotTablesLocked);
            sheet.LockPivotTables(true);
            ClassicAssert.IsFalse(sheet.IsPivotTablesLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsPivotTablesLocked);
            sheet.LockPivotTables(false);
            ClassicAssert.IsFalse(sheet.IsPivotTablesLocked);
        }

        [Test]
        public void TestWriteSort()
        {
            ClassicAssert.IsFalse(sheet.IsSortLocked);
            sheet.LockSort(true);
            ClassicAssert.IsFalse(sheet.IsSortLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsSortLocked);
            sheet.LockSort(false);
            ClassicAssert.IsFalse(sheet.IsSortLocked);
        }

        [Test]
        public void TestWriteObjects()
        {
            ClassicAssert.IsFalse(sheet.IsObjectsLocked);
            sheet.LockObjects(true);
            ClassicAssert.IsFalse(sheet.IsObjectsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsObjectsLocked);
            sheet.LockObjects(false);
            ClassicAssert.IsFalse(sheet.IsObjectsLocked);
        }

        [Test]
        public void TestWriteScenarios()
        {
            ClassicAssert.IsFalse(sheet.IsScenariosLocked);
            sheet.LockScenarios(true);
            ClassicAssert.IsFalse(sheet.IsScenariosLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsScenariosLocked);
            sheet.LockScenarios(false);
            ClassicAssert.IsFalse(sheet.IsScenariosLocked);
        }

        [Test]
        public void TestWriteSelectLockedCells()
        {
            ClassicAssert.IsFalse(sheet.IsSelectLockedCellsLocked);
            sheet.LockSelectLockedCells(true);
            ClassicAssert.IsFalse(sheet.IsSelectLockedCellsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsSelectLockedCellsLocked);
            sheet.LockSelectLockedCells(false);
            ClassicAssert.IsFalse(sheet.IsSelectLockedCellsLocked);
        }

        [Test]
        public void TestWriteSelectUnlockedCells()
        {
            ClassicAssert.IsFalse(sheet.IsSelectUnlockedCellsLocked);
            sheet.LockSelectUnlockedCells(true);
            ClassicAssert.IsFalse(sheet.IsSelectUnlockedCellsLocked);
            sheet.EnableLocking();
            ClassicAssert.IsTrue(sheet.IsSelectUnlockedCellsLocked);
            sheet.LockSelectUnlockedCells(false);
            ClassicAssert.IsFalse(sheet.IsSelectUnlockedCellsLocked);
        }

        [Test]
        public void TestWriteSelectEnableLocking()
        {
            sheet = (XSSFSheet) XSSFTestDataSamples.OpenSampleWorkbook("sheetProtection_allLocked.xlsx").GetSheetAt(0);

            ClassicAssert.IsTrue(sheet.IsAutoFilterLocked);
            ClassicAssert.IsTrue(sheet.IsDeleteColumnsLocked);
            ClassicAssert.IsTrue(sheet.IsDeleteRowsLocked);
            ClassicAssert.IsTrue(sheet.IsFormatCellsLocked);
            ClassicAssert.IsTrue(sheet.IsFormatColumnsLocked);
            ClassicAssert.IsTrue(sheet.IsFormatRowsLocked);
            ClassicAssert.IsTrue(sheet.IsInsertColumnsLocked);
            ClassicAssert.IsTrue(sheet.IsInsertHyperlinksLocked);
            ClassicAssert.IsTrue(sheet.IsInsertRowsLocked);
            ClassicAssert.IsTrue(sheet.IsPivotTablesLocked);
            ClassicAssert.IsTrue(sheet.IsSortLocked);
            ClassicAssert.IsTrue(sheet.IsObjectsLocked);
            ClassicAssert.IsTrue(sheet.IsScenariosLocked);
            ClassicAssert.IsTrue(sheet.IsSelectLockedCellsLocked);
            ClassicAssert.IsTrue(sheet.IsSelectUnlockedCellsLocked);
            ClassicAssert.IsTrue(sheet.IsSheetLocked);

            sheet.DisableLocking();

            ClassicAssert.IsFalse(sheet.IsAutoFilterLocked);
            ClassicAssert.IsFalse(sheet.IsDeleteColumnsLocked);
            ClassicAssert.IsFalse(sheet.IsDeleteRowsLocked);
            ClassicAssert.IsFalse(sheet.IsFormatCellsLocked);
            ClassicAssert.IsFalse(sheet.IsFormatColumnsLocked);
            ClassicAssert.IsFalse(sheet.IsFormatRowsLocked);
            ClassicAssert.IsFalse(sheet.IsInsertColumnsLocked);
            ClassicAssert.IsFalse(sheet.IsInsertHyperlinksLocked);
            ClassicAssert.IsFalse(sheet.IsInsertRowsLocked);
            ClassicAssert.IsFalse(sheet.IsPivotTablesLocked);
            ClassicAssert.IsFalse(sheet.IsSortLocked);
            ClassicAssert.IsFalse(sheet.IsObjectsLocked);
            ClassicAssert.IsFalse(sheet.IsScenariosLocked);
            ClassicAssert.IsFalse(sheet.IsSelectLockedCellsLocked);
            ClassicAssert.IsFalse(sheet.IsSelectUnlockedCellsLocked);
            ClassicAssert.IsFalse(sheet.IsSheetLocked);
        }
    }

}