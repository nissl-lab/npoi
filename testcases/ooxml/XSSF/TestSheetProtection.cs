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
namespace NPOI.xssf
{
    using System;

using NUnit.Framework;

using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel;

    [TestFixture]
    public class TestSheetProtection 
{
	private XSSFSheet sheet;
	
	
	protected void SetUp()  {
		XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("sheetProtection_not_protected.xlsx");
		sheet = workbook.GetSheetAt(0);
	}
	
	    [Test]
    public void TestShouldReadWorkbookProtection(){
		Assert.IsFalse(sheet.IsAutoFilterLocked());
		Assert.IsFalse(sheet.IsDeleteColumnsLocked());
		Assert.IsFalse(sheet.IsDeleteRowsLocked());
		Assert.IsFalse(sheet.IsFormatCellsLocked());
		Assert.IsFalse(sheet.IsFormatColumnsLocked());
		Assert.IsFalse(sheet.IsFormatRowsLocked());
		Assert.IsFalse(sheet.IsInsertColumnsLocked());
		Assert.IsFalse(sheet.IsInsertHyperlinksLocked());
		Assert.IsFalse(sheet.IsInsertRowsLocked());
		Assert.IsFalse(sheet.IsPivotTablesLocked());
		Assert.IsFalse(sheet.IsSortLocked());
		Assert.IsFalse(sheet.IsObjectsLocked());
		Assert.IsFalse(sheet.IsScenariosLocked());
		Assert.IsFalse(sheet.IsSelectLockedCellsLocked());
		Assert.IsFalse(sheet.IsSelectUnlockedCellsLocked());
		Assert.IsFalse(sheet.IsSheetLocked());

		sheet = XSSFTestDataSamples.OpenSampleWorkbook("sheetProtection_allLocked.xlsx").GetSheetAt(0);

		Assert.IsTrue(sheet.IsAutoFilterLocked());
		Assert.IsTrue(sheet.IsDeleteColumnsLocked());
		Assert.IsTrue(sheet.IsDeleteRowsLocked());
		Assert.IsTrue(sheet.IsFormatCellsLocked());
		Assert.IsTrue(sheet.IsFormatColumnsLocked());
		Assert.IsTrue(sheet.IsFormatRowsLocked());
		Assert.IsTrue(sheet.IsInsertColumnsLocked());
		Assert.IsTrue(sheet.IsInsertHyperlinksLocked());
		Assert.IsTrue(sheet.IsInsertRowsLocked());
		Assert.IsTrue(sheet.IsPivotTablesLocked());
		Assert.IsTrue(sheet.IsSortLocked());
		Assert.IsTrue(sheet.IsObjectsLocked());
		Assert.IsTrue(sheet.IsScenariosLocked());
		Assert.IsTrue(sheet.IsSelectLockedCellsLocked());
		Assert.IsTrue(sheet.IsSelectUnlockedCellsLocked());
		Assert.IsTrue(sheet.IsSheetLocked());
	}
	
	    [Test]
    public void TestWriteAutoFilter(){
		Assert.IsFalse(sheet.IsAutoFilterLocked());
		sheet.LockAutoFilter();
		Assert.IsFalse(sheet.IsAutoFilterLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsAutoFilterLocked());
		sheet.LockAutoFilter(false);
		Assert.IsFalse(sheet.IsAutoFilterLocked());
	}
	
	    [Test]
    public void TestWriteDeleteColumns(){
		Assert.IsFalse(sheet.IsDeleteColumnsLocked());
		sheet.LockDeleteColumns();
		Assert.IsFalse(sheet.IsDeleteColumnsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsDeleteColumnsLocked());
		sheet.LockDeleteColumns(false);
		Assert.IsFalse(sheet.IsDeleteColumnsLocked());
	}
	
	    [Test]
    public void TestWriteDeleteRows(){
		Assert.IsFalse(sheet.IsDeleteRowsLocked());
		sheet.LockDeleteRows();
		Assert.IsFalse(sheet.IsDeleteRowsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsDeleteRowsLocked());
        sheet.LockDeleteRows(false);
        Assert.IsFalse(sheet.IsDeleteRowsLocked());
	}
	
	    [Test]
    public void TestWriteFormatCells(){
		Assert.IsFalse(sheet.IsFormatCellsLocked());
		sheet.LockFormatCells();
		Assert.IsFalse(sheet.IsFormatCellsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsFormatCellsLocked());
        sheet.LockFormatCells(false);
        Assert.IsFalse(sheet.IsFormatCellsLocked());
	}
	
	    [Test]
    public void TestWriteFormatColumns(){
		Assert.IsFalse(sheet.IsFormatColumnsLocked());
		sheet.LockFormatColumns();
		Assert.IsFalse(sheet.IsFormatColumnsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsFormatColumnsLocked());
        sheet.LockFormatColumns(false);
        Assert.IsFalse(sheet.IsFormatColumnsLocked());
	}
	
	    [Test]
    public void TestWriteFormatRows(){
		Assert.IsFalse(sheet.IsFormatRowsLocked());
		sheet.LockFormatRows();
		Assert.IsFalse(sheet.IsFormatRowsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsFormatRowsLocked());
        sheet.LockFormatRows(false);
        Assert.IsFalse(sheet.IsFormatRowsLocked());
	}
	
	    [Test]
    public void TestWriteInsertColumns(){
		Assert.IsFalse(sheet.IsInsertColumnsLocked());
		sheet.LockInsertColumns();
		Assert.IsFalse(sheet.IsInsertColumnsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsInsertColumnsLocked());
        sheet.LockInsertColumns(false);
        Assert.IsFalse(sheet.IsInsertColumnsLocked());
	}
	
	    [Test]
    public void TestWriteInsertHyperlinks(){
		Assert.IsFalse(sheet.IsInsertHyperlinksLocked());
		sheet.LockInsertHyperlinks();
		Assert.IsFalse(sheet.IsInsertHyperlinksLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsInsertHyperlinksLocked());
        sheet.LockInsertHyperlinks(false);
        Assert.IsFalse(sheet.IsInsertHyperlinksLocked());
	}
	
	    [Test]
    public void TestWriteInsertRows(){
		Assert.IsFalse(sheet.IsInsertRowsLocked());
		sheet.LockInsertRows();
		Assert.IsFalse(sheet.IsInsertRowsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsInsertRowsLocked());
        sheet.LockInsertRows(false);
        Assert.IsFalse(sheet.IsInsertRowsLocked());
	}
	
	    [Test]
    public void TestWritePivotTables(){
		Assert.IsFalse(sheet.IsPivotTablesLocked());
		sheet.LockPivotTables();
		Assert.IsFalse(sheet.IsPivotTablesLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsPivotTablesLocked());
        sheet.LockPivotTables(false);
        Assert.IsFalse(sheet.IsPivotTablesLocked());
	}
	
	    [Test]
    public void TestWriteSort(){
		Assert.IsFalse(sheet.IsSortLocked());
		sheet.LockSort();
		Assert.IsFalse(sheet.IsSortLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsSortLocked());
        sheet.LockSort(false);
        Assert.IsFalse(sheet.IsSortLocked());
	}
	
	    [Test]
    public void TestWriteObjects(){
		Assert.IsFalse(sheet.IsObjectsLocked());
		sheet.LockObjects();
		Assert.IsFalse(sheet.IsObjectsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsObjectsLocked());
        sheet.LockObjects(false);
        Assert.IsFalse(sheet.IsObjectsLocked());
	}
	
	    [Test]
    public void TestWriteScenarios(){
		Assert.IsFalse(sheet.IsScenariosLocked());
		sheet.LockScenarios();
		Assert.IsFalse(sheet.IsScenariosLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsScenariosLocked());
        sheet.LockScenarios(false);
        Assert.IsFalse(sheet.IsScenariosLocked());
	}
	
	    [Test]
    public void TestWriteSelectLockedCells(){
		Assert.IsFalse(sheet.IsSelectLockedCellsLocked());
		sheet.LockSelectLockedCells();
		Assert.IsFalse(sheet.IsSelectLockedCellsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsSelectLockedCellsLocked());
        sheet.LockSelectLockedCells(false);
        Assert.IsFalse(sheet.IsSelectLockedCellsLocked());
	}
	
	    [Test]
    public void TestWriteSelectUnlockedCells(){
		Assert.IsFalse(sheet.IsSelectUnlockedCellsLocked());
		sheet.LockSelectUnlockedCells();
		Assert.IsFalse(sheet.IsSelectUnlockedCellsLocked());
		sheet.EnableLocking();
		Assert.IsTrue(sheet.IsSelectUnlockedCellsLocked());
        sheet.LockSelectUnlockedCells(false);
        Assert.IsFalse(sheet.IsSelectUnlockedCellsLocked());
	}

	    [Test]
    public void TestWriteSelectEnableLocking(){
		sheet = XSSFTestDataSamples.OpenSampleWorkbook("sheetProtection_allLocked.xlsx").GetSheetAt(0);
		
		Assert.IsTrue(sheet.IsAutoFilterLocked());
		Assert.IsTrue(sheet.IsDeleteColumnsLocked());
		Assert.IsTrue(sheet.IsDeleteRowsLocked());
		Assert.IsTrue(sheet.IsFormatCellsLocked());
		Assert.IsTrue(sheet.IsFormatColumnsLocked());
		Assert.IsTrue(sheet.IsFormatRowsLocked());
		Assert.IsTrue(sheet.IsInsertColumnsLocked());
		Assert.IsTrue(sheet.IsInsertHyperlinksLocked());
		Assert.IsTrue(sheet.IsInsertRowsLocked());
		Assert.IsTrue(sheet.IsPivotTablesLocked());
		Assert.IsTrue(sheet.IsSortLocked());
		Assert.IsTrue(sheet.IsObjectsLocked());
		Assert.IsTrue(sheet.IsScenariosLocked());
		Assert.IsTrue(sheet.IsSelectLockedCellsLocked());
		Assert.IsTrue(sheet.IsSelectUnlockedCellsLocked());
		Assert.IsTrue(sheet.IsSheetLocked());
		
		sheet.DisableLocking();
		
		Assert.IsFalse(sheet.IsAutoFilterLocked());
		Assert.IsFalse(sheet.IsDeleteColumnsLocked());
		Assert.IsFalse(sheet.IsDeleteRowsLocked());
		Assert.IsFalse(sheet.IsFormatCellsLocked());
		Assert.IsFalse(sheet.IsFormatColumnsLocked());
		Assert.IsFalse(sheet.IsFormatRowsLocked());
		Assert.IsFalse(sheet.IsInsertColumnsLocked());
		Assert.IsFalse(sheet.IsInsertHyperlinksLocked());
		Assert.IsFalse(sheet.IsInsertRowsLocked());
		Assert.IsFalse(sheet.IsPivotTablesLocked());
		Assert.IsFalse(sheet.IsSortLocked());
		Assert.IsFalse(sheet.IsObjectsLocked());
		Assert.IsFalse(sheet.IsScenariosLocked());
		Assert.IsFalse(sheet.IsSelectLockedCellsLocked());
		Assert.IsFalse(sheet.IsSelectUnlockedCellsLocked());
		Assert.IsFalse(sheet.IsSheetLocked());
	}
}

