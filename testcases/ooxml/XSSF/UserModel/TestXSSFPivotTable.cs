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
namespace NPOI.XSSF.UserModel
{
    using System;

using NUnit.Framework;

using NPOI.SS.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.SS.Util;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main;

    [TestFixture]
    public class TestXSSFPivotTable 
{
    private XSSFPivotTable pivotTable;
    
    
    public void SetUp(){
        IWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.CreateSheet();

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

        AreaReference source = new AreaReference("A1:C2");
        pivotTable = sheet.CreatePivotTable(source, new CellReference("H5"));
    }

    /**
     * Verify that when creating a row label it's  Created on the correct row
     * and the count is increased by one.
     */
        [Test]
    public void TestAddRowLabelToPivotTable(){
        int columnIndex = 0;

        Assert.AreEqual(0, pivotTable.RowLabelColumns.Count);
        
        pivotTable.AddRowLabel(columnIndex);
        CTPivotTableDefInition defintion = pivotTable.CTPivotTableDefInition;

        Assert.AreEqual(defintion.RowFields.GetFieldArray(0).X, columnIndex);
        Assert.AreEqual(defintion.RowFields.Count, 1);
        Assert.AreEqual(1, pivotTable.RowLabelColumns.Count);
        
        columnIndex = 1;
        pivotTable.AddRowLabel(columnIndex);
        Assert.AreEqual(2, pivotTable.RowLabelColumns.Count);
        
        Assert.AreEqual(0, (int)pivotTable.RowLabelColumns.Get(0));
        Assert.AreEqual(1, (int)pivotTable.RowLabelColumns.Get(1));
    }
    /**
     * Verify that it's not possible to create a row label outside of the referenced area.
     */
        [Test]
    public void TestAddRowLabelOutOfRangeThrowsException(){
        int columnIndex = 5;

        try {
            pivotTable.AddRowLabel(columnIndex);
        } catch(IndexOutOfBoundsException e) {
            return;
        }
        Assert.Fail();
    }

    /**
     * Verify that when creating one column label, no col fields are being Created.
     */
        [Test]
    public void TestAddOneColumnLabelToPivotTableDoesNotCreateColField(){
        int columnIndex = 0;

        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnIndex);
        CTPivotTableDefInition defintion = pivotTable.CTPivotTableDefInition;

        Assert.AreEqual(defintion.ColFields, null);
    }

    /**
     * Verify that it's possible to create three column labels with different DataConsolidateFunction
     */
        [Test]
    public void TestAddThreeDifferentColumnLabelsToPivotTable(){
        int columnOne = 0;
        int columnTwo = 1;
        int columnThree = 2;

        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnOne);
        pivotTable.AddColumnLabel(DataConsolidateFunction.MAX, columnTwo);
        pivotTable.AddColumnLabel(DataConsolidateFunction.MIN, columnThree);
        CTPivotTableDefInition defintion = pivotTable.CTPivotTableDefInition;

        Assert.AreEqual(defintion.DataFields.DataFieldList.Count, 3);
    }
    
    
    /**
     * Verify that it's possible to create three column labels with the same DataConsolidateFunction
     */
        [Test]
    public void TestAddThreeSametColumnLabelsToPivotTable(){
        int columnOne = 0;
        int columnTwo = 1;
        int columnThree = 2;

        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnOne);
        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnTwo);
        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnThree);
        CTPivotTableDefInition defintion = pivotTable.CTPivotTableDefInition;

        Assert.AreEqual(defintion.DataFields.DataFieldList.Count, 3);
    }
    
    /**
     * Verify that when creating two column labels, a col field is being Created and X is Set to -2.
     */
        [Test]
    public void TestAddTwoColumnLabelsToPivotTable(){
        int columnOne = 0;
        int columnTwo = 1;

        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnOne);
        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnTwo);
        CTPivotTableDefInition defintion = pivotTable.CTPivotTableDefInition;

        Assert.AreEqual(defintion.ColFields.GetFieldArray(0).X, -2);
    }

    /**
     * Verify that a data field is Created when creating a data column
     */
        [Test]
    public void TestColumnLabelCreatesDataField(){
        int columnIndex = 0;

        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnIndex);

        CTPivotTableDefInition defintion = pivotTable.CTPivotTableDefInition;

        Assert.AreEqual(defintion.DataFields.GetDataFieldArray(0).Fld, columnIndex);
        Assert.AreEqual(defintion.DataFields.GetDataFieldArray(0).Subtotal,
                STDataConsolidateFunction.Enum.ForInt(DataConsolidateFunction.SUM.Value));
    }
    
    /**
     * Verify that it's possible to Set a custom name when creating a data column
     */
        [Test]
    public void TestColumnLabelSetCustomName(){
        int columnIndex = 0;

        String customName = "Custom Name";
        
        pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnIndex, customName);

        CTPivotTableDefInition defintion = pivotTable.CTPivotTableDefInition;

        Assert.AreEqual(defintion.DataFields.GetDataFieldArray(0).Fld, columnIndex);
        Assert.AreEqual(defintion.DataFields.GetDataFieldArray(0).Name, customName);
    }

    /**
     * Verify that it's not possible to create a column label outside of the referenced area.
     */
        [Test]
    public void TestAddColumnLabelOutOfRangeThrowsException(){
        int columnIndex = 5;

        try {
            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnIndex);
        } catch(IndexOutOfBoundsException e) {
            return;
        }
        Assert.Fail();
    }

     /**
     * Verify when creating a data column Set to a data field, the data field with the corresponding
     * column index will be Set to true.
     */
        [Test]
    public void TestAddDataColumn(){
        int columnIndex = 0;
        bool IsDataField = true;

        pivotTable.AddDataColumn(columnIndex, isDataField);
        CTPivotFields pivotFields = pivotTable.CTPivotTableDefInition.PivotFields;
        Assert.AreEqual(pivotFields.GetPivotFieldArray(columnIndex).DataField, isDataField);
    }

    /**
     * Verify that it's not possible to create a data column outside of the referenced area.
     */
        [Test]
    public void TestAddDataColumnOutOfRangeThrowsException(){
        int columnIndex = 5;
        bool IsDataField = true;

        try {
            pivotTable.AddDataColumn(columnIndex, isDataField);
        } catch(IndexOutOfBoundsException e) {
            return;
        }
        Assert.Fail();
    }

     /**
     * Verify that it's possible to create a new filter
     */
        [Test]
    public void TestAddReportFilter(){
        int columnIndex = 0;

        pivotTable.AddReportFilter(columnIndex);
        CTPageFields fields = pivotTable.CTPivotTableDefInition.PageFields;
        CTPageField field = fields.GetPageFieldArray(0);
        Assert.AreEqual(field.Fld, columnIndex);
        Assert.AreEqual(field.Hier, -1);
        Assert.AreEqual(fields.Count, 1);
    }

     /**
     * Verify that it's not possible to create a new filter outside of the referenced area.
     */
        [Test]
    public void TestAddReportFilterOutOfRangeThrowsException(){
        int columnIndex = 5;
        try {
            pivotTable.AddReportFilter(columnIndex);
        } catch(IndexOutOfBoundsException e) {
            return;
        }
        Assert.Fail();
    }
}

