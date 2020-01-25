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
    using System;
    using System.Collections.Generic;
    using System.IO;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;


    public abstract class BaseTestXSSFPivotTable
    {
        private static XSSFITestDataProvider _testDataProvider = XSSFITestDataProvider.instance;
        protected XSSFWorkbook wb;
        protected XSSFPivotTable pivotTable;
        protected XSSFPivotTable offsetPivotTable;
        protected ICell offsetOuterCell;
        public abstract void SetUp();

        [TearDown]
        public void TearDown()
        {
            if (wb != null)
            {
                XSSFWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb) as XSSFWorkbook;
                wb.Close();
                wb2.Close();
            }
        }

        /**
         * Verify that when creating a row label it's  Created on the correct row
         * and the count is increased by one.
         */
        [Test]
        public void TestAddRowLabelToPivotTable()
        {
            int columnIndex = 0;

            Assert.AreEqual(0, pivotTable.GetRowLabelColumns().Count);

            pivotTable.AddRowLabel(columnIndex);
            CT_PivotTableDefinition defintion = pivotTable.GetCTPivotTableDefinition();

            Assert.AreEqual(defintion.rowFields.GetFieldArray(0).x, columnIndex);
            Assert.AreEqual(defintion.rowFields.count, 1);
            Assert.AreEqual(1, pivotTable.GetRowLabelColumns().Count);

            columnIndex = 1;
            pivotTable.AddRowLabel(columnIndex);
            Assert.AreEqual(2, pivotTable.GetRowLabelColumns().Count);

            Assert.AreEqual(0, (int)pivotTable.GetRowLabelColumns()[(0)]);
            Assert.AreEqual(1, (int)pivotTable.GetRowLabelColumns()[(1)]);
        }
        /**
         * Verify that it's not possible to create a row label outside of the referenced area.
         */
        [Test]
        public void TestAddRowLabelOutOfRangeThrowsException()
        {
            int columnIndex = 5;

            try
            {
                pivotTable.AddRowLabel(columnIndex);
            }
            catch (IndexOutOfRangeException)
            {
                return;
            }
            Assert.Fail();
        }

        /**
         * Verify that when creating one column label, no col fields are being Created.
         */
        [Test]
        public void TestAddOneColumnLabelToPivotTableDoesNotCreateColField()
        {
            int columnIndex = 0;

            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnIndex);
            CT_PivotTableDefinition defintion = pivotTable.GetCTPivotTableDefinition();

            Assert.AreEqual(defintion.colFields, null);
        }

        /**
         * Verify that it's possible to create three column labels with different DataConsolidateFunction
         */
        [Test]
        public void TestAddThreeDifferentColumnLabelsToPivotTable()
        {
            int columnOne = 0;
            int columnTwo = 1;
            int columnThree = 2;

            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnOne);
            pivotTable.AddColumnLabel(DataConsolidateFunction.MAX, columnTwo);
            pivotTable.AddColumnLabel(DataConsolidateFunction.MIN, columnThree);
            CT_PivotTableDefinition defintion = pivotTable.GetCTPivotTableDefinition();

            Assert.AreEqual(defintion.dataFields.dataField.Count, 3);
        }


        /**
         * Verify that it's possible to create three column labels with the same DataConsolidateFunction
         */
        [Test]
        public void TestAddThreeSametColumnLabelsToPivotTable()
        {
            int columnOne = 0;
            int columnTwo = 1;
            int columnThree = 2;

            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnOne);
            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnTwo);
            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnThree);
            CT_PivotTableDefinition defintion = pivotTable.GetCTPivotTableDefinition();

            Assert.AreEqual(defintion.dataFields.dataField.Count, 3);
        }

        /**
         * Verify that when creating two column labels, a col field is being Created and X is Set to -2.
         */
        [Test]
        public void TestAddTwoColumnLabelsToPivotTable()
        {
            int columnOne = 0;
            int columnTwo = 1;

            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnOne);
            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnTwo);
            CT_PivotTableDefinition defintion = pivotTable.GetCTPivotTableDefinition();

            Assert.AreEqual(defintion.colFields.field[0].x, -2);
        }

        /**
         * Verify that a data field is Created when creating a data column
         */
        [Test]
        public void TestColumnLabelCreatesDataField()
        {
            int columnIndex = 0;

            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnIndex);

            CT_PivotTableDefinition defintion = pivotTable.GetCTPivotTableDefinition();

            Assert.AreEqual(defintion.dataFields.dataField[(0)].fld, columnIndex);
            Assert.AreEqual(defintion.dataFields.dataField[(0)].subtotal,
                    (ST_DataConsolidateFunction)(DataConsolidateFunction.SUM.Value));
        }

        /**
         * Verify that it's possible to Set a custom name when creating a data column
         */
        [Test]
        public void TestColumnLabelSetCustomName()
        {
            int columnIndex = 0;

            String customName = "Custom Name";

            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnIndex, customName);

            CT_PivotTableDefinition defintion = pivotTable.GetCTPivotTableDefinition();

            Assert.AreEqual(defintion.dataFields.dataField[(0)].fld, columnIndex);
            Assert.AreEqual(defintion.dataFields.dataField[(0)].name, customName);
        }

        /**
         * Verify that it's not possible to create a column label outside of the referenced area.
         */
        [Test]
        public void TestAddColumnLabelOutOfRangeThrowsException()
        {
            int columnIndex = 5;

            try
            {
                pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, columnIndex);
            }
            catch (IndexOutOfRangeException)
            {
                return;
            }
            Assert.Fail();
        }

        /**
        * Verify when creating a data column Set to a data field, the data field with the corresponding
        * column index will be Set to true.
        */
        [Test]
        public void TestAddDataColumn()
        {
            int columnIndex = 0;
            bool isDataField = true;

            pivotTable.AddDataColumn(columnIndex, isDataField);
            CT_PivotFields pivotFields = pivotTable.GetCTPivotTableDefinition().pivotFields;
            Assert.AreEqual(pivotFields.GetPivotFieldArray(columnIndex).dataField, isDataField);
        }

        /**
         * Verify that it's not possible to create a data column outside of the referenced area.
         */
        [Test]
        public void TestAddDataColumnOutOfRangeThrowsException()
        {
            int columnIndex = 5;
            bool isDataField = true;

            try
            {
                pivotTable.AddDataColumn(columnIndex, isDataField);
            }
            catch (IndexOutOfRangeException)
            {
                return;
            }
            Assert.Fail();
        }

        /**
        * Verify that it's possible to create a new filter
        */
        [Test]
        public void TestAddReportFilter()
        {
            int columnIndex = 0;

            pivotTable.AddReportFilter(columnIndex);
            CT_PageFields fields = pivotTable.GetCTPivotTableDefinition().pageFields;
            CT_PageField field = fields.pageField[(0)];
            Assert.AreEqual(field.fld, columnIndex);
            Assert.AreEqual(field.hier, -1);
            Assert.AreEqual(fields.count, 1);
        }

        /**
        * Verify that it's not possible to create a new filter outside of the referenced area.
        */
        [Test]
        public void TestAddReportFilterOutOfRangeThrowsException()
        {
            int columnIndex = 5;
            try
            {
                pivotTable.AddReportFilter(columnIndex);
            }
            catch (IndexOutOfRangeException)
            {
                return;
            }
            Assert.Fail();
        }

        /**
         * Verify that the Pivot Table operates only within the referenced area, even when the
         * first column of the referenced area is not index 0.
         */
        [Test]
        public void TestAddDataColumnWithOffsetData()
        {
            offsetPivotTable.AddColumnLabel(DataConsolidateFunction.SUM, 1);
            Assert.AreEqual(CellType.Numeric, offsetOuterCell.CellType);

            offsetPivotTable.AddColumnLabel(DataConsolidateFunction.SUM, 0);
        }
        [Ignore("not found in poi")]
        [Test]
        public void Test58294()
        {
            //XSSFWorkbook wb = new XSSFWorkbook("C:\\temp\\test1.xlsx");
            //XSSFSheet sheet = wb.GetSheetAt(1) as XSSFSheet;
            //XSSFSheet sheet0 = wb.GetSheetAt(0) as XSSFSheet;
            //sheet0.SetActiveCell("A4");
            //XSSFPivotTable pivotTable = sheet0.CreatePivotTable(new AreaReference("A3:H6"), new CellReference("A4"), sheet);
            //pivotTable.AddRowLabel(1);
            //pivotTable.AddRowLabel(3);
            //pivotTable.AddRowLabel(5);
            //pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, 6, "Sum of days with hauls");
            //pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, 7, "Sum of days site cutoff");
            ////checkPivotTables(wb);
            //FileStream fileOut = new FileStream("c:\\temp\\test2new.xlsx", FileMode.Create, FileAccess.ReadWrite);
            //try
            //{
            //    wb.Write(fileOut);
            //}
            //finally
            //{
            //    fileOut.Close();
            //}
            //XSSFWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            ////checkPivotTables(wbBack);
            //wb.Close();
        }
        private void checkPivotTables(XSSFWorkbook wb)
        {
            IList<XSSFPivotTable> pivotTables = (wb.GetSheetAt(0) as XSSFSheet).GetPivotTables();
            Assert.IsNotNull(pivotTables);
            Assert.AreEqual(3, pivotTables.Count);
            XSSFPivotTable pivotTable = pivotTables[0];
            checkPivotTable(pivotTable);
        }
        private void checkPivotTable(XSSFPivotTable pivotTableBack)
        {
            Assert.IsNotNull(pivotTableBack.GetPivotCacheDefinition());
            Assert.IsNotNull(pivotTableBack.GetPivotCacheDefinition().GetCTPivotCacheDefinition());
            CT_CacheFields cacheFields = pivotTableBack.GetPivotCacheDefinition().GetCTPivotCacheDefinition().cacheFields;
            Assert.IsNotNull(cacheFields);
            Assert.AreEqual(8, cacheFields.SizeOfCacheFieldArray());
            Assert.AreEqual("A", cacheFields.cacheField[0].name);
        }

    }
}
