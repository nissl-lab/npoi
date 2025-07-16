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
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;


    [TestFixture]
    public class TestXSSFTable
    {

        [Test]
        public void Bug56274()
        {
            // read sample file
            XSSFWorkbook inputWorkbook = XSSFTestDataSamples.OpenSampleWorkbook("56274.xlsx");

            // read the original sheet header order
            XSSFRow row = inputWorkbook.GetSheetAt(0).GetRow(0) as XSSFRow;
            List<string> headers = new List<string>();
            foreach (ICell cell in row)
            {
                headers.Add(cell.StringCellValue);
            }

            // no SXSSF class
            // save the worksheet as-is using SXSSF
            //File outputFile = File.CreateTempFile("poi-56274", ".xlsx");
            //SXSSFWorkbook outputWorkbook = new NPOI.XSSF.streaming.SXSSFWorkbook(inputWorkbook);
            //outputWorkbook.Write(new FileOutputStream(outputFile));

            // re-read the saved file and make sure headers in the xml are in the original order
            //inputWorkbook = new NPOI.XSSF.UserModel.XSSFWorkbook(new FileStream(outputFile));
            inputWorkbook = XSSFTestDataSamples.WriteOutAndReadBack(inputWorkbook);
            CT_Table ctTable = (inputWorkbook.GetSheetAt(0) as XSSFSheet).GetTables()[0].GetCTTable();
            List<CT_TableColumn> ctTableColumnList = ctTable.tableColumns.tableColumn;

            ClassicAssert.AreEqual(headers.Count, ctTableColumnList.Count,
                    "number of headers in xml table should match number of header cells in worksheet");
            for (int i = 0; i < headers.Count; i++)
            {
                ClassicAssert.AreEqual(headers[i], ctTableColumnList[i].name,
                    "header name in xml table should match number of header cells in worksheet");
            }
            //ClassicAssert.IsTrue(outputFile.Delete());
            inputWorkbook.Close();
        }
        [Test]
        public void TestCTTableStyleInfo()
        {
            XSSFWorkbook outputWorkbook = new XSSFWorkbook();
            XSSFSheet sheet = outputWorkbook.CreateSheet() as XSSFSheet;

            //Create
            XSSFTable outputTable = sheet.CreateTable();
            outputTable.DisplayName = ("Test");
            CT_Table outputCTTable = outputTable.GetCTTable();

            //Style configurations
            CT_TableStyleInfo outputStyleInfo = outputCTTable.AddNewTableStyleInfo();
            outputStyleInfo.name = ("TableStyleLight1");
            outputStyleInfo.showColumnStripes = (false);
            outputStyleInfo.showRowStripes = (true);

            XSSFWorkbook inputWorkbook = XSSFTestDataSamples.WriteOutAndReadBack(outputWorkbook);
            List<XSSFTable> tables = (inputWorkbook.GetSheetAt(0) as XSSFSheet).GetTables();
            ClassicAssert.AreEqual(1, tables.Count, "Tables number");

            XSSFTable inputTable = tables[0];
            ClassicAssert.AreEqual(outputTable.DisplayName, inputTable.DisplayName, "Table display name");

            CT_TableStyleInfo inputStyleInfo = inputTable.GetCTTable().tableStyleInfo;
            ClassicAssert.AreEqual(outputStyleInfo.name, inputStyleInfo.name, "Style name");
            ClassicAssert.AreEqual(outputStyleInfo.showColumnStripes, inputStyleInfo.showColumnStripes, "Show column stripes");
            ClassicAssert.AreEqual(outputStyleInfo.showRowStripes, inputStyleInfo.showRowStripes, "Show row stripes");
            outputWorkbook.Close();

        }

        [Test]
        public void FindColumnIndex()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            // FIXME: use a worksheet where upper left cell of table is not A1 so that we test
            // that XSSFTable.findColumnIndex returns the column index relative to the first
            // column in the table, not the column number in the sheet
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.IsNotNull(table);
            ClassicAssert.AreEqual(0, table.FindColumnIndex("calc='#*'#"),
                "column header has special escaped characters");
            ClassicAssert.AreEqual(1, table.FindColumnIndex("Name"));
            ClassicAssert.AreEqual(2, table.FindColumnIndex("Number"));
            ClassicAssert.AreEqual(2, table.FindColumnIndex("NuMbEr"), "case insensitive");
            // findColumnIndex should return -1 if no column header name matches
            ClassicAssert.AreEqual(-1, table.FindColumnIndex(null));
            ClassicAssert.AreEqual(-1, table.FindColumnIndex(""));
            ClassicAssert.AreEqual(-1, table.FindColumnIndex("one"));
            wb.Close();
        }

        [Test]
        public void FindColumnIndexIsRelativeToTableNotSheet()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("DataTableCities.xlsx");
            XSSFTable table = wb.GetTable("SmallCity");
            // Make sure that XSSFTable.findColumnIndex returns the column index relative to the first
            // column in the table, not the column number in the sheet
            ClassicAssert.AreEqual(0, table.FindColumnIndex("City")); // column I in worksheet but 0th column in table
            ClassicAssert.AreEqual(1, table.FindColumnIndex("Latitude"));
            ClassicAssert.AreEqual(2, table.FindColumnIndex("Longitude"));
            ClassicAssert.AreEqual(3, table.FindColumnIndex("Population"));
            wb.Close();
        }

        [Test]
        public void GetSheetName()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual("Table", table.SheetName);
            wb.Close();
        }
        [Test]
        public void IsHasTotalsRow()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.IsFalse(table.TotalsRowCount > 0);
            wb.Close();
        }
        [Test]
        public void GetStartColIndex()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual(0, table.StartColIndex);
            wb.Close();
        }
        [Test]
        public void GetEndColIndex()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual(2, table.EndColIndex);
            wb.Close();
        }
        [Test]
        public void GetStartRowIndex()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual(0, table.StartRowIndex);
            wb.Close();
        }
        [Test]
        public void GetEndRowIndex()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual(6, table.EndRowIndex);
            wb.Close();
        }
        [Test]
        public void GetStartCellReference()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual(new CellReference("A1"), table.StartCellReference);
            wb.Close();
        }
        [Test]
        public void GetEndCellReference()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual(new CellReference("C7"), table.EndCellReference);
            wb.Close();
        }
        [Test]
        [Obsolete]
        public void GetNumberOfMappedColumns()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual(3, table.NumberOfMappedColumns);
            wb.Close();
        }
        [Test]
        public void GetAndSetDisplayName()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            XSSFTable table = wb.GetTable("\\_Prime.1");
            ClassicAssert.AreEqual("\\_Prime.1", table.DisplayName);
            table.DisplayName = null;
            ClassicAssert.IsNull(table.DisplayName);
            ClassicAssert.AreEqual("\\_Prime.1", table.Name); // name and display name are different
            table.DisplayName = "Display name";
            ClassicAssert.AreEqual("Display name", table.DisplayName);
            ClassicAssert.AreEqual("\\_Prime.1", table.Name); // name and display name are different
            wb.Close();
        }

        [Test]
        public void GetCellReferences()
        {
            // make sure that cached start and end cell references
            // can be synchronized with the underlying CTTable
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sh = wb.CreateSheet() as XSSFSheet;
            XSSFTable table = sh.CreateTable();
            CT_Table ctTable = table.GetCTTable();
            ctTable.@ref = "B2:E8";
            ClassicAssert.AreEqual(new CellReference("B2"), table.StartCellReference);
            ClassicAssert.AreEqual(new CellReference("E8"), table.EndCellReference);
            // At this point start and end cell reference are cached
            // and may not follow changes to the underlying CTTable
            ctTable.@ref = "C1:M3";
            ClassicAssert.AreEqual(new CellReference("B2"), table.StartCellReference);
            ClassicAssert.AreEqual(new CellReference("E8"), table.EndCellReference);
            // Force a synchronization between CTTable and XSSFTable
            // start and end cell references
            table.UpdateReferences();
            ClassicAssert.AreEqual(new CellReference("C1"), table.StartCellReference);
            ClassicAssert.AreEqual(new CellReference("M3"), table.EndCellReference);
            IOUtils.CloseQuietly(wb);
        }

        [Test]
        public void GetRowCount()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sh = wb.CreateSheet() as XSSFSheet;
            XSSFTable table = sh.CreateTable();
            CT_Table ctTable = table.GetCTTable();
            ClassicAssert.AreEqual(0, table.RowCount);
            ctTable.@ref = "B2:B2";
            // update cell references to clear the cache
            table.UpdateReferences();
            ClassicAssert.AreEqual(1, table.RowCount);
            ctTable.@ref = "B2:B12";
            // update cell references to clear the cache
            table.UpdateReferences();
            ClassicAssert.AreEqual(11, table.RowCount);
            IOUtils.CloseQuietly(wb);
        }

        [Test]
        public void FormatAsTable()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sh = wb.CreateSheet() as XSSFSheet;

            XSSFRow row = sh.CreateRow(0) as XSSFRow;
            row.CreateCell(0).SetCellValue("Col1");
            row.CreateCell(1).SetCellValue("Col2");
            row.CreateCell(2).SetCellValue("Col3");

            row = sh.CreateRow(1) as XSSFRow;
            row.CreateCell(0).SetCellValue("Value1");
            row.CreateCell(1).SetCellValue("Value2");
            row.CreateCell(2).SetCellValue("Value3");

            XSSFTable table = sh.CreateTable();
            table.CellReferences = new AreaReference(new CellReference(0, 0), new CellReference(1, 2));
            wb.Close();
        }

        [Test]
        public void GetEndCellReferenceFromSingleCellTable()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("SingleCellTable.xlsx");
            XSSFTable table = wb.GetTable("Table3");
            ClassicAssert.AreEqual(new CellReference("A2"), table.EndCellReference);
            wb.Close();
        }

        [Test]
        public void TestDifferentHeaderTypes()
        {

            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("TablesWithDifferentHeaders.xlsx");
            ClassicAssert.AreEqual(3, wb.NumberOfSheets);
            XSSFSheet s;
            XSSFTable t;

            // TODO Nicer column fetching

            s = wb.GetSheet("IntHeaders") as XSSFSheet;
            ClassicAssert.AreEqual(1, s.GetTables().Count);
            t = s.GetTables()[0];
            ClassicAssert.AreEqual("A1:B2", t.CellReferences.FormatAsString());
            ClassicAssert.AreEqual("12", t.GetCTTable().tableColumns.GetTableColumnArray(0).name);
            ClassicAssert.AreEqual("34", t.GetCTTable().tableColumns.GetTableColumnArray(1).name);

            s = wb.GetSheet("FloatHeaders") as XSSFSheet;
            ClassicAssert.AreEqual(1, s.GetTables().Count);
            t = s.GetTables()[0];
            ClassicAssert.AreEqual("A1:B2", t.CellReferences.FormatAsString());
            ClassicAssert.AreEqual("12.34", t.GetCTTable().tableColumns.GetTableColumnArray(0).name);
            ClassicAssert.AreEqual("34.56", t.GetCTTable().tableColumns.GetTableColumnArray(1).name);

            s = wb.GetSheet("NoExplicitHeaders") as XSSFSheet;
            ClassicAssert.AreEqual(1, s.GetTables().Count);
            t = s.GetTables()[0];
            ClassicAssert.AreEqual("A1:B3", t.CellReferences.FormatAsString());
            ClassicAssert.AreEqual("Column1", t.GetCTTable().tableColumns.GetTableColumnArray(0).name);
            ClassicAssert.AreEqual("Column2", t.GetCTTable().tableColumns.GetTableColumnArray(1).name);

            wb.Close();        
        }

        /// <summary>
        /// See https://stackoverflow.com/questions/44407111/apache-poi-cant-format-filled-cells-as-numeric
        /// </summary>
        [Test]
        public void TestNumericCellsInTable()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet s = wb.CreateSheet() as XSSFSheet;

            // Create some cells, some numeric, some not
            ICell c1 = s.CreateRow(0).CreateCell(0);
            ICell c2 = s.GetRow(0).CreateCell(1);
            ICell c3 = s.GetRow(0).CreateCell(2);
            ICell c4 = s.CreateRow(1).CreateCell(0);
            ICell c5 = s.GetRow(1).CreateCell(1);
            ICell c6 = s.GetRow(1).CreateCell(2);

            // Inserting values; some numeric strings, some alphabetical strings
            c1.SetCellValue(12);
            c2.SetCellValue(34.56);
            c3.SetCellValue("ABCD");
            c4.SetCellValue("AB");
            c5.SetCellValue("CD");
            c6.SetCellValue("EF");

            // Setting up the CTTable
            XSSFTable t = s.CreateTable();
            t.Name = "TableTest";
            t.DisplayName = "CT_Table_Test";
            t.AddColumn();
            t.AddColumn();
            t.AddColumn();
            t.CellReferences = (wb.GetCreationHelper().CreateAreaReference(
                    new CellReference(c1), new CellReference(c6)
            ));

            // Save and re-load
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            IOUtils.CloseQuietly(wb);
            s = wb2.GetSheetAt(0) as XSSFSheet;

            // Check
            ClassicAssert.AreEqual(1, s.GetTables().Count);
            t = s.GetTables()[0];
            ClassicAssert.AreEqual("A1", t.StartCellReference.FormatAsString());
            ClassicAssert.AreEqual("C2", t.EndCellReference.FormatAsString());

            // TODO Nicer column fetching
            ClassicAssert.AreEqual("12", t.GetCTTable().tableColumns.GetTableColumnArray(0).name);
            ClassicAssert.AreEqual("34.56", t.GetCTTable().tableColumns.GetTableColumnArray(1).name);
            ClassicAssert.AreEqual("ABCD", t.GetCTTable().tableColumns.GetTableColumnArray(2).name);
            
            IOUtils.CloseQuietly(wb2);
        }
    }
}
