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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace TestCases.XSSF.Streaming
{
    [TestFixture]
    class SXSSFWorkbookTests
    {
        private SXSSFWorkbook _objectToTest { get; set; }

        [Test]
        public void
            CallingEmptyConstructorShouldInstanstiateNewXssfWorkbookDefaultRowAccessWindowSizeCompressTempFilesAsFalseAndUseSharedStringsTableFalse
            ()
        {
            _objectToTest = new SXSSFWorkbook();
            Assert.AreEqual(100, _objectToTest.RandomAccessWindowSize);
            Assert.NotNull(_objectToTest.XssfWorkbook);
        }

        [Test]
        public void
            CallingConstructorWithNullWorkbookShouldInstanstiateNewXssfWorkbookDefaultRowAccessWindowSizeCompressTempFilesAsFalseAndUseSharedStringsTableFalse
            ()
        {
            _objectToTest = new SXSSFWorkbook(null);
            Assert.AreEqual(100, _objectToTest.RandomAccessWindowSize);
            Assert.NotNull(_objectToTest.XssfWorkbook);
        }

        [Test]
        public void
    CallingConstructorWithExistingWorkbookShouldInstanstiateNewXssfWorkbookDefaultRowAccessWindowSizeCompressTempFilesAsFalseAndUseSharedStringsTableFalse
    ()
        {
            var wb = new XSSFWorkbook();
            var name = wb.CreateName();

            name.NameName = "test";
            var sheet = wb.CreateSheet("test1");

            _objectToTest = new SXSSFWorkbook(wb);
            Assert.AreEqual(100, _objectToTest.RandomAccessWindowSize);
            Assert.AreEqual("test", _objectToTest.XssfWorkbook.GetName("test").NameName);
            Assert.AreEqual(1, _objectToTest.NumberOfSheets);
        }

        [Test]
        public void IfCompressTmpFilesIsSetToTrueShouldReturnGZIPSheetDataWriter()
        {
            _objectToTest = new SXSSFWorkbook(null, 100, true);
            var result = _objectToTest.CreateSheetDataWriter();

            Assert.IsTrue(result is GZIPSheetDataWriter);

        }

        [Test]
        public void IfCompressTmpFilesIsSetToFalseShouldReturnSheetDataWriter()
        {
            _objectToTest = new SXSSFWorkbook();
            var result = _objectToTest.CreateSheetDataWriter();

            Assert.IsTrue(result is SheetDataWriter);

        }

        [Test]
        public void IfSettingSheetOrderShouldSetSheetOrderOfXssfWorkbook()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("test1");
            _objectToTest.CreateSheet("test2");

            _objectToTest.SetSheetOrder("test2", 0);

            Assert.AreEqual("test2", _objectToTest.GetSheetName(0));
            Assert.AreEqual("test1", _objectToTest.GetSheetName(1));
        }

        [Test]
        public void IfSettingSelectedTabShouldSetSelectedTabOfXssfWorkbook()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("test1");
            _objectToTest.CreateSheet("test2");

            _objectToTest.SetSelectedTab(0);

            Assert.IsTrue(_objectToTest.GetSheetAt(0).IsSelected);

        }

        [Test]
        public void IfSheetNameByIndexShouldGetSheetNameFromXssfWorkbook()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("test1");
            _objectToTest.CreateSheet("test2");

            _objectToTest.SetSelectedTab(0);

            Assert.IsTrue(_objectToTest.GetSheetAt(0).IsSelected);

        }

        [Test]
        public void IfSettingSheetNameShouldChangeTheSheetNameAtTheSpecifiedIndex()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("test1");
            _objectToTest.SetSheetName(0, "renamed");

            Assert.AreEqual("renamed", _objectToTest.GetSheetAt(0).SheetName);

        }

        [Test]
        public void IfRequestingTheSheetIndexBySheetNameShouldReturnTheIndexOfTheXssfSheet()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("test0");
            _objectToTest.CreateSheet("test1");
            _objectToTest.CreateSheet("test2");

            var first = _objectToTest.GetSheetIndex("test0");
            var second = _objectToTest.GetSheetIndex("test1");
            var third = _objectToTest.GetSheetIndex("test2");
            Assert.AreEqual(0, first);
            Assert.AreEqual(1, second);
            Assert.AreEqual(2, third);
        }

        [Test]
        public void IfGivenASheetOfAWorkbookShouldGetTheIndexIfTheSheetExists()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheet = _objectToTest.CreateSheet("test");

            var index = _objectToTest.GetSheetIndex(sheet);

            Assert.AreEqual(0, index);
        }

        [Test]
        public void IfCreatingASheetShouldCreateASheetInTheXssfWorkbookWithTheGivenName()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheet = _objectToTest.CreateSheet("test");

            Assert.NotNull(sheet);
            Assert.AreEqual("test", sheet.SheetName);
        }

        [Test]
        public void IfCreatingASheetShouldCreateASheetInTheXssfWorkbookWithDefaultName()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheet = _objectToTest.CreateSheet();
            var sheet2 = _objectToTest.CreateSheet();

            Assert.NotNull(sheet);
            Assert.AreEqual("Sheet0", sheet.SheetName);
            Assert.AreEqual("Sheet1", sheet2.SheetName);
        }

        [Test]
        public void IfGivenTheNameOfAnExistingSheetShouldReturnTheSheet()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("1");
            _objectToTest.CreateSheet("2");

            var sheet1 = _objectToTest.GetSheet("1");
            var sheet2 = _objectToTest.GetSheet("2");

            Assert.AreEqual("1", sheet1.SheetName);
            Assert.AreEqual("2", sheet2.SheetName);

        }


        [Test]
        public void IfGivenTheIndexOfAnExistingSheetShouldReturnTheSheet()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("1");
            _objectToTest.CreateSheet("2");

            var sheet1 = _objectToTest.GetSheetAt(0);
            var sheet2 = _objectToTest.GetSheetAt(1);

            Assert.AreEqual("1", sheet1.SheetName);
            Assert.AreEqual("2", sheet2.SheetName);

        }

        [Test]
        public void IfGivenThePositionOfAnExistingSheetShouldRemoveThatSheet()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateSheet("1");
            _objectToTest.CreateSheet("2");

            _objectToTest.RemoveSheetAt(0);
            var sheet = _objectToTest.GetSheetAt(0);
            Assert.IsTrue(_objectToTest.NumberOfSheets == 1);
            Assert.AreEqual("2", sheet.SheetName);

        }

        [Test]
        public void IfAFontIsCreatedItShouldBeReturnedAndAddedToTheExistingWorkbook()
        {
            _objectToTest = new SXSSFWorkbook();
            var font = _objectToTest.CreateFont();

            Assert.NotNull(font);
        }

        [Test]
        public void IfGivenZeroBasedIndexShouldReturnExistingFont()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateFont();

            var font = _objectToTest.GetFontAt(0);
            Assert.NotNull(font);
        }

        [Test]
        public void IfACellStyleIsCreatedItShouldBeReturnedAndAddedToTheExistingWorkbook()
        {
            _objectToTest = new SXSSFWorkbook();
            var cellStyle = _objectToTest.CreateCellStyle();

            Assert.NotNull(cellStyle);
        }

        [Test]
        public void IfGivenZeroBasedIndexShouldReturnExistingCellStyle()
        {
            _objectToTest = new SXSSFWorkbook();
            _objectToTest.CreateFont();

            var cellStyle = _objectToTest.GetCellStyleAt(0);
            Assert.NotNull(cellStyle);
        }

        [Test]
        public void IfWriting10x10CellsShouldWriteNumericValuesForCells()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheets = 1;
            var rows = 10;
            var cols = 10;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Numeric);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "numericTest.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        public void IfOpeningExistingWorkbookShouldWriteAllPreviouslyExistingColumns()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheets = 1;
            var rows = 10;
            var cols = 10;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Numeric);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "numericTest.xlsx");
            var reSavePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "numericTest2.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));

            var xssfWorkbook = new XSSFWorkbook(savePath);

            var result = new SXSSFWorkbook(xssfWorkbook);

            WriteFile(reSavePath, result);
            xssfWorkbook.Close();

            File.Delete(reSavePath);
            File.Delete(savePath);
        }

        [Test]
        public void IfWriting10x10CellsShouldWriteStringValuesForCells()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheets = 1;
            var rows = 10;
            var cols = 10;
            AddCells(_objectToTest, sheets, rows, cols, CellType.String);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "plainStringTest.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        public void IfWriting10x10CellsShouldWriteBooleanValuesForCells()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheets = 1;
            var rows = 10;
            var cols = 10;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Boolean);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "boolTest.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        public void IfWriting10x10CellsShouldWriteBlankValuesForCells()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheets = 1;
            var rows = 10;
            var cols = 10;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Blank);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "blankTest.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        public void IfWriting10x10CellsShouldWriteFormulaValuesForCells()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheets = 1;
            var rows = 10;
            var cols = 10;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Formula);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "formulaTest.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        public void IfWriting10x10CellsShouldWriteErrorValuesForCells()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheets = 1;
            var rows = 10;
            var cols = 10;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Error);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "errorTest.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        [Ignore("This takes a long time to run.")]
        public void IfWritingMaxCellsForWorksheetShouldNotThrowOutOfMemoryException()
        {
            Assert.Fail("This takes a long time to run.");
            _objectToTest = new SXSSFWorkbook();
            var sheets = 1;
            var rows = 1048576;
            var cols = 16384;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Numeric);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "maxCellsWorksheet.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        [Ignore("consume memory")]
        public void IfWorkbookIsSetToUseCompressionShouldUseGZIPDataWriter()
        {
            //Assert.Fail("This takes a long time to run.");
            _objectToTest = new SXSSFWorkbook(null, 100, true);
            var sheets = 1;
            var rows = 10000;
            var cols = 100;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Numeric);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "maxCellsWorksheetZip.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        [Ignore("consume memory")]
        public void IfWriting20WorksheetsWith10000x100CellsShouldNotThrowOutOfMemoryException()
        {
            _objectToTest = new SXSSFWorkbook();
            var sheets = 20;
            var rows = 10000;
            var cols = 100;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Numeric);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "largeWorksheet.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }

        [Test]
        [Ignore("consume memory")]
        public void IfWriting20WorksheetsWith10000x100CellsUsingGzipShouldNotThrowOutOfMemoryException()
        {
            _objectToTest = new SXSSFWorkbook(null, 100, true);
            var sheets = 20;
            var rows = 10000;
            var cols = 100;
            AddCells(_objectToTest, sheets, rows, cols, CellType.Numeric);
            var savePath = Path.Combine(TestContext.CurrentContext.TestDirectory, "largeGzipWorksheet.xlsx");
            WriteFile(savePath, _objectToTest);

            Assert.True(File.Exists(savePath));
            File.Delete(savePath);
        }


        private void AddCells(IWorkbook wb, int sheets, int rows, int columns, CellType type)
        {
            for (int j = 0; j < sheets; j++)
            {
                var sheet = wb.CreateSheet(j.ToString());
                for (int k = 0; k < rows; k++)
                {
                    var row = sheet.CreateRow(k);
                    for (int i = 0; i < columns; i++)
                    {
                        WriteCellValue(row, type, i, i);
                    }
                }
            }
        }


        private void WriteFile(string saveAsPath, SXSSFWorkbook wb)
        {
            //Passing SXSSFWorkbook because IWorkbook does not implement .Dispose which cleans ups temporary files.
            using (FileStream fs = new FileStream(saveAsPath, FileMode.Create, FileAccess.ReadWrite))
            {
                wb.Write(fs);
            }

            wb.Dispose();
        }

        private void WriteCellValue(IRow row, CellType type, int col, object val)
        {
            if (type == CellType.Numeric)
            {
                row.CreateCell(col).SetCellValue(Convert.ToInt32(val));
            }
            else if (type == CellType.String)
            {
                row.CreateCell(col).SetCellValue("\"\'\'<>\\t\\n\\r&\\\"?         test:SLDFKj    \"");
            }
            else if (type == CellType.Boolean)
            {
                row.CreateCell(col).SetCellValue(true);
            }
            else if (type == CellType.Blank)
            {
                row.CreateCell(col);
            }
            else if (type == CellType.Error)
            {
                row.CreateCell(col).SetCellErrorValue(0);
            }
            else if (type == CellType.Formula)
            {
                row.CreateCell(col).SetCellFormula("SUM(A1:A2)");
            }
        }

        [Test]
        public void StreamShouldBeLeavedOpen()
        {
            using (SXSSFWorkbook workbook = new SXSSFWorkbook())
            {
                ISheet sheet = workbook.CreateSheet("Sheet1");

                // Write a large number of rows and columns to cause OutOfMemoryException
                for (int rowNumber = 0; rowNumber < 10; rowNumber++) // Increase this number for more rows
                {
                    IRow row = sheet.CreateRow(rowNumber);
                    for (int colNumber = 0; colNumber < 100; colNumber++) // Increase this number for more columns
                    {
                        ICell cell = row.CreateCell(colNumber);
                        cell.SetCellValue($"Row {rowNumber + 1}, Column {colNumber + 1}");
                    }
                }

                // Write the workbook data to a MemoryStream
                using (var stream = new MemoryStream())
                {
                    workbook.Write(stream, true);
                    Assert.IsTrue(stream.CanRead);
                }
            }
        }
    }
}
