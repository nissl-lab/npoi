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

using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFColumn
    {
        public TestXSSFColumn()
        {

        }

        [Test]
        public void Sheet_SheetIsNotNull()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column = sheet.CreateColumn(columnIndex);

            Assert.NotNull(column.Sheet);
            Assert.AreEqual(column.Sheet, sheet);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(columnIndex);

            Assert.NotNull(columnLoaded.Sheet);
            Assert.AreEqual(columnLoaded.Sheet, sheetLoaded);
        }

        [Test]
        public void FirstCellNumTest()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column1 = sheet.CreateColumn(columnIndex);
            IColumn column2 = sheet.CreateColumn(columnIndex + 1);
            int firstRowNum = 5;
            int numOfCells = 10;

            for (int i = 0; i < numOfCells; i++)
            {
                _ = column2.CreateCell(firstRowNum + i);
            }

            for (int i = 0; i < numOfCells; i++)
            {
                _ = column2.CreateCell(firstRowNum + 20 + i);
            }

            Assert.AreEqual(-1, column1.FirstCellNum);
            Assert.AreEqual(firstRowNum, column2.FirstCellNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn column1Loaded = sheetLoaded.GetColumn(columnIndex);
            IColumn column2Loaded = sheetLoaded.GetColumn(columnIndex + 1);

            Assert.AreEqual(-1, column1Loaded.FirstCellNum);
            Assert.AreEqual(firstRowNum, column2Loaded.FirstCellNum);
        }

        [Test]
        public void LastCellNumTest()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column1 = sheet.CreateColumn(columnIndex);
            IColumn column2 = sheet.CreateColumn(columnIndex + 1);
            int firstRowNum = 5;
            int numOfCells = 10;

            for (int i = 0; i < numOfCells; i++)
            {
                _ = column2.CreateCell(firstRowNum + i);
            }

            for (int i = 0; i < numOfCells; i++)
            {
                _ = column2.CreateCell(firstRowNum + 20 + i);
            }

            Assert.AreEqual(-1, column1.LastCellNum);
            Assert.AreEqual(firstRowNum + 30, column2.LastCellNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn column1Loaded = sheetLoaded.GetColumn(columnIndex);
            IColumn column2Loaded = sheetLoaded.GetColumn(columnIndex + 1);

            Assert.AreEqual(-1, column1Loaded.LastCellNum);
            Assert.AreEqual(firstRowNum + 30, column2Loaded.LastCellNum);
        }

        [Test]
        public void WidthTest()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column1 = sheet.CreateColumn(columnIndex);
            IColumn column2 = sheet.CreateColumn(columnIndex + 1);
            short width = 20;

            column2.Width = width;

            Assert.AreEqual(sheet.DefaultColumnWidth, column1.Width);
            Assert.AreEqual(width, column2.Width);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn column1Loaded = sheetLoaded.GetColumn(columnIndex);
            IColumn column2Loaded = sheetLoaded.GetColumn(columnIndex + 1);

            Assert.AreEqual(sheetLoaded.DefaultColumnWidth, column1Loaded.Width);
            Assert.AreEqual(width, column2Loaded.Width);
        }

        [Test]
        public void PhysicalNumberOfCellsTest()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column1 = sheet.CreateColumn(columnIndex);
            IColumn column2 = sheet.CreateColumn(columnIndex + 1);
            int firstRowNum = 5;
            int numOfCells = 10;

            for (int i = 0; i < numOfCells; i++)
            {
                _ = column2.CreateCell(firstRowNum + i);
            }

            for (int i = 0; i < numOfCells; i++)
            {
                _ = column2.CreateCell(firstRowNum + 20 + i);
            }

            column2.RemoveCell(
                sheet.GetRow(firstRowNum).GetCell(column2.ColumnNum));

            Assert.AreEqual(0, column1.PhysicalNumberOfCells);
            Assert.AreEqual((numOfCells * 2) - 1, column2.PhysicalNumberOfCells);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn column1Loaded = sheetLoaded.GetColumn(columnIndex);
            IColumn column2Loaded = sheetLoaded.GetColumn(columnIndex + 1);

            Assert.AreEqual(0, column1Loaded.PhysicalNumberOfCells);
            Assert.AreEqual((numOfCells * 2) - 1, column2Loaded.PhysicalNumberOfCells);
        }

        [Test]
        public void ZeroWidthTest()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column1 = sheet.CreateColumn(columnIndex);
            IColumn column2 = sheet.CreateColumn(columnIndex + 1);
            short width = 20;

            column2.Width = width;
            column2.ZeroWidth = true;

            Assert.IsFalse(column1.ZeroWidth);
            Assert.AreEqual(sheet.DefaultColumnWidth, column1.Width);
            Assert.IsTrue(column2.ZeroWidth);
            Assert.AreEqual(width, column2.Width);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn column1Loaded = sheetLoaded.GetColumn(columnIndex);
            IColumn column2Loaded = sheetLoaded.GetColumn(columnIndex + 1);

            Assert.IsFalse(column1Loaded.ZeroWidth);
            Assert.AreEqual(sheetLoaded.DefaultColumnWidth, column1Loaded.Width);
            Assert.IsTrue(column2Loaded.ZeroWidth);
            Assert.AreEqual(width, column2Loaded.Width);
        }

        [Test]
        public void ColumnStyleTest()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column1 = sheet.CreateColumn(columnIndex);
            IColumn column2 = sheet.CreateColumn(columnIndex + 1);
            IColumn column3 = sheet.CreateColumn(columnIndex + 2);

            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.BorderLeft = BorderStyle.Double;

            column2.ColumnStyle = cellStyle;
            column3.ColumnStyle = cellStyle;

            Assert.IsNull(column1.ColumnStyle);
            Assert.AreEqual(BorderStyle.None, column1.CreateCell(1).CellStyle.BorderLeft);
            Assert.IsNotNull(column2.ColumnStyle);
            Assert.AreEqual(BorderStyle.Double, column2.ColumnStyle.BorderLeft);
            Assert.AreEqual(BorderStyle.Double, column2.CreateCell(1).CellStyle.BorderLeft);
            Assert.AreEqual(BorderStyle.Double, column3.ColumnStyle.BorderLeft);
            Assert.AreEqual(BorderStyle.Double, column3.CreateCell(1).CellStyle.BorderLeft);

            column3.ColumnStyle = null;

            Assert.IsNull(column3.ColumnStyle);
            Assert.AreEqual(BorderStyle.None, column3.Cells.FirstOrDefault().CellStyle.BorderLeft);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn column1Loaded = sheetLoaded.GetColumn(columnIndex);
            IColumn column2Loaded = sheetLoaded.GetColumn(columnIndex + 1);
            IColumn column3Loaded = sheetLoaded.GetColumn(columnIndex + 2);

            Assert.IsNull(column1Loaded.ColumnStyle);
            Assert.AreEqual(BorderStyle.None, column1Loaded.GetCell(1).CellStyle.BorderLeft);
            Assert.IsNotNull(column2Loaded.ColumnStyle);
            Assert.AreEqual(BorderStyle.Double, column2Loaded.ColumnStyle.BorderLeft);
            Assert.AreEqual(BorderStyle.Double, column2Loaded.GetCell(1).CellStyle.BorderLeft);
            Assert.IsNull(column3Loaded.ColumnStyle);
            Assert.AreEqual(BorderStyle.None, column3Loaded.Cells.FirstOrDefault().CellStyle.BorderLeft);
        }

        [Test]
        public void IsFormattedTest()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column1 = sheet.CreateColumn(columnIndex);
            IColumn column2 = sheet.CreateColumn(columnIndex + 1);

            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.BorderLeft = BorderStyle.Double;

            column2.ColumnStyle = cellStyle;

            Assert.IsFalse(column1.IsFormatted);
            Assert.IsTrue(column2.IsFormatted);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn column1Loaded = sheetLoaded.GetColumn(columnIndex);
            IColumn column2Loaded = sheetLoaded.GetColumn(columnIndex + 1);

            Assert.IsFalse(column1Loaded.IsFormatted);
            Assert.IsTrue(column2Loaded.IsFormatted);
        }

        [Test]
        public void CreateCell_WithValidIndex_CellCreated()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            int rowIndex = 10;
            IColumn column = sheet.CreateColumn(columnIndex);
            ICell cell = column.CreateCell(rowIndex);

            Assert.NotNull(cell);
            Assert.AreEqual(CellType.Blank, cell.CellType);
            Assert.NotNull(sheet.GetRow(rowIndex));
            Assert.NotNull(sheet.GetRow(rowIndex).GetCell(columnIndex));

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            ICell cellLoaded = sheetLoaded.GetColumn(columnIndex).GetCell(rowIndex);

            Assert.NotNull(cellLoaded);
            Assert.AreEqual(CellType.Blank, cellLoaded.CellType);
            Assert.NotNull(sheetLoaded.GetRow(rowIndex));
            Assert.NotNull(sheetLoaded.GetRow(rowIndex).GetCell(columnIndex));
        }

        [Test]
        public void CreateCell_WithInValidIndex_ExceptiomnIsThrown()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;

            IColumn column = sheet.CreateColumn(columnIndex);

            _ = Assert.Throws<ArgumentOutOfRangeException>(() => column.CreateCell(-1));
            _ = Assert.Throws<ArgumentOutOfRangeException>(() => column.CreateCell(2000000));
        }

        [Test]
        public void CreateCell_WithValidIndexVariousTypes_CellsCreatedWithCorrectTypes()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            IColumn column = sheet.CreateColumn(columnIndex);
            _ = column.CreateCell(0, CellType.Blank);
            _ = column.CreateCell(1, CellType.Boolean);
            _ = column.CreateCell(2, CellType.Error);
            _ = column.CreateCell(3, CellType.Formula);
            _ = column.CreateCell(4, CellType.Numeric);
            _ = column.CreateCell(5, CellType.String);

            for (int i = 0; i <= 5; i++)
            {
                Assert.NotNull(sheet.GetRow(i));
                Assert.NotNull(sheet.GetRow(i).GetCell(columnIndex));
            }

            Assert.AreEqual(CellType.Blank, sheet.GetRow(0).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.Boolean, sheet.GetRow(1).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.Error, sheet.GetRow(2).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.Formula, sheet.GetRow(3).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.Blank, sheet.GetRow(4).GetCell(columnIndex).CellType); // Numeric cell with no data will return blank by default
            Assert.AreEqual(CellType.String, sheet.GetRow(5).GetCell(columnIndex).CellType);

            _ = Assert.Throws<ArgumentException>(() => column.CreateCell(6, CellType.Unknown));
            Assert.IsNull(sheet.GetRow(6));
            Assert.IsNull(column.GetCell(6));

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            for (int i = 0; i <= 5; i++)
            {
                Assert.NotNull(sheetLoaded.GetRow(i));
                Assert.NotNull(sheetLoaded.GetRow(i).GetCell(columnIndex));
            }

            Assert.AreEqual(CellType.Blank, sheetLoaded.GetRow(0).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.Boolean, sheetLoaded.GetRow(1).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.Error, sheetLoaded.GetRow(2).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.Formula, sheetLoaded.GetRow(3).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.Blank, sheetLoaded.GetRow(4).GetCell(columnIndex).CellType);
            Assert.AreEqual(CellType.String, sheetLoaded.GetRow(5).GetCell(columnIndex).CellType);

            Assert.IsNull(sheetLoaded.GetRow(6));
            Assert.IsNull(sheetLoaded.GetColumn(columnIndex).GetCell(6));
        }

        [Test]
        public void GetCell_GetExistingAndNonExistinCells_CellsReturnedWhenExistsNullsWhenNot()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            int rowIndex = 10;
            IColumn column = sheet.CreateColumn(columnIndex);
            _ = column.CreateCell(rowIndex);
            ICell cell = column.GetCell(rowIndex);
            _ = sheet.CreateRow(rowIndex + 1).CreateCell(columnIndex);

            Assert.IsNotNull(cell);
            Assert.AreEqual(columnIndex, cell.ColumnIndex);
            Assert.AreEqual(rowIndex, cell.RowIndex);
            Assert.IsNull(column.GetCell(0));
            Assert.IsNotNull(column.GetCell(rowIndex + 1));

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(columnIndex);
            ICell cellLoaded = columnLoaded.GetCell(rowIndex);

            Assert.IsNotNull(cellLoaded);
            Assert.AreEqual(columnIndex, cellLoaded.ColumnIndex);
            Assert.AreEqual(rowIndex, cellLoaded.RowIndex);
            Assert.IsNull(columnLoaded.GetCell(0));
            Assert.IsNotNull(columnLoaded.GetCell(rowIndex + 1));
        }

        [Test]
        public void GetCell_GetExistingCells_CellsReturned()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            int rowIndex = 10;
            IColumn column = sheet.CreateColumn(columnIndex);

            _ = column.CreateCell(rowIndex, CellType.Blank);
            column.CreateCell(rowIndex + 1, CellType.Blank).SetCellValue(1);
            _ = column.CreateCell(rowIndex + 2, CellType.Numeric);
            column.CreateCell(rowIndex + 3, CellType.Numeric).SetCellValue(1);

            Assert.IsNotNull(column.GetCell(rowIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK));
            Assert.IsNotNull(column.GetCell(rowIndex + 1, MissingCellPolicy.RETURN_NULL_AND_BLANK));
            Assert.IsNotNull(column.GetCell(rowIndex + 2, MissingCellPolicy.RETURN_NULL_AND_BLANK));
            Assert.IsNotNull(column.GetCell(rowIndex + 3, MissingCellPolicy.RETURN_NULL_AND_BLANK));
            Assert.IsNull(column.GetCell(rowIndex + 4, MissingCellPolicy.RETURN_NULL_AND_BLANK));

            Assert.IsNull(column.GetCell(rowIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            Assert.IsNotNull(column.GetCell(rowIndex + 1, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            Assert.IsNull(column.GetCell(rowIndex + 2, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            Assert.IsNotNull(column.GetCell(rowIndex + 3, MissingCellPolicy.RETURN_BLANK_AS_NULL));
            Assert.IsNull(column.GetCell(rowIndex + 4, MissingCellPolicy.RETURN_BLANK_AS_NULL));

            Assert.IsNotNull(column.GetCell(rowIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK));
            Assert.IsNotNull(column.GetCell(rowIndex + 1, MissingCellPolicy.CREATE_NULL_AS_BLANK));
            Assert.IsNotNull(column.GetCell(rowIndex + 2, MissingCellPolicy.CREATE_NULL_AS_BLANK));
            Assert.IsNotNull(column.GetCell(rowIndex + 3, MissingCellPolicy.CREATE_NULL_AS_BLANK));
            Assert.IsNotNull(column.GetCell(rowIndex + 4, MissingCellPolicy.CREATE_NULL_AS_BLANK));
        }

        [Test]
        public void RemoveCell_RemoveExistingCells_CellsAreRemoved()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            int rowIndex = 10;
            int middleRowIndex = 15;
            int initialNumberOfCells = 10;
            IColumn column = sheet.CreateColumn(columnIndex);

            for (int i = 0; i < initialNumberOfCells; i++)
            {
                _ = column.CreateCell(rowIndex + i);
            }

            Assert.AreEqual(initialNumberOfCells, column.PhysicalNumberOfCells);
            Assert.AreEqual(rowIndex, column.FirstCellNum);
            Assert.AreEqual(rowIndex + initialNumberOfCells, column.LastCellNum);

            column.RemoveCell(column.GetCell(column.FirstCellNum));
            column.RemoveCell(column.GetCell(column.LastCellNum - 1));
            column.RemoveCell(column.GetCell(middleRowIndex));

            Assert.AreEqual(initialNumberOfCells - 3, column.PhysicalNumberOfCells);
            Assert.AreEqual(rowIndex + 1, column.FirstCellNum);
            Assert.AreEqual(rowIndex + initialNumberOfCells - 1, column.LastCellNum);
            Assert.IsNull(column.GetCell(rowIndex));
            Assert.IsNull(column.GetCell(middleRowIndex));
            Assert.IsNull(column.GetCell(rowIndex + initialNumberOfCells));
            Assert.IsNull(sheet.GetRow(rowIndex).GetCell(columnIndex));
            Assert.IsNull(sheet.GetRow(middleRowIndex).GetCell(columnIndex));
            Assert.IsNull(sheet.GetRow(rowIndex + initialNumberOfCells - 1).GetCell(columnIndex));
        }

        [Test]
        public void RemoveCell_RemoveCellNotBelongingToColumn_ExceptionThrownCellIsNotRemoved()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int columnIndex = 10;
            int rowIndex = 10;
            IColumn column = sheet.CreateColumn(columnIndex);
            IRow row = sheet.CreateRow(rowIndex);
            ICell cell = row.CreateCell(columnIndex + 1);

            _ = Assert.Throws<ArgumentException>(() => column.RemoveCell(cell));
            _ = Assert.Throws<ArgumentException>(() => column.RemoveCell(null));
            Assert.IsNotNull(row.GetCell(columnIndex + 1));
        }

        [Test]
        public void CopyColumnFrom_CopyOverExistingCells_ExistingCellsAreReplacedByCellsFromCopiedColumn()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(wb);
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            XSSFDrawing dg = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFComment comment = (XSSFComment)dg.CreateCellComment(new XSSFClientAnchor());
            comment.Author = "POI";
            comment.String = new XSSFRichTextString("A new comment");
            int originalColIndex = 1;
            int copyColIndex = 4;
            short width = 20;
            int mergedRegionFirstRow = 10;
            int mergedRegionLastRow = 20;
            int firstFormulaCellValue = 10;
            int secondFormulaCellValue = 20;
            int formulaResult = firstFormulaCellValue + secondFormulaCellValue;
            int cellToBeErasedRowIndex = 25;

            XSSFColumn anotherColumn = (XSSFColumn)sheet.CreateColumn(copyColIndex);
            anotherColumn.CreateCell(0).SetCellValue("POI");
            anotherColumn.Width = (short)(width * 2);
            anotherColumn.CreateCell(0).SetCellValue("POI");
            anotherColumn.CreateCell(cellToBeErasedRowIndex).SetCellValue("POI");
            XSSFColumn originalColumn = (XSSFColumn)sheet.CreateColumn(originalColIndex);
            originalColumn.Width = width;
            XSSFCell formulaCell1 = (XSSFCell)originalColumn.CreateCell(0);
            formulaCell1.SetCellFormula("C1 + D1");
            formulaCell1.CellComment = comment;
            sheet.CreateColumn(originalColIndex + 1).CreateCell(0).SetCellValue(firstFormulaCellValue);
            sheet.CreateColumn(originalColIndex + 2).CreateCell(0).SetCellValue(secondFormulaCellValue);
            sheet.GetRow(0).CreateCell(originalColIndex + 1 + 3).SetCellValue(firstFormulaCellValue);
            sheet.GetRow(0).CreateCell(originalColIndex + 2 + 3).SetCellValue(secondFormulaCellValue);
            originalColumn.CreateCell(mergedRegionFirstRow).SetCellValue("POI");
            _ = sheet.AddMergedRegion(
                new CellRangeAddress(
                    mergedRegionFirstRow,
                    mergedRegionLastRow,
                    originalColumn.ColumnNum,
                    originalColumn.ColumnNum));
            CellRangeAddress originaMergedRegion = sheet.GetMergedRegion(0);

            fe.EvaluateAll();

            Assert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(originalColIndex).StringCellValue);
            Assert.AreEqual("POI", formulaCell1.CellComment.Author);
            Assert.AreEqual(formulaResult, formulaCell1.NumericCellValue);
            Assert.AreEqual(1, sheet.NumMergedRegions);
            Assert.NotNull(originaMergedRegion);

            anotherColumn.CopyColumnFrom(originalColumn, new CellCopyPolicy());
            CellRangeAddress copyMergedRegion = sheet.GetMergedRegion(1);
            fe.EvaluateAll();

            XSSFCell formulaCell2 = (XSSFCell)sheet.GetColumn(copyColIndex).GetCell(0);

            Assert.AreEqual(copyColIndex, anotherColumn.ColumnNum);
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(4, sheet.LastColumnNum);
            Assert.AreEqual(width, anotherColumn.Width);
            Assert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            Assert.AreEqual("F1+G1", formulaCell2.CellFormula);
            Assert.AreEqual(formulaResult, formulaCell2.NumericCellValue);
            //Assert.IsNull(anotherColumn.GetCell(cellToBeErasedRowIndex));
            Assert.AreEqual(2, sheet.NumMergedRegions);
            Assert.NotNull(originaMergedRegion);
            Assert.NotNull(copyMergedRegion);
            Assert.AreEqual(mergedRegionFirstRow, copyMergedRegion.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, copyMergedRegion.LastRow);
            Assert.AreEqual(anotherColumn.ColumnNum, copyMergedRegion.FirstColumn);
            Assert.AreEqual(anotherColumn.ColumnNum, copyMergedRegion.LastColumn);

            FileInfo file = TempFile.CreateTempFile("CopyColumnFrom-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFFormulaEvaluator feLoaded = new XSSFFormulaEvaluator(wbLoaded);
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            XSSFColumn originalColumnLoaded = (XSSFColumn)sheetLoaded.GetColumn(originalColIndex);
            XSSFColumn copyColumnLoaded = (XSSFColumn)sheetLoaded.GetColumn(copyColIndex);
            XSSFCell formulaCellLoaded = (XSSFCell)copyColumnLoaded.GetCell(0);
            CellRangeAddress originalMergedRegionLoaded = sheet.GetMergedRegion(0);
            CellRangeAddress copyMergedRegionLoaded = sheet.GetMergedRegion(1);

            feLoaded.EvaluateAll();

            Assert.AreEqual(1, sheetLoaded.FirstColumnNum);
            Assert.AreEqual(4, sheetLoaded.LastColumnNum);
            Assert.AreEqual(width, copyColumnLoaded.Width);
            Assert.AreEqual(firstFormulaCellValue, sheetLoaded.GetColumn(originalColIndex + 1).GetCell(0).NumericCellValue);
            Assert.AreEqual(secondFormulaCellValue, sheetLoaded.GetColumn(originalColIndex + 2).GetCell(0).NumericCellValue);
            Assert.AreEqual("POI", sheetLoaded.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            Assert.AreEqual("F1+G1", formulaCellLoaded.CellFormula);
            Assert.AreEqual(formulaResult, formulaCellLoaded.NumericCellValue);
            Assert.NotNull(originalMergedRegionLoaded);
            Assert.AreEqual(mergedRegionFirstRow, originalMergedRegionLoaded.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, originalMergedRegionLoaded.LastRow);
            Assert.AreEqual(originalColumnLoaded.ColumnNum, originalMergedRegionLoaded.FirstColumn);
            Assert.AreEqual(originalColumnLoaded.ColumnNum, originalMergedRegionLoaded.LastColumn);
            Assert.NotNull(copyMergedRegionLoaded);
            Assert.AreEqual(mergedRegionFirstRow, copyMergedRegionLoaded.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, copyMergedRegionLoaded.LastRow);
            Assert.AreEqual(copyColumnLoaded.ColumnNum, copyMergedRegionLoaded.FirstColumn);
            Assert.AreEqual(copyColumnLoaded.ColumnNum, copyMergedRegionLoaded.LastColumn);
        }

        [Test]
        public void CopyColumnFrom_CopyOntoEmptyColumn_CopiedCellsAreAddedToColumn()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(wb);
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            XSSFDrawing dg = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFComment comment = (XSSFComment)dg.CreateCellComment(new XSSFClientAnchor());
            comment.Author = "POI";
            comment.String = new XSSFRichTextString("A new comment");
            int originalColIndex = 1;
            int copyColIndex = 4;
            short width = 20;
            int mergedRegionFirstRow = 10;
            int mergedRegionLastRow = 20;
            int firstFormulaCellValue = 10;
            int secondFormulaCellValue = 20;
            int formulaResult = firstFormulaCellValue + secondFormulaCellValue;

            XSSFColumn anotherColumn = (XSSFColumn)sheet.CreateColumn(copyColIndex);
            XSSFColumn originalColumn = (XSSFColumn)sheet.CreateColumn(originalColIndex);
            originalColumn.Width = width;
            XSSFCell formulaCell1 = (XSSFCell)originalColumn.CreateCell(0);
            formulaCell1.SetCellFormula("C1 + D1");
            formulaCell1.CellComment = comment;
            sheet.CreateColumn(originalColIndex + 1).CreateCell(0).SetCellValue(firstFormulaCellValue);
            sheet.CreateColumn(originalColIndex + 2).CreateCell(0).SetCellValue(secondFormulaCellValue);
            sheet.GetRow(0).CreateCell(originalColIndex + 1 + 3).SetCellValue(firstFormulaCellValue);
            sheet.GetRow(0).CreateCell(originalColIndex + 2 + 3).SetCellValue(secondFormulaCellValue);
            originalColumn.CreateCell(mergedRegionFirstRow).SetCellValue("POI");
            _ = sheet.AddMergedRegion(
                new CellRangeAddress(
                    mergedRegionFirstRow,
                    mergedRegionLastRow,
                    originalColumn.ColumnNum,
                    originalColumn.ColumnNum));
            CellRangeAddress originaMergedRegion = sheet.GetMergedRegion(0);

            fe.EvaluateAll();

            Assert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(originalColIndex).StringCellValue);
            Assert.AreEqual("POI", formulaCell1.CellComment.Author);
            Assert.AreEqual(formulaResult, formulaCell1.NumericCellValue);
            Assert.AreEqual(1, sheet.NumMergedRegions);
            Assert.NotNull(originaMergedRegion);

            anotherColumn.CopyColumnFrom(originalColumn, new CellCopyPolicy() { IsCopyColumnWidth = false });
            CellRangeAddress copyMergedRegion = sheet.GetMergedRegion(1);
            fe.EvaluateAll();

            XSSFCell formulaCell2 = (XSSFCell)sheet.GetColumn(copyColIndex).GetCell(0);

            Assert.AreEqual(copyColIndex, anotherColumn.ColumnNum);
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(4, sheet.LastColumnNum);
            Assert.AreNotEqual(width, anotherColumn.Width);
            Assert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            Assert.AreEqual("F1+G1", formulaCell2.CellFormula);
            Assert.AreEqual(formulaResult, formulaCell2.NumericCellValue);
            Assert.AreEqual(2, sheet.NumMergedRegions);
            Assert.NotNull(originaMergedRegion);
            Assert.NotNull(copyMergedRegion);
            Assert.AreEqual(mergedRegionFirstRow, copyMergedRegion.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, copyMergedRegion.LastRow);
            Assert.AreEqual(anotherColumn.ColumnNum, copyMergedRegion.FirstColumn);
            Assert.AreEqual(anotherColumn.ColumnNum, copyMergedRegion.LastColumn);

            FileInfo file = TempFile.CreateTempFile("CopyColumnFrom-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFFormulaEvaluator feLoaded = new XSSFFormulaEvaluator(wbLoaded);
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            XSSFColumn originalColumnLoaded = (XSSFColumn)sheetLoaded.GetColumn(originalColIndex);
            XSSFColumn copyColumnLoaded = (XSSFColumn)sheetLoaded.GetColumn(copyColIndex);
            XSSFCell formulaCellLoaded = (XSSFCell)copyColumnLoaded.GetCell(0);
            CellRangeAddress originalMergedRegionLoaded = sheet.GetMergedRegion(0);
            CellRangeAddress copyMergedRegionLoaded = sheet.GetMergedRegion(1);

            feLoaded.EvaluateAll();

            Assert.AreEqual(1, sheetLoaded.FirstColumnNum);
            Assert.AreEqual(4, sheetLoaded.LastColumnNum);

            Assert.AreNotEqual(width, copyColumnLoaded.Width);
            Assert.AreEqual(firstFormulaCellValue, sheetLoaded.GetColumn(originalColIndex + 1).GetCell(0).NumericCellValue);
            Assert.AreEqual(secondFormulaCellValue, sheetLoaded.GetColumn(originalColIndex + 2).GetCell(0).NumericCellValue);
            Assert.AreEqual("POI", sheetLoaded.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            Assert.AreEqual("F1+G1", formulaCellLoaded.CellFormula);
            Assert.AreEqual(formulaResult, formulaCellLoaded.NumericCellValue);
            Assert.NotNull(originalMergedRegionLoaded);
            Assert.AreEqual(mergedRegionFirstRow, originalMergedRegionLoaded.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, originalMergedRegionLoaded.LastRow);
            Assert.AreEqual(originalColumnLoaded.ColumnNum, originalMergedRegionLoaded.FirstColumn);
            Assert.AreEqual(originalColumnLoaded.ColumnNum, originalMergedRegionLoaded.LastColumn);
            Assert.NotNull(copyMergedRegionLoaded);
            Assert.AreEqual(mergedRegionFirstRow, copyMergedRegionLoaded.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, copyMergedRegionLoaded.LastRow);
            Assert.AreEqual(copyColumnLoaded.ColumnNum, copyMergedRegionLoaded.FirstColumn);
            Assert.AreEqual(copyColumnLoaded.ColumnNum, copyMergedRegionLoaded.LastColumn);
        }

        [Test]
        public void CopyColumnFrom_CopyFromNull_ExistingCellsAreReplacedWithNull()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int copyColIndex = 4;
            int cellToBeErasedRowIndex = 25;

            XSSFColumn anotherColumn = (XSSFColumn)sheet.CreateColumn(copyColIndex);
            anotherColumn.CreateCell(0).SetCellValue("POI");
            anotherColumn.CreateCell(cellToBeErasedRowIndex).SetCellValue("POI");

            anotherColumn.CopyColumnFrom(null, new CellCopyPolicy());

            Assert.AreEqual(copyColIndex, anotherColumn.ColumnNum);
            Assert.AreEqual(4, sheet.FirstColumnNum);
            Assert.AreEqual(4, sheet.LastColumnNum);
            Assert.AreEqual("", anotherColumn.GetCell(0).StringCellValue);
            Assert.AreEqual("", anotherColumn.GetCell(cellToBeErasedRowIndex).StringCellValue);
        }

        [Test]
        public void CopyColumnFrom_FromExternalSheet()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet srcSheet = (XSSFSheet)workbook.CreateSheet("src");
            XSSFSheet destSheet = workbook.CreateSheet("dest") as XSSFSheet;
            _ = workbook.CreateSheet("other");

            IColumn srcColumn = srcSheet.CreateColumn(0);
            int col = 0;
            //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
            srcColumn.CreateCell(col++).CellFormula = "E5";
            srcColumn.CreateCell(col++).CellFormula = "src!E5";
            srcColumn.CreateCell(col++).CellFormula = "dest!E5";
            srcColumn.CreateCell(col++).CellFormula = "other!E5";

            //Test 2D and 3D Ref Ptgs with absolute column
            srcColumn.CreateCell(col++).CellFormula = "$E5";
            srcColumn.CreateCell(col++).CellFormula = "src!$E5";
            srcColumn.CreateCell(col++).CellFormula = "dest!$E5";
            srcColumn.CreateCell(col++).CellFormula = "other!$E5";

            //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
            srcColumn.CreateCell(col++).CellFormula = "SUM(E5:$E10)";
            srcColumn.CreateCell(col++).CellFormula = "SUM(src!E5:$E10)";
            srcColumn.CreateCell(col++).CellFormula = "SUM(dest!E5:$E10)";
            srcColumn.CreateCell(col++).CellFormula = "SUM(other!E5:$E10)";
            //////////////////
            XSSFColumn destColumn = destSheet.CreateColumn(1) as XSSFColumn;
            destColumn.CopyColumnFrom(srcColumn, new CellCopyPolicy());

            //////////////////

            //Test 2D and 3D Ref Ptgs (Pxg for OOXML Workbooks)
            col = 0;
            ICell cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("F5", cell.CellFormula, "RefPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("src!F5", cell.CellFormula, "Ref3DPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("dest!F5", cell.CellFormula, "Ref3DPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("other!F5", cell.CellFormula, "Ref3DPtg");

            /////////////////////////////////////////////

            //Test 2D and 3D Ref Ptgs with absolute column (Ptg column number shouldn't change)
            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("$E5", cell.CellFormula, "RefPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("src!$E5", cell.CellFormula, "Ref3DPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("dest!$E5", cell.CellFormula, "Ref3DPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("other!$E5", cell.CellFormula, "Ref3DPtg");

            //////////////////////////////////////////

            //Test 2D and 3D Area Ptgs (Pxg for OOXML Workbooks)
            // Note: absolute column changes from last cell to first cell in order
            // to maintain topLeft:bottomRight order
            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("SUM($E5:F10)", cell.CellFormula, "Area2DPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(cell);
            Assert.AreEqual("SUM(src!$E5:F10)", cell.CellFormula, "Area3DPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(destColumn.GetCell(6));
            Assert.AreEqual("SUM(dest!$E5:F10)", cell.CellFormula, "Area3DPtg");

            cell = destColumn.GetCell(col++);
            Assert.IsNotNull(destColumn.GetCell(7));
            Assert.AreEqual("SUM(other!$E5:F10)", cell.CellFormula, "Area3DPtg");

            workbook.Close();
        }//*/
    }
}
