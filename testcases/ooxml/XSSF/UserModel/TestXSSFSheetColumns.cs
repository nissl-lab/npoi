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

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFSheetColumns
    {
        public TestXSSFSheetColumns()
        {

        }

        [Test]
        public void FirstColumnNum_ReturnsCorrectNumber()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            _ = sheet.CreateColumn(10);

            Assert.AreEqual(10, sheet.FirstColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            Assert.AreEqual(10, sheetLoaded.FirstColumnNum);
        }

        [Test]
        public void LastColumnNum_ReturnsCorrectNumber()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            _ = sheet.CreateColumn(10);

            Assert.AreEqual(10, sheet.LastColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            Assert.AreEqual(10, sheetLoaded.LastColumnNum);
        }

        [Test]
        public void PhysicalNumberOfColumns_ReturnsCorrectNumber()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int colNumber = 12;

            for (int i = 0; i < colNumber; i++)
            {
                _ = sheet.CreateColumn(i);
            }

            Assert.AreEqual(0, sheet.FirstColumnNum);
            Assert.AreEqual(colNumber - 1, sheet.LastColumnNum);
            Assert.AreEqual(colNumber, sheet.PhysicalNumberOfColumns);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            Assert.AreEqual(0, sheetLoaded.FirstColumnNum);
            Assert.AreEqual(colNumber - 1, sheetLoaded.LastColumnNum);
            Assert.AreEqual(colNumber, sheetLoaded.PhysicalNumberOfColumns);
        }

        [Test]
        public void CreateColumn_AtIndex6_ColumnIsCreatedWithCorrectIndex()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int colIndex = 6;

            _ = sheet.CreateColumn(colIndex);

            IColumn column = sheet.GetColumn(colIndex);

            Assert.NotNull(column);
            Assert.AreEqual(colIndex, column.ColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(colIndex);

            Assert.NotNull(columnLoaded);
            Assert.AreEqual(colIndex, columnLoaded.ColumnNum);
        }

        [Test]
        public void CreateColumn_AtIndex6Twice_ColumnIsCreatedWithCorrectIndexAndNoCells()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int colIndex = 6;

            sheet.CreateColumn(colIndex).CreateCell(0).SetCellValue(1);
            IColumn column = sheet.GetColumn(colIndex);

            Assert.NotNull(column);
            Assert.AreEqual(colIndex, column.ColumnNum);
            Assert.AreEqual(1, column.GetCell(0).NumericCellValue);

            IColumn columnOverwrite = sheet.CreateColumn(colIndex);

            Assert.NotNull(columnOverwrite);
            Assert.AreEqual(colIndex, columnOverwrite.ColumnNum);
            Assert.IsNull(columnOverwrite.GetCell(0));

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(colIndex);

            Assert.NotNull(columnLoaded);
            Assert.AreEqual(colIndex, columnLoaded.ColumnNum);
            Assert.IsNull(columnLoaded.GetCell(0));
        }

        [Test]
        public void GetColumn_GetExistingColumn_ReturnsAColumn()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            _ = sheet.CreateColumn(10);
            IColumn column = sheet.GetColumn(10);

            Assert.NotNull(column);
            Assert.AreEqual(10, column.ColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(10);

            Assert.NotNull(columnLoaded);
            Assert.AreEqual(10, columnLoaded.ColumnNum);
        }

        [Test]
        public void GetColumn_GetNonExistingColumn_ReturnsNull()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            IColumn column = sheet.GetColumn(10);

            Assert.IsNull(column);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(10);

            Assert.IsNull(columnLoaded);
        }

        [Test]
        public void RemoveColumn_RemoveExistingColumn_ColumnIsRemoved()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            _ = sheet.CreateColumn(1);
            _ = sheet.CreateColumn(2);
            _ = sheet.CreateColumn(3);

            Assert.NotNull(sheet.GetColumn(1));
            Assert.NotNull(sheet.GetColumn(2));
            Assert.NotNull(sheet.GetColumn(3));
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(3, sheet.LastColumnNum);

            sheet.RemoveColumn(sheet.GetColumn(3));

            Assert.IsNull(sheet.GetColumn(3));
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(2, sheet.LastColumnNum);

            _ = sheet.CreateColumn(3);

            sheet.RemoveColumn(sheet.GetColumn(1));

            Assert.IsNull(sheet.GetColumn(1));
            Assert.AreEqual(2, sheet.FirstColumnNum);
            Assert.AreEqual(3, sheet.LastColumnNum);

            _ = sheet.CreateColumn(1);

            Assert.NotNull(sheet.GetColumn(3));
            Assert.AreEqual(3, sheet.LastColumnNum);

            sheet.RemoveColumn(sheet.GetColumn(2));

            Assert.IsNull(sheet.GetColumn(2));
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(3, sheet.LastColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            Assert.IsNull(sheetLoaded.GetColumn(2));
            Assert.AreEqual(1, sheetLoaded.FirstColumnNum);
            Assert.AreEqual(3, sheetLoaded.LastColumnNum);
        }

        [Test]
        public void RemoveColumn_RemoveNonExistingColumn_ExceptionIsThrown()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            _ = sheet.CreateColumn(1);
            _ = sheet.CreateColumn(2);
            _ = sheet.CreateColumn(3);

            Assert.NotNull(sheet.GetColumn(1));
            Assert.NotNull(sheet.GetColumn(2));
            Assert.NotNull(sheet.GetColumn(3));
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(3, sheet.LastColumnNum);

            _ = Assert.Throws<ArgumentException>(() => sheet.RemoveColumn(sheet.GetColumn(4)));
        }

        /// <summary>
        /// Tests sets up a column with formula, merged region with data in it
        /// custom width and comments and copies in to a place where there is no
        /// column exsits yet. The outcome is exact copy of the column is created
        /// in a new index with no data loss.
        /// </summary>
        [Test]
        public void CopyColumn_ColumnWithDataCommentsFormulasMergedRegionsCopiedToNewPlace_ColumncopiedWithAllItsContents()
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

            XSSFColumn originalColumn = (XSSFColumn)sheet.CreateColumn(originalColIndex);
            originalColumn.Width = width;
            XSSFCell formulaCell1 = (XSSFCell)originalColumn.CreateCell(0);
            formulaCell1.SetCellFormula("C1 + D1");
            formulaCell1.CellComment = comment;
            sheet.CreateColumn(originalColIndex + 1).CreateCell(0).SetCellValue(10);
            sheet.CreateColumn(originalColIndex + 2).CreateCell(0).SetCellValue(20);
            originalColumn.CreateCell(mergedRegionFirstRow).SetCellValue("POI");
            _ = sheet.AddMergedRegion(
                new CellRangeAddress(
                    mergedRegionFirstRow,
                    mergedRegionLastRow,
                    originalColumn.ColumnNum,
                    originalColumn.ColumnNum + 1));
            CellRangeAddress originaMergedRegion = sheet.GetMergedRegion(0);

            fe.EvaluateAll();

            Assert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(originalColIndex).StringCellValue);
            Assert.AreEqual("POI", formulaCell1.CellComment.Author);
            Assert.AreEqual(30, formulaCell1.NumericCellValue);
            Assert.AreEqual(1, sheet.NumMergedRegions);
            Assert.NotNull(originaMergedRegion);

            XSSFColumn copyColumn = (XSSFColumn)sheet.CopyColumn(originalColIndex, copyColIndex);
            CellRangeAddress copyMergedRegion = sheet.GetMergedRegion(1);
            fe.EvaluateAll();

            XSSFCell formulaCell2 = (XSSFCell)sheet.GetColumn(copyColIndex).GetCell(0);

            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(4, sheet.LastColumnNum);
            Assert.AreEqual(width, copyColumn.Width);
            //Assert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            Assert.AreEqual("POI", formulaCell2.CellComment.Author);
            Assert.AreEqual("C1 + D1", formulaCell2.CellFormula);
            Assert.AreEqual(30, formulaCell2.NumericCellValue);
            Assert.AreEqual(2, sheet.NumMergedRegions);
            Assert.NotNull(originaMergedRegion);
            Assert.NotNull(copyMergedRegion);
            Assert.AreEqual(mergedRegionFirstRow, copyMergedRegion.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, copyMergedRegion.LastRow);
            Assert.AreEqual(copyColumn.ColumnNum, copyMergedRegion.FirstColumn);
            Assert.AreEqual(copyColumn.ColumnNum + 1, copyMergedRegion.LastColumn);

            FileInfo file = TempFile.CreateTempFile("ShiftCols-", ".xlsx");
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
            Assert.AreEqual(10, sheetLoaded.GetColumn(originalColIndex + 1).GetCell(0).NumericCellValue);
            Assert.AreEqual(20, sheetLoaded.GetColumn(originalColIndex + 2).GetCell(0).NumericCellValue);
            Assert.AreEqual("POI", sheetLoaded.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            Assert.AreEqual("POI", formulaCellLoaded.CellComment.Author);
            Assert.AreEqual("C1 + D1", formulaCellLoaded.CellFormula);
            Assert.AreEqual(30, formulaCellLoaded.NumericCellValue);
            Assert.NotNull(originalMergedRegionLoaded);
            Assert.AreEqual(mergedRegionFirstRow, originalMergedRegionLoaded.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, originalMergedRegionLoaded.LastRow);
            Assert.AreEqual(originalColumnLoaded.ColumnNum, originalMergedRegionLoaded.FirstColumn);
            Assert.AreEqual(originalColumnLoaded.ColumnNum + 1, originalMergedRegionLoaded.LastColumn);
            Assert.NotNull(copyMergedRegionLoaded);
            Assert.AreEqual(mergedRegionFirstRow, copyMergedRegionLoaded.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, copyMergedRegionLoaded.LastRow);
            Assert.AreEqual(copyColumnLoaded.ColumnNum, copyMergedRegionLoaded.FirstColumn);
            Assert.AreEqual(copyColumnLoaded.ColumnNum + 1, copyMergedRegionLoaded.LastColumn);
        }

        /// <summary>
        /// Tests sets up a column with formula, merged region with data in it
        /// custom width and comments and copies in to a place of pre-existing
        /// column. The outcome is pre-existing column gets shifted to the right,
        /// exact copy of the copied column is created in it's place with no data loss.
        /// </summary>
        [Test]
        public void CopyColumn_ColumnWithDataCommentsFormulasMergedRegionsCopiedToPlaceOfExistingColumn_ExistingColumnIsShiftedColumnCopied()
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

            XSSFColumn anotherColumn = (XSSFColumn)sheet.CreateColumn(copyColIndex);
            anotherColumn.CreateCell(0).SetCellValue("POI");
            anotherColumn.Width = (short)(width * 2);
            XSSFColumn originalColumn = (XSSFColumn)sheet.CreateColumn(originalColIndex);
            originalColumn.Width = width;
            XSSFCell formulaCell1 = (XSSFCell)originalColumn.CreateCell(0);
            formulaCell1.SetCellFormula("C1 + D1");
            formulaCell1.CellComment = comment;
            sheet.CreateColumn(originalColIndex + 1).CreateCell(0).SetCellValue(10);
            sheet.CreateColumn(originalColIndex + 2).CreateCell(0).SetCellValue(20);
            originalColumn.CreateCell(mergedRegionFirstRow).SetCellValue("POI");
            _ = sheet.AddMergedRegion(
                new CellRangeAddress(
                    mergedRegionFirstRow,
                    mergedRegionLastRow,
                    originalColumn.ColumnNum,
                    originalColumn.ColumnNum + 1));
            CellRangeAddress originaMergedRegion = sheet.GetMergedRegion(0);

            fe.EvaluateAll();

            Assert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(originalColIndex).StringCellValue);
            Assert.AreEqual("POI", formulaCell1.CellComment.Author);
            Assert.AreEqual(30, formulaCell1.NumericCellValue);
            Assert.AreEqual(1, sheet.NumMergedRegions);
            Assert.NotNull(originaMergedRegion);

            XSSFColumn copyColumn = (XSSFColumn)sheet.CopyColumn(originalColIndex, copyColIndex);
            CellRangeAddress copyMergedRegion = sheet.GetMergedRegion(1);
            fe.EvaluateAll();

            XSSFCell formulaCell2 = (XSSFCell)sheet.GetColumn(copyColIndex).GetCell(0);

            Assert.AreEqual(copyColIndex + 1, anotherColumn.ColumnNum);
            Assert.AreEqual("POI", sheet.GetRow(0).GetCell(copyColIndex + 1).StringCellValue);
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(5, sheet.LastColumnNum);
            Assert.AreEqual(width, copyColumn.Width);
            Assert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            Assert.AreEqual("POI", formulaCell2.CellComment.Author);
            Assert.AreEqual("C1 + D1", formulaCell2.CellFormula);
            Assert.AreEqual(30, formulaCell2.NumericCellValue);
            Assert.AreEqual(2, sheet.NumMergedRegions);
            Assert.NotNull(originaMergedRegion);
            Assert.NotNull(copyMergedRegion);
            Assert.AreEqual(mergedRegionFirstRow, copyMergedRegion.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, copyMergedRegion.LastRow);
            Assert.AreEqual(copyColumn.ColumnNum, copyMergedRegion.FirstColumn);
            Assert.AreEqual(copyColumn.ColumnNum + 1, copyMergedRegion.LastColumn);

            FileInfo file = TempFile.CreateTempFile("ShiftCols-", ".xlsx");
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
            Assert.AreEqual(5, sheetLoaded.LastColumnNum);
            Assert.AreEqual(width, copyColumnLoaded.Width);
            Assert.AreEqual(10, sheetLoaded.GetColumn(originalColIndex + 1).GetCell(0).NumericCellValue);
            Assert.AreEqual(20, sheetLoaded.GetColumn(originalColIndex + 2).GetCell(0).NumericCellValue);
            Assert.AreEqual("POI", sheetLoaded.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            Assert.AreEqual("POI", formulaCellLoaded.CellComment.Author);
            Assert.AreEqual("C1 + D1", formulaCellLoaded.CellFormula);
            Assert.AreEqual(30, formulaCellLoaded.NumericCellValue);
            Assert.NotNull(originalMergedRegionLoaded);
            Assert.AreEqual(mergedRegionFirstRow, originalMergedRegionLoaded.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, originalMergedRegionLoaded.LastRow);
            Assert.AreEqual(originalColumnLoaded.ColumnNum, originalMergedRegionLoaded.FirstColumn);
            Assert.AreEqual(originalColumnLoaded.ColumnNum + 1, originalMergedRegionLoaded.LastColumn);
            Assert.NotNull(copyMergedRegionLoaded);
            Assert.AreEqual(mergedRegionFirstRow, copyMergedRegionLoaded.FirstRow);
            Assert.AreEqual(mergedRegionLastRow, copyMergedRegionLoaded.LastRow);
            Assert.AreEqual(copyColumnLoaded.ColumnNum, copyMergedRegionLoaded.FirstColumn);
            Assert.AreEqual(copyColumnLoaded.ColumnNum + 1, copyMergedRegionLoaded.LastColumn);
        }

        [Test]
        public void ShiftColumns_Shift2ColumnsWithCommentsAndFormulasAndOverlapingMergedRegion_ColumnsShiftedCommentsShiftedFormulasAreUpdatedMergedRegionRemoved()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(wb);
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            XSSFDrawing dg = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFComment comment = (XSSFComment)dg.CreateCellComment(new XSSFClientAnchor());
            comment.Author = "POI";
            comment.String = new XSSFRichTextString("A new comment");

            XSSFCell mergedRegionStart = (XSSFCell)sheet.CreateRow(10).CreateCell(0);
            mergedRegionStart.SetCellValue("MERGED");
            int mergedRegionIndex = sheet.AddMergedRegion(new CellRangeAddress(10, 11, 0, 10));
            CellRangeAddress mergedRegion = sheet.GetMergedRegion(mergedRegionIndex);

            XSSFCell formulaCell = (XSSFCell)sheet.CreateColumn(1).CreateCell(0);
            formulaCell.SetCellFormula("C1 + D1");
            XSSFCell cell1 = (XSSFCell)sheet.CreateColumn(2).CreateCell(0);
            cell1.SetCellValue(10);
            cell1.CellComment = comment;
            sheet.CreateColumn(3).CreateCell(0).SetCellValue(20);

            fe.EvaluateAll();

            Assert.AreEqual("POI", sheet.GetColumn(2).GetCell(0).CellComment.Author);
            Assert.AreEqual(30, formulaCell.NumericCellValue);
            Assert.AreEqual("MERGED", sheet.GetRow(10).GetCell(0).StringCellValue);
            Assert.AreEqual(1, sheet.NumMergedRegions);
            Assert.NotNull(mergedRegion);
            Assert.AreEqual("A11:K12", mergedRegion.FormatAsString());

            sheet.ShiftColumns(2, 3, 4);
            fe.EvaluateAll();

            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(7, sheet.LastColumnNum);
            Assert.AreEqual(10, sheet.GetColumn(6).GetCell(0).NumericCellValue);
            Assert.AreEqual(20, sheet.GetColumn(7).GetCell(0).NumericCellValue);
            cell1 = (XSSFCell)sheet.GetColumn(6).GetCell(0);
            Assert.AreEqual("POI", cell1.CellComment.Author);
            Assert.AreEqual("G1+H1", formulaCell.CellFormula);
            Assert.AreEqual(30, formulaCell.NumericCellValue);
            Assert.AreEqual("MERGED", sheet.GetRow(10).GetCell(0).StringCellValue);
            Assert.AreEqual(0, sheet.NumMergedRegions);

            FileInfo file = TempFile.CreateTempFile("ShiftCols-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFFormulaEvaluator feLoaded = new XSSFFormulaEvaluator(wbLoaded);
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            XSSFCell formulaCellLoaded = (XSSFCell)sheetLoaded.GetColumn(1).GetCell(0);

            feLoaded.EvaluateAll();

            Assert.IsNull(sheetLoaded.GetColumn(2));
            Assert.AreEqual(1, sheetLoaded.FirstColumnNum);
            Assert.AreEqual(7, sheetLoaded.LastColumnNum);
            Assert.AreEqual(10, sheetLoaded.GetColumn(6).GetCell(0).NumericCellValue);
            Assert.AreEqual(20, sheetLoaded.GetColumn(7).GetCell(0).NumericCellValue);
            Assert.AreEqual("POI", sheetLoaded.GetColumn(6).GetCell(0).CellComment.Author);
            Assert.AreEqual("G1+H1", formulaCellLoaded.CellFormula);
            Assert.AreEqual(30, formulaCellLoaded.NumericCellValue);
            Assert.AreEqual("MERGED", sheetLoaded.GetRow(10).GetCell(0).StringCellValue);
            Assert.AreEqual(0, sheetLoaded.NumMergedRegions);
        }

        [Test]
        public void ShiftColumns_Shift1ColumnWithCommentsAndFormulas_ColumnShiftedCommentShiftedFormulaUpdated()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(wb);
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            XSSFDrawing dg = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFComment comment = (XSSFComment)dg.CreateCellComment(new XSSFClientAnchor());
            comment.Author = "POI";
            comment.String = new XSSFRichTextString("A new comment");

            XSSFCell formulaCell = (XSSFCell)sheet.CreateColumn(1).CreateCell(0);
            formulaCell.SetCellFormula("C1 + D1");
            XSSFCell cell1 = (XSSFCell)sheet.CreateColumn(2).CreateCell(0);
            cell1.SetCellValue(10);
            cell1.CellComment = comment;
            sheet.CreateColumn(3).CreateCell(0).SetCellValue(20);

            fe.EvaluateAll();

            Assert.AreEqual("POI", sheet.GetColumn(2).GetCell(0).CellComment.Author);
            Assert.AreEqual(30, formulaCell.NumericCellValue);

            sheet.ShiftColumns(2, 2, 4);
            fe.EvaluateAll();

            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(6, sheet.LastColumnNum);
            Assert.AreEqual(10, sheet.GetColumn(6).GetCell(0).NumericCellValue);
            Assert.AreEqual(20, sheet.GetColumn(3).GetCell(0).NumericCellValue);
            cell1 = (XSSFCell)sheet.GetColumn(6).GetCell(0);
            Assert.AreEqual("POI", cell1.CellComment.Author);
            Assert.AreEqual("G1+D1", formulaCell.CellFormula);
            Assert.AreEqual(30, formulaCell.NumericCellValue);

            FileInfo file = TempFile.CreateTempFile("ShiftCols-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFFormulaEvaluator feLoaded = new XSSFFormulaEvaluator(wbLoaded);
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            XSSFCell formulaCellLoaded = (XSSFCell)sheetLoaded.GetColumn(1).GetCell(0);

            feLoaded.EvaluateAll();

            Assert.IsNull(sheetLoaded.GetColumn(2));
            Assert.AreEqual(1, sheetLoaded.FirstColumnNum);
            Assert.AreEqual(6, sheetLoaded.LastColumnNum);
            Assert.AreEqual(10, sheetLoaded.GetColumn(6).GetCell(0).NumericCellValue);
            Assert.AreEqual(20, sheetLoaded.GetColumn(3).GetCell(0).NumericCellValue);
            Assert.AreEqual("POI", sheetLoaded.GetColumn(6).GetCell(0).CellComment.Author);
            Assert.AreEqual("G1+D1", formulaCellLoaded.CellFormula);
            Assert.AreEqual(30, formulaCellLoaded.NumericCellValue);
        }

        /// <summary>
        /// This test trys to shift a column which doesn't exist, but there are
        /// cells that would fall into that column. The expected behaviour is 
        /// that system will create a column object, shift it updating formulas
        /// and moving cell values and comments and then will destroy created
        /// column objects, leaving cells in new places
        /// </summary>
        [Test]
        public void ShiftColumns_ShiftNonExistingColumnWithCellsAndComments_CellsAreShiftedNocolumsCreated()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(wb);
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            XSSFDrawing dg = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFComment comment = (XSSFComment)dg.CreateCellComment(new XSSFClientAnchor());
            comment.Author = "POI";
            comment.String = new XSSFRichTextString("A new comment");

            XSSFCell formulaCell = (XSSFCell)sheet.CreateColumn(1).CreateCell(0);
            formulaCell.SetCellFormula("C1 + D2");
            sheet.CreateColumn(2).CreateCell(0).SetCellValue(10);
            XSSFRow row = (XSSFRow)sheet.CreateRow(1);
            XSSFCell cell1 = (XSSFCell)row.CreateCell(3);
            cell1.SetCellValue(20);
            cell1.CellComment = comment;

            fe.EvaluateAll();

            Assert.AreEqual("POI", row.GetCell(3).CellComment.Author);
            Assert.AreEqual(30, formulaCell.NumericCellValue);
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(2, sheet.LastColumnNum);

            sheet.ShiftColumns(3, 3, 4);
            fe.EvaluateAll();

            Assert.AreEqual("POI", row.GetCell(7).CellComment.Author);
            Assert.AreEqual("C1+H2", formulaCell.CellFormula);
            Assert.AreEqual(30, formulaCell.NumericCellValue);
            Assert.AreEqual(1, sheet.FirstColumnNum);
            Assert.AreEqual(2, sheet.LastColumnNum);

            FileInfo file = TempFile.CreateTempFile("ShiftCols-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFFormulaEvaluator feLoaded = new XSSFFormulaEvaluator(wbLoaded);
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            XSSFRow rowLoaded = (XSSFRow)sheet.GetRow(1);
            XSSFCell formulaCellLoaded = (XSSFCell)sheet.GetRow(0).GetCell(1);

            feLoaded.EvaluateAll();

            Assert.AreEqual("POI", rowLoaded.GetCell(7).CellComment.Author);
            Assert.AreEqual("C1+H2", formulaCellLoaded.CellFormula);
            Assert.AreEqual(30, formulaCellLoaded.NumericCellValue);
            Assert.AreEqual(1, sheetLoaded.FirstColumnNum);
            Assert.AreEqual(2, sheetLoaded.LastColumnNum);
        }//*/
    }
}