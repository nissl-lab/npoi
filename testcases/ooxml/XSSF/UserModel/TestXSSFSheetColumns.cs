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
using NUnit.Framework;using NUnit.Framework.Legacy;
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

            ClassicAssert.AreEqual(10, sheet.FirstColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            ClassicAssert.AreEqual(10, sheetLoaded.FirstColumnNum);
        }

        [Test]
        public void LastColumnNum_ReturnsCorrectNumber()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            _ = sheet.CreateColumn(10);

            ClassicAssert.AreEqual(10, sheet.LastColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            ClassicAssert.AreEqual(10, sheetLoaded.LastColumnNum);
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

            ClassicAssert.AreEqual(0, sheet.FirstColumnNum);
            ClassicAssert.AreEqual(colNumber - 1, sheet.LastColumnNum);
            ClassicAssert.AreEqual(colNumber, sheet.PhysicalNumberOfColumns);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            ClassicAssert.AreEqual(0, sheetLoaded.FirstColumnNum);
            ClassicAssert.AreEqual(colNumber - 1, sheetLoaded.LastColumnNum);
            ClassicAssert.AreEqual(colNumber, sheetLoaded.PhysicalNumberOfColumns);
        }

        [Test]
        public void CreateColumn_AtIndex6_ColumnIsCreatedWithCorrectIndex()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int colIndex = 6;

            _ = sheet.CreateColumn(colIndex);

            IColumn column = sheet.GetColumn(colIndex);

            ClassicAssert.NotNull(column);
            ClassicAssert.AreEqual(colIndex, column.ColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(colIndex);

            ClassicAssert.NotNull(columnLoaded);
            ClassicAssert.AreEqual(colIndex, columnLoaded.ColumnNum);
        }

        [Test]
        public void CreateColumn_AtIndex6Twice_ColumnIsCreatedWithCorrectIndexAndNoCells()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
            int colIndex = 6;

            sheet.CreateColumn(colIndex).CreateCell(0).SetCellValue(1);
            IColumn column = sheet.GetColumn(colIndex);

            ClassicAssert.NotNull(column);
            ClassicAssert.AreEqual(colIndex, column.ColumnNum);
            ClassicAssert.AreEqual(1, column.GetCell(0).NumericCellValue);

            IColumn columnOverwrite = sheet.CreateColumn(colIndex);

            ClassicAssert.NotNull(columnOverwrite);
            ClassicAssert.AreEqual(colIndex, columnOverwrite.ColumnNum);
            ClassicAssert.IsNull(columnOverwrite.GetCell(0));

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(colIndex);

            ClassicAssert.NotNull(columnLoaded);
            ClassicAssert.AreEqual(colIndex, columnLoaded.ColumnNum);
            ClassicAssert.IsNull(columnLoaded.GetCell(0));
        }

        [Test]
        public void GetColumn_GetExistingColumn_ReturnsAColumn()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            _ = sheet.CreateColumn(10);
            IColumn column = sheet.GetColumn(10);

            ClassicAssert.NotNull(column);
            ClassicAssert.AreEqual(10, column.ColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            IColumn columnLoaded = sheetLoaded.GetColumn(10);

            ClassicAssert.NotNull(columnLoaded);
            ClassicAssert.AreEqual(10, columnLoaded.ColumnNum);
        }

        [Test]
        public void GetColumn_GetNonExistingColumn_ReturnsNull()
        {
            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            using(XSSFWorkbook wb = new XSSFWorkbook())
            {
                XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");
                IColumn column = sheet.GetColumn(10);
                ClassicAssert.IsNull(column);

                using(Stream output = File.OpenWrite(file.FullName))
                {
                    wb.Write(output);
                }
            }

            using(XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString()))
            {
                XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
                IColumn columnLoaded = sheetLoaded.GetColumn(10);
                ClassicAssert.IsNull(columnLoaded);
            }
        }

        [Test]
        public void RemoveColumn_RemoveExistingColumn_ColumnIsRemoved()
        {
            FileInfo file;
            using(XSSFWorkbook wb = new XSSFWorkbook())
            {
                XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

                _ = sheet.CreateColumn(1);
                _ = sheet.CreateColumn(2);
                _ = sheet.CreateColumn(3);

                ClassicAssert.NotNull(sheet.GetColumn(1));
                ClassicAssert.NotNull(sheet.GetColumn(2));
                ClassicAssert.NotNull(sheet.GetColumn(3));
                ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
                ClassicAssert.AreEqual(3, sheet.LastColumnNum);

                sheet.RemoveColumn(sheet.GetColumn(3));

                ClassicAssert.IsNull(sheet.GetColumn(3));
                ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
                ClassicAssert.AreEqual(2, sheet.LastColumnNum);

                _ = sheet.CreateColumn(3);

                sheet.RemoveColumn(sheet.GetColumn(1));

                ClassicAssert.IsNull(sheet.GetColumn(1));
                ClassicAssert.AreEqual(2, sheet.FirstColumnNum);
                ClassicAssert.AreEqual(3, sheet.LastColumnNum);

                _ = sheet.CreateColumn(1);

                ClassicAssert.NotNull(sheet.GetColumn(3));
                ClassicAssert.AreEqual(3, sheet.LastColumnNum);

                sheet.RemoveColumn(sheet.GetColumn(2));

                ClassicAssert.IsNull(sheet.GetColumn(2));
                ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
                ClassicAssert.AreEqual(3, sheet.LastColumnNum);

                file = TempFile.CreateTempFile("poi-", ".xlsx");
                using(Stream output = File.OpenWrite(file.FullName))
                {
                    wb.Write(output);
                }
            }
            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");

            ClassicAssert.IsNull(sheetLoaded.GetColumn(2));
            ClassicAssert.AreEqual(1, sheetLoaded.FirstColumnNum);
            ClassicAssert.AreEqual(3, sheetLoaded.LastColumnNum);
        }

        [Test]
        public void RemoveColumn_RemoveNonExistingColumn_ExceptionIsThrown()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("sheet1");

            _ = sheet.CreateColumn(1);
            _ = sheet.CreateColumn(2);
            _ = sheet.CreateColumn(3);

            ClassicAssert.NotNull(sheet.GetColumn(1));
            ClassicAssert.NotNull(sheet.GetColumn(2));
            ClassicAssert.NotNull(sheet.GetColumn(3));
            ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
            ClassicAssert.AreEqual(3, sheet.LastColumnNum);

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

            ClassicAssert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(originalColIndex).StringCellValue);
            ClassicAssert.AreEqual("POI", formulaCell1.CellComment.Author);
            ClassicAssert.AreEqual(30, formulaCell1.NumericCellValue);
            ClassicAssert.AreEqual(1, sheet.NumMergedRegions);
            ClassicAssert.NotNull(originaMergedRegion);

            XSSFColumn copyColumn = (XSSFColumn)sheet.CopyColumn(originalColIndex, copyColIndex);
            CellRangeAddress copyMergedRegion = sheet.GetMergedRegion(1);
            fe.EvaluateAll();

            XSSFCell formulaCell2 = (XSSFCell)sheet.GetColumn(copyColIndex).GetCell(0);

            ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
            ClassicAssert.AreEqual(4, sheet.LastColumnNum);
            ClassicAssert.AreEqual(width, copyColumn.Width);
            //ClassicAssert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            ClassicAssert.AreEqual("POI", formulaCell2.CellComment.Author);
            ClassicAssert.AreEqual("C1 + D1", formulaCell2.CellFormula);
            ClassicAssert.AreEqual(30, formulaCell2.NumericCellValue);
            ClassicAssert.AreEqual(2, sheet.NumMergedRegions);
            ClassicAssert.NotNull(originaMergedRegion);
            ClassicAssert.NotNull(copyMergedRegion);
            ClassicAssert.AreEqual(mergedRegionFirstRow, copyMergedRegion.FirstRow);
            ClassicAssert.AreEqual(mergedRegionLastRow, copyMergedRegion.LastRow);
            ClassicAssert.AreEqual(copyColumn.ColumnNum, copyMergedRegion.FirstColumn);
            ClassicAssert.AreEqual(copyColumn.ColumnNum + 1, copyMergedRegion.LastColumn);

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

            ClassicAssert.AreEqual(1, sheetLoaded.FirstColumnNum);
            ClassicAssert.AreEqual(4, sheetLoaded.LastColumnNum);
            ClassicAssert.AreEqual(width, copyColumnLoaded.Width);
            ClassicAssert.AreEqual(10, sheetLoaded.GetColumn(originalColIndex + 1).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual(20, sheetLoaded.GetColumn(originalColIndex + 2).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual("POI", sheetLoaded.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            ClassicAssert.AreEqual("POI", formulaCellLoaded.CellComment.Author);
            ClassicAssert.AreEqual("C1 + D1", formulaCellLoaded.CellFormula);
            ClassicAssert.AreEqual(30, formulaCellLoaded.NumericCellValue);
            ClassicAssert.NotNull(originalMergedRegionLoaded);
            ClassicAssert.AreEqual(mergedRegionFirstRow, originalMergedRegionLoaded.FirstRow);
            ClassicAssert.AreEqual(mergedRegionLastRow, originalMergedRegionLoaded.LastRow);
            ClassicAssert.AreEqual(originalColumnLoaded.ColumnNum, originalMergedRegionLoaded.FirstColumn);
            ClassicAssert.AreEqual(originalColumnLoaded.ColumnNum + 1, originalMergedRegionLoaded.LastColumn);
            ClassicAssert.NotNull(copyMergedRegionLoaded);
            ClassicAssert.AreEqual(mergedRegionFirstRow, copyMergedRegionLoaded.FirstRow);
            ClassicAssert.AreEqual(mergedRegionLastRow, copyMergedRegionLoaded.LastRow);
            ClassicAssert.AreEqual(copyColumnLoaded.ColumnNum, copyMergedRegionLoaded.FirstColumn);
            ClassicAssert.AreEqual(copyColumnLoaded.ColumnNum + 1, copyMergedRegionLoaded.LastColumn);
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

            ClassicAssert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(originalColIndex).StringCellValue);
            ClassicAssert.AreEqual("POI", formulaCell1.CellComment.Author);
            ClassicAssert.AreEqual(30, formulaCell1.NumericCellValue);
            ClassicAssert.AreEqual(1, sheet.NumMergedRegions);
            ClassicAssert.NotNull(originaMergedRegion);

            XSSFColumn copyColumn = (XSSFColumn)sheet.CopyColumn(originalColIndex, copyColIndex);
            CellRangeAddress copyMergedRegion = sheet.GetMergedRegion(1);
            fe.EvaluateAll();

            XSSFCell formulaCell2 = (XSSFCell)sheet.GetColumn(copyColIndex).GetCell(0);

            ClassicAssert.AreEqual(copyColIndex + 1, anotherColumn.ColumnNum);
            ClassicAssert.AreEqual("POI", sheet.GetRow(0).GetCell(copyColIndex + 1).StringCellValue);
            ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
            ClassicAssert.AreEqual(5, sheet.LastColumnNum);
            ClassicAssert.AreEqual(width, copyColumn.Width);
            ClassicAssert.AreEqual("POI", sheet.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            ClassicAssert.AreEqual("POI", formulaCell2.CellComment.Author);
            ClassicAssert.AreEqual("C1 + D1", formulaCell2.CellFormula);
            ClassicAssert.AreEqual(30, formulaCell2.NumericCellValue);
            ClassicAssert.AreEqual(2, sheet.NumMergedRegions);
            ClassicAssert.NotNull(originaMergedRegion);
            ClassicAssert.NotNull(copyMergedRegion);
            ClassicAssert.AreEqual(mergedRegionFirstRow, copyMergedRegion.FirstRow);
            ClassicAssert.AreEqual(mergedRegionLastRow, copyMergedRegion.LastRow);
            ClassicAssert.AreEqual(copyColumn.ColumnNum, copyMergedRegion.FirstColumn);
            ClassicAssert.AreEqual(copyColumn.ColumnNum + 1, copyMergedRegion.LastColumn);

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

            ClassicAssert.AreEqual(1, sheetLoaded.FirstColumnNum);
            ClassicAssert.AreEqual(5, sheetLoaded.LastColumnNum);
            ClassicAssert.AreEqual(width, copyColumnLoaded.Width);
            ClassicAssert.AreEqual(10, sheetLoaded.GetColumn(originalColIndex + 1).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual(20, sheetLoaded.GetColumn(originalColIndex + 2).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual("POI", sheetLoaded.GetRow(mergedRegionFirstRow).GetCell(copyColIndex).StringCellValue);
            ClassicAssert.AreEqual("POI", formulaCellLoaded.CellComment.Author);
            ClassicAssert.AreEqual("C1 + D1", formulaCellLoaded.CellFormula);
            ClassicAssert.AreEqual(30, formulaCellLoaded.NumericCellValue);
            ClassicAssert.NotNull(originalMergedRegionLoaded);
            ClassicAssert.AreEqual(mergedRegionFirstRow, originalMergedRegionLoaded.FirstRow);
            ClassicAssert.AreEqual(mergedRegionLastRow, originalMergedRegionLoaded.LastRow);
            ClassicAssert.AreEqual(originalColumnLoaded.ColumnNum, originalMergedRegionLoaded.FirstColumn);
            ClassicAssert.AreEqual(originalColumnLoaded.ColumnNum + 1, originalMergedRegionLoaded.LastColumn);
            ClassicAssert.NotNull(copyMergedRegionLoaded);
            ClassicAssert.AreEqual(mergedRegionFirstRow, copyMergedRegionLoaded.FirstRow);
            ClassicAssert.AreEqual(mergedRegionLastRow, copyMergedRegionLoaded.LastRow);
            ClassicAssert.AreEqual(copyColumnLoaded.ColumnNum, copyMergedRegionLoaded.FirstColumn);
            ClassicAssert.AreEqual(copyColumnLoaded.ColumnNum + 1, copyMergedRegionLoaded.LastColumn);
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

            ClassicAssert.AreEqual("POI", sheet.GetColumn(2).GetCell(0).CellComment.Author);
            ClassicAssert.AreEqual(30, formulaCell.NumericCellValue);
            ClassicAssert.AreEqual("MERGED", sheet.GetRow(10).GetCell(0).StringCellValue);
            ClassicAssert.AreEqual(1, sheet.NumMergedRegions);
            ClassicAssert.NotNull(mergedRegion);
            ClassicAssert.AreEqual("A11:K12", mergedRegion.FormatAsString());

            sheet.ShiftColumns(2, 3, 4);
            fe.EvaluateAll();

            ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
            ClassicAssert.AreEqual(7, sheet.LastColumnNum);
            ClassicAssert.AreEqual(10, sheet.GetColumn(6).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual(20, sheet.GetColumn(7).GetCell(0).NumericCellValue);
            cell1 = (XSSFCell)sheet.GetColumn(6).GetCell(0);
            ClassicAssert.AreEqual("POI", cell1.CellComment.Author);
            ClassicAssert.AreEqual("G1+H1", formulaCell.CellFormula);
            ClassicAssert.AreEqual(30, formulaCell.NumericCellValue);
            ClassicAssert.AreEqual("MERGED", sheet.GetRow(10).GetCell(0).StringCellValue);
            ClassicAssert.AreEqual(0, sheet.NumMergedRegions);

            FileInfo file = TempFile.CreateTempFile("ShiftCols-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFFormulaEvaluator feLoaded = new XSSFFormulaEvaluator(wbLoaded);
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            XSSFCell formulaCellLoaded = (XSSFCell)sheetLoaded.GetColumn(1).GetCell(0);

            feLoaded.EvaluateAll();

            ClassicAssert.IsNull(sheetLoaded.GetColumn(2));
            ClassicAssert.AreEqual(1, sheetLoaded.FirstColumnNum);
            ClassicAssert.AreEqual(7, sheetLoaded.LastColumnNum);
            ClassicAssert.AreEqual(10, sheetLoaded.GetColumn(6).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual(20, sheetLoaded.GetColumn(7).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual("POI", sheetLoaded.GetColumn(6).GetCell(0).CellComment.Author);
            ClassicAssert.AreEqual("G1+H1", formulaCellLoaded.CellFormula);
            ClassicAssert.AreEqual(30, formulaCellLoaded.NumericCellValue);
            ClassicAssert.AreEqual("MERGED", sheetLoaded.GetRow(10).GetCell(0).StringCellValue);
            ClassicAssert.AreEqual(0, sheetLoaded.NumMergedRegions);
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

            ClassicAssert.AreEqual("POI", sheet.GetColumn(2).GetCell(0).CellComment.Author);
            ClassicAssert.AreEqual(30, formulaCell.NumericCellValue);

            sheet.ShiftColumns(2, 2, 4);
            fe.EvaluateAll();

            ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
            ClassicAssert.AreEqual(6, sheet.LastColumnNum);
            ClassicAssert.AreEqual(10, sheet.GetColumn(6).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual(20, sheet.GetColumn(3).GetCell(0).NumericCellValue);
            cell1 = (XSSFCell)sheet.GetColumn(6).GetCell(0);
            ClassicAssert.AreEqual("POI", cell1.CellComment.Author);
            ClassicAssert.AreEqual("G1+D1", formulaCell.CellFormula);
            ClassicAssert.AreEqual(30, formulaCell.NumericCellValue);

            FileInfo file = TempFile.CreateTempFile("ShiftCols-", ".xlsx");
            Stream output = File.OpenWrite(file.FullName);
            wb.Write(output);
            output.Close();

            XSSFWorkbook wbLoaded = new XSSFWorkbook(file.ToString());
            XSSFFormulaEvaluator feLoaded = new XSSFFormulaEvaluator(wbLoaded);
            XSSFSheet sheetLoaded = (XSSFSheet)wbLoaded.GetSheet("sheet1");
            XSSFCell formulaCellLoaded = (XSSFCell)sheetLoaded.GetColumn(1).GetCell(0);

            feLoaded.EvaluateAll();

            ClassicAssert.IsNull(sheetLoaded.GetColumn(2));
            ClassicAssert.AreEqual(1, sheetLoaded.FirstColumnNum);
            ClassicAssert.AreEqual(6, sheetLoaded.LastColumnNum);
            ClassicAssert.AreEqual(10, sheetLoaded.GetColumn(6).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual(20, sheetLoaded.GetColumn(3).GetCell(0).NumericCellValue);
            ClassicAssert.AreEqual("POI", sheetLoaded.GetColumn(6).GetCell(0).CellComment.Author);
            ClassicAssert.AreEqual("G1+D1", formulaCellLoaded.CellFormula);
            ClassicAssert.AreEqual(30, formulaCellLoaded.NumericCellValue);
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

            ClassicAssert.AreEqual("POI", row.GetCell(3).CellComment.Author);
            ClassicAssert.AreEqual(30, formulaCell.NumericCellValue);
            ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
            ClassicAssert.AreEqual(2, sheet.LastColumnNum);

            sheet.ShiftColumns(3, 3, 4);
            fe.EvaluateAll();

            ClassicAssert.AreEqual("POI", row.GetCell(7).CellComment.Author);
            ClassicAssert.AreEqual("C1+H2", formulaCell.CellFormula);
            ClassicAssert.AreEqual(30, formulaCell.NumericCellValue);
            ClassicAssert.AreEqual(1, sheet.FirstColumnNum);
            ClassicAssert.AreEqual(2, sheet.LastColumnNum);

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

            ClassicAssert.AreEqual("POI", rowLoaded.GetCell(7).CellComment.Author);
            ClassicAssert.AreEqual("C1+H2", formulaCellLoaded.CellFormula);
            ClassicAssert.AreEqual(30, formulaCellLoaded.NumericCellValue);
            ClassicAssert.AreEqual(1, sheetLoaded.FirstColumnNum);
            ClassicAssert.AreEqual(2, sheetLoaded.LastColumnNum);
        }//*/
    }
}