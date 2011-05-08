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

namespace TestCases.SS.UserModel
{
    using System;
    using NPOI.SS;
    using NPOI.SS.UserModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     * Common superclass for testing implementatiosn of
     * {@link Comment}
     */
    public abstract class BaseTestCellComment
    {

        private ITestDataProvider _testDataProvider;

        protected BaseTestCellComment(ITestDataProvider testDataProvider) {
            _testDataProvider = testDataProvider;
        }
        [TestMethod]
        public void TestFind()
        {
            Workbook book = _testDataProvider.CreateWorkbook();
            Sheet sheet = book.CreateSheet();
            Assert.IsNull(sheet.GetCellComment(0, 0));

            Row row = sheet.CreateRow(0);
            Cell cell = row.CreateCell(0);
            Assert.IsNull(sheet.GetCellComment(0, 0));
            Assert.IsNull(cell.CellComment);
        }
        [TestMethod]
        public void TestCreate()
        {
            String cellText = "Hello, World";
            String commentText = "We can set comments in POI";
            String commentAuthor = "Apache Software Foundation";
            int cellRow = 3;
            int cellColumn = 1;

            Workbook wb = _testDataProvider.CreateWorkbook();
            CreationHelper factory = wb.GetCreationHelper();

            Sheet sheet = wb.CreateSheet();
            Assert.IsNull(sheet.GetCellComment(cellRow, cellColumn));

            Cell cell = sheet.CreateRow(cellRow).CreateCell(cellColumn);
            cell.SetCellValue(factory.CreateRichTextString(cellText));
            Assert.IsNull(cell.CellComment);
            Assert.IsNull(sheet.GetCellComment(cellRow, cellColumn));

            Drawing patr = sheet.CreateDrawingPatriarch();
            ClientAnchor anchor = factory.CreateClientAnchor();
            anchor.Col1=(2);
            anchor.Col2=(5);
            anchor.Row1=(1);
            anchor.Row2=(2);
            Comment comment = patr.CreateCellComment(anchor);
            Assert.IsFalse(comment.Visible);
            comment.Visible = (true);
            Assert.IsTrue(comment.Visible);
            RichTextString string1 = factory.CreateRichTextString(commentText);
            comment.String=(string1);
            comment.Author=(commentAuthor);
            cell.CellComment=(comment);
            Assert.IsNotNull(cell.CellComment);
            Assert.IsNotNull(sheet.GetCellComment(cellRow, cellColumn));

            //verify our Settings
            Assert.AreEqual(commentAuthor, comment.Author);
            Assert.AreEqual(commentText, comment.String.String);
            Assert.AreEqual(cellRow, comment.Row);
            Assert.AreEqual(cellColumn, comment.Column);

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);
            cell = sheet.GetRow(cellRow).GetCell(cellColumn);
            comment = cell.CellComment;

            Assert.IsNotNull(comment);
            Assert.AreEqual(commentAuthor, comment.Author);
            Assert.AreEqual(commentText, comment.String.String);
            Assert.AreEqual(cellRow, comment.Row);
            Assert.AreEqual(cellColumn, comment.Column);
            Assert.IsTrue(comment.Visible);

            // Change slightly, and re-test
            comment.String = (factory.CreateRichTextString("New Comment Text"));
            comment.Visible = (false);

            wb = _testDataProvider.WriteOutAndReadBack(wb);

            sheet = wb.GetSheetAt(0);
            cell = sheet.GetRow(cellRow).GetCell(cellColumn);
            comment = cell.CellComment;

            Assert.IsNotNull(comment);
            Assert.AreEqual(commentAuthor, comment.Author);
            Assert.AreEqual("New Comment Text", comment.String.String);
            Assert.AreEqual(cellRow, comment.Row);
            Assert.AreEqual(cellColumn, comment.Column);
            Assert.IsFalse(comment.Visible);
        }

        /**
         * test that we can read cell comments from an existing workbook.
         */
        [TestMethod]
        public void TestReadComments()
        {
            Workbook wb = _testDataProvider.OpenSampleWorkbook("SimpleWithComments." + _testDataProvider.StandardFileNameExtension);

            Sheet sheet = wb.GetSheetAt(0);

            Cell cell;
            Row row;
            Comment comment;

            for (int rownum = 0; rownum < 3; rownum++)
            {
                row = sheet.GetRow(rownum);
                cell = row.GetCell(0);
                comment = cell.CellComment;
                Assert.IsNull(comment, "Cells in the first column are not commented");
                Assert.IsNull(sheet.GetCellComment(rownum, 0));
            }

            for (int rownum = 0; rownum < 3; rownum++)
            {
                row = sheet.GetRow(rownum);
                cell = row.GetCell(1);
                comment = cell.CellComment;
                Assert.IsNotNull(comment, "Cells in the second column have comments");
                Assert.IsNotNull(sheet.GetCellComment(rownum, 1), "Cells in the second column have comments");

                Assert.AreEqual("Yegor Kozlov", comment.Author);
                Assert.IsFalse(comment.String.String == string.Empty, "cells in the second column have not empyy notes");
                Assert.AreEqual(rownum, comment.Row);
                Assert.AreEqual(cell.ColumnIndex, comment.Column);
            }
        }

        /**
         * test that we can modify existing cell comments
         */
        [TestMethod]
        public void TestModifyComments()
        {

            Workbook wb = _testDataProvider.OpenSampleWorkbook("SimpleWithComments." + _testDataProvider.StandardFileNameExtension);
            CreationHelper factory = wb.GetCreationHelper();

            Sheet sheet = wb.GetSheetAt(0);

            Cell cell;
            Row row;
            Comment comment;

            for (int rownum = 0; rownum < 3; rownum++)
            {
                row = sheet.GetRow(rownum);
                cell = row.GetCell(1);
                comment = cell.CellComment;
                comment.Author = ("Mofified[" + rownum + "] by Yegor");
                comment.String = (factory.CreateRichTextString("Modified comment at row " + rownum));
            }

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);

            for (int rownum = 0; rownum < 3; rownum++)
            {
                row = sheet.GetRow(rownum);
                cell = row.GetCell(1);
                comment = cell.CellComment;

                Assert.AreEqual("Mofified[" + rownum + "] by Yegor", comment.Author);
                Assert.AreEqual("Modified comment at row " + rownum, comment.String.String);
            }

        }
        [TestMethod]
        public void TestDeleteComments()
        {
            Workbook wb = _testDataProvider.OpenSampleWorkbook("SimpleWithComments." + _testDataProvider.StandardFileNameExtension);
            Sheet sheet = wb.GetSheetAt(0);

            // Zap from rows 1 and 3
            Assert.IsNotNull(sheet.GetRow(0).GetCell(1).CellComment);
            Assert.IsNotNull(sheet.GetRow(1).GetCell(1).CellComment);
            Assert.IsNotNull(sheet.GetRow(2).GetCell(1).CellComment);

            sheet.GetRow(0).GetCell(1).RemoveCellComment();
            sheet.GetRow(2).GetCell(1).CellComment = (null);

            // Check gone so far
            Assert.IsNull(sheet.GetRow(0).GetCell(1).CellComment);
            Assert.IsNotNull(sheet.GetRow(1).GetCell(1).CellComment);
            Assert.IsNull(sheet.GetRow(2).GetCell(1).CellComment);

            // Save and re-load
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);
            // Check
            Assert.IsNull(sheet.GetRow(0).GetCell(1).CellComment);
            Assert.IsNotNull(sheet.GetRow(1).GetCell(1).CellComment);
            Assert.IsNull(sheet.GetRow(2).GetCell(1).CellComment);

        }

        /**
         * code from the quick guide
         */
        [TestMethod]
        public void TestQuickGuide()
        {
            Workbook wb = _testDataProvider.CreateWorkbook();

            CreationHelper factory = wb.GetCreationHelper();

            Sheet sheet = wb.CreateSheet();

            Cell cell = sheet.CreateRow(3).CreateCell(5);
            cell.SetCellValue("F4");

            Drawing drawing = sheet.CreateDrawingPatriarch();

            ClientAnchor anchor = factory.CreateClientAnchor();
            Comment comment = drawing.CreateCellComment(anchor);
            RichTextString str = factory.CreateRichTextString("Hello, World!");
            comment.String = (str);
            comment.Author = ("Apache POI");
            //assign the comment to the cell
            cell.CellComment = (comment);

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0);
            cell = sheet.GetRow(3).GetCell(5);
            comment = cell.CellComment;
            Assert.IsNotNull(comment);
            Assert.AreEqual("Hello, World!", comment.String.String);
            Assert.AreEqual("Apache POI", comment.Author);
            Assert.AreEqual(3, comment.Row);
            Assert.AreEqual(5, comment.Column);
        }
    }
}



