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
using NUnit.Framework;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace NPOI.XSSF.Model
{


    [TestFixture]
    public class TestCommentsTable
    {

        private static String TEST_A2_TEXT = "test A2 text";
        private static String TEST_A1_TEXT = "test A1 text";
        private static String TEST_AUTHOR = "test author";
        [Test]
        public void TestFindAuthor()
        {
            CommentsTable sheetComments = new CommentsTable();
            Assert.AreEqual(1, sheetComments.GetNumberOfAuthors());
            Assert.AreEqual(0, sheetComments.FindAuthor(""));
            Assert.AreEqual("", sheetComments.GetAuthor(0));

            Assert.AreEqual(1, sheetComments.FindAuthor(TEST_AUTHOR));
            Assert.AreEqual(2, sheetComments.FindAuthor("another author"));
            Assert.AreEqual(1, sheetComments.FindAuthor(TEST_AUTHOR));
            Assert.AreEqual(3, sheetComments.FindAuthor("YAA"));
            Assert.AreEqual(2, sheetComments.FindAuthor("another author"));
        }
        [Test]
        public void TestGetCellComment()
        {
            CommentsTable sheetComments = new CommentsTable();

            CT_Comments comments = sheetComments.GetCTComments();
            CT_CommentList commentList = comments.commentList;

            // Create 2 comments for A1 and A" cells
            CT_Comment comment0 = commentList.InsertNewComment(0);
            comment0.@ref = "A1";
            CT_Rst ctrst0 = new CT_Rst();
            ctrst0.t = (TEST_A1_TEXT);
            comment0.text = (ctrst0);
            CT_Comment comment1 = commentList.InsertNewComment(0);
            comment1.@ref = ("A2");
            CT_Rst ctrst1 = new CT_Rst();
            ctrst1.t = (TEST_A2_TEXT);
            comment1.text = (ctrst1);

            // Test Finding the right comment for a cell
            Assert.AreSame(comment0, sheetComments.GetCTComment("A1"));
            Assert.AreSame(comment1, sheetComments.GetCTComment("A2"));
            Assert.IsNull(sheetComments.GetCTComment("A3"));
        }

        [Test]
        public void TestExisting()
        {
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("WithVariousData.xlsx");
            ISheet sheet1 = workbook.GetSheetAt(0);
            ISheet sheet2 = workbook.GetSheetAt(1);

            Assert.IsTrue(((XSSFSheet)sheet1).HasComments);
            Assert.IsFalse(((XSSFSheet)sheet2).HasComments);

            // Comments should be in C5 and C7
            IRow r5 = sheet1.GetRow(4);
            IRow r7 = sheet1.GetRow(6);
            Assert.IsNotNull(r5.GetCell(2).CellComment);
            Assert.IsNotNull(r7.GetCell(2).CellComment);

            // Check they have what we expect
            // TODO: Rich text formatting
            IComment cc5 = r5.GetCell(2).CellComment;
            IComment cc7 = r7.GetCell(2).CellComment;

            Assert.AreEqual("Nick Burch", cc5.Author);
            Assert.AreEqual("Nick Burch:\nThis is a comment", cc5.String.String);
            Assert.AreEqual(4, cc5.Row);
            Assert.AreEqual(2, cc5.Column);

            Assert.AreEqual("Nick Burch", cc7.Author);
            Assert.AreEqual("Nick Burch:\nComment #1\n", cc7.String.String);
            Assert.AreEqual(6, cc7.Row);
            Assert.AreEqual(2, cc7.Column);
        }
        [Test]
        public void TestWriteRead()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("WithVariousData.xlsx");
            XSSFSheet sheet1 = (XSSFSheet)workbook.GetSheetAt(0);
            XSSFSheet sheet2 = (XSSFSheet)workbook.GetSheetAt(1);

            Assert.IsTrue(sheet1.HasComments);
            Assert.IsFalse(sheet2.HasComments);

            // Change on comment on sheet 1, and add another into
            //  sheet 2
            IRow r5 = sheet1.GetRow(4);
            IComment cc5 = r5.GetCell(2).CellComment;
            cc5.Author = ("Apache POI");
            cc5.String = (new XSSFRichTextString("Hello!"));

            IRow r2s2 = sheet2.CreateRow(2);
            ICell c1r2s2 = r2s2.CreateCell(1);
            Assert.IsNull(c1r2s2.CellComment);

            IDrawing dg = sheet2.CreateDrawingPatriarch();
            IComment cc2 = dg.CreateCellComment(new XSSFClientAnchor());
            cc2.Author = ("Also POI");
            cc2.String = (new XSSFRichTextString("A new comment"));
            c1r2s2.CellComment = (cc2);


            // Save, and re-load the file
            workbook = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);

            // Check we still have comments where we should do
            sheet1 = (XSSFSheet)workbook.GetSheetAt(0);
            sheet2 = (XSSFSheet)workbook.GetSheetAt(1);
            Assert.IsNotNull(sheet1.GetRow(4).GetCell(2).CellComment);
            Assert.IsNotNull(sheet1.GetRow(6).GetCell(2).CellComment);
            Assert.IsNotNull(sheet2.GetRow(2).GetCell(1).CellComment);

            // And check they still have the contents they should do
            Assert.AreEqual("Apache POI",
                    sheet1.GetRow(4).GetCell(2).CellComment.Author);
            Assert.AreEqual("Nick Burch",
                    sheet1.GetRow(6).GetCell(2).CellComment.Author);
            Assert.AreEqual("Also POI",
                    sheet2.GetRow(2).GetCell(1).CellComment.Author);

            Assert.AreEqual("Hello!",
                    sheet1.GetRow(4).GetCell(2).CellComment.String.String);
        }
        [Test]
        public void TestReadWriteMultipleAuthors()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("WithMoreVariousData.xlsx");
            XSSFSheet sheet1 = (XSSFSheet)workbook.GetSheetAt(0);
            XSSFSheet sheet2 = (XSSFSheet)workbook.GetSheetAt(1);

            Assert.IsTrue(sheet1.HasComments);
            Assert.IsFalse(sheet2.HasComments);

            Assert.AreEqual("Nick Burch",
                    sheet1.GetRow(4).GetCell(2).CellComment.Author);
            Assert.AreEqual("Nick Burch",
                    sheet1.GetRow(6).GetCell(2).CellComment.Author);
            Assert.AreEqual("Torchbox",
                    sheet1.GetRow(12).GetCell(2).CellComment.Author);

            // Save, and re-load the file
            workbook = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);

            // Check we still have comments where we should do
            sheet1 = (XSSFSheet)workbook.GetSheetAt(0);
            Assert.IsNotNull(sheet1.GetRow(4).GetCell(2).CellComment);
            Assert.IsNotNull(sheet1.GetRow(6).GetCell(2).CellComment);
            Assert.IsNotNull(sheet1.GetRow(12).GetCell(2).CellComment);

            // And check they still have the contents they should do
            Assert.AreEqual("Nick Burch",
                    sheet1.GetRow(4).GetCell(2).CellComment.Author);
            Assert.AreEqual("Nick Burch",
                    sheet1.GetRow(6).GetCell(2).CellComment.Author);
            Assert.AreEqual("Torchbox",
                    sheet1.GetRow(12).GetCell(2).CellComment.Author);

            // Todo - check text too, once bug fixed
        }
        [Test]
        public void TestRemoveComment()
        {
            CommentsTable sheetComments = new CommentsTable();
            CT_Comment a1 = sheetComments.CreateComment();
            a1.@ref = ("A1");
            CT_Comment a2 = sheetComments.CreateComment();
            a2.@ref = ("A2");
            CT_Comment a3 = sheetComments.CreateComment();
            a3.@ref = ("A3");

            Assert.AreSame(a1, sheetComments.GetCTComment("A1"));
            Assert.AreSame(a2, sheetComments.GetCTComment("A2"));
            Assert.AreSame(a3, sheetComments.GetCTComment("A3"));
            Assert.AreEqual(3, sheetComments.GetNumberOfComments());

            Assert.IsTrue(sheetComments.RemoveComment("A1"));
            Assert.AreEqual(2, sheetComments.GetNumberOfComments());
            Assert.IsNull(sheetComments.GetCTComment("A1"));
            Assert.AreSame(a2, sheetComments.GetCTComment("A2"));
            Assert.AreSame(a3, sheetComments.GetCTComment("A3"));

            Assert.IsTrue(sheetComments.RemoveComment("A2"));
            Assert.AreEqual(1, sheetComments.GetNumberOfComments());
            Assert.IsNull(sheetComments.GetCTComment("A1"));
            Assert.IsNull(sheetComments.GetCTComment("A2"));
            Assert.AreSame(a3, sheetComments.GetCTComment("A3"));

            Assert.IsTrue(sheetComments.RemoveComment("A3"));
            Assert.AreEqual(0, sheetComments.GetNumberOfComments());
            Assert.IsNull(sheetComments.GetCTComment("A1"));
            Assert.IsNull(sheetComments.GetCTComment("A2"));
            Assert.IsNull(sheetComments.GetCTComment("A3"));
        }
    }

}