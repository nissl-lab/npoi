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
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF;
using NPOI.XSSF.Model;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using TestCases.SS.UserModel;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFComment : BaseTestCellComment
    {

        private static String TEST_RICHTEXTSTRING = "test richtextstring";

        public TestXSSFComment()
            : base(XSSFITestDataProvider.instance)
        {

        }

        /**
         * Test properties of a newly constructed comment
         */
        [Test]
        public void Constructor()
        {
            CommentsTable sheetComments = new CommentsTable();
            Assert.IsNotNull(sheetComments.GetCTComments().commentList);
            Assert.IsNotNull(sheetComments.GetCTComments().authors);
            Assert.AreEqual(1, sheetComments.GetCTComments().authors.SizeOfAuthorArray());
            Assert.AreEqual(1, sheetComments.GetNumberOfAuthors());

            CT_Comment ctComment = sheetComments.NewComment(CellAddress.A1);
            CT_Shape vmlShape = new CT_Shape();

            XSSFComment comment = new XSSFComment(sheetComments, ctComment, vmlShape);
            Assert.AreEqual(null, comment.String.String);
            Assert.AreEqual(0, comment.Row);
            Assert.AreEqual(0, comment.Column);
            Assert.AreEqual("", comment.Author);
            Assert.AreEqual(false, comment.Visible);
        }
        [Test]
        public void GetSetCol()
        {
            CommentsTable sheetComments = new CommentsTable();
            XSSFVMLDrawing vml = new XSSFVMLDrawing();
            CT_Comment ctComment = sheetComments.NewComment(CellAddress.A1);
            CT_Shape vmlShape = vml.newCommentShape();

            XSSFComment comment = new XSSFComment(sheetComments, ctComment, vmlShape);
            comment.Column = (1);
            Assert.AreEqual(1, comment.Column);
            Assert.AreEqual(1, new CellReference(ctComment.@ref).Col);
            Assert.AreEqual(1, vmlShape.GetClientDataArray(0).GetColumnArray(0));

            comment.Column = (5);
            Assert.AreEqual(5, comment.Column);
            Assert.AreEqual(5, new CellReference(ctComment.@ref).Col);
            Assert.AreEqual(5, vmlShape.GetClientDataArray(0).GetColumnArray(0));
        }
        [Test]
        public void GetSetRow()
        {
            CommentsTable sheetComments = new CommentsTable();
            XSSFVMLDrawing vml = new XSSFVMLDrawing();
            CT_Comment ctComment = sheetComments.NewComment(CellAddress.A1);
            CT_Shape vmlShape = vml.newCommentShape();

            XSSFComment comment = new XSSFComment(sheetComments, ctComment, vmlShape);
            comment.Row = (1);
            Assert.AreEqual(1, comment.Row);
            Assert.AreEqual(1, new CellReference(ctComment.@ref).Row);
            Assert.AreEqual(1, vmlShape.GetClientDataArray(0).GetRowArray(0));

            comment.Row = (5);
            Assert.AreEqual(5, comment.Row);
            Assert.AreEqual(5, new CellReference(ctComment.@ref).Row);
            Assert.AreEqual(5, vmlShape.GetClientDataArray(0).GetRowArray(0));
        }
        [Test]
        public void SetString()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sh = (XSSFSheet)wb.CreateSheet();
            XSSFComment comment = (XSSFComment)sh.CreateDrawingPatriarch().CreateCellComment(new XSSFClientAnchor());

            //passing HSSFRichTextString is incorrect
            try
            {
                comment.String = (new HSSFRichTextString(TEST_RICHTEXTSTRING));
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Only XSSFRichTextString argument is supported", e.Message);
            }

            //simple string argument
            comment.SetString(TEST_RICHTEXTSTRING);
            Assert.AreEqual(TEST_RICHTEXTSTRING, comment.String.String);

            //if the text is already Set, it should be overridden, not Added twice!
            comment.SetString(TEST_RICHTEXTSTRING);

            CT_Comment ctComment = comment.GetCTComment();
            //  Assert.Fail("TODO test case incomplete!?");
            //XmlObject[] obj = ctComment.selectPath(
            //        "declare namespace w='"+XSSFRelation.NS_SPREADSHEETML+"' .//w:text");
            //Assert.AreEqual(1, obj.Length);
            Assert.AreEqual(TEST_RICHTEXTSTRING, comment.String.String);

            //sequential call of comment.String should return the same XSSFRichTextString object
            Assert.AreSame(comment.String, comment.String);

            XSSFRichTextString richText = new XSSFRichTextString(TEST_RICHTEXTSTRING);
            XSSFFont font1 = (XSSFFont)wb.CreateFont();
            font1.FontName = ("Tahoma");
            font1.FontHeightInPoints = 8.5;
            font1.IsItalic = true;
            font1.Color = IndexedColors.BlueGrey.Index;
            richText.ApplyFont(0, 5, font1);

            //check the low-level stuff
            comment.String = richText;
            //obj = ctComment.selectPath(
            //        "declare namespace w='"+XSSFRelation.NS_SPREADSHEETML+"' .//w:text");
            //Assert.AreEqual(1, obj.Length);
            Assert.AreSame(comment.String, richText);
            //check that the rich text is Set in the comment
            CT_RPrElt rPr = richText.GetCTRst().GetRArray(0).rPr;
            Assert.AreEqual(true, rPr.GetIArray(0).val);
            Assert.AreEqual(8.5, rPr.GetSzArray(0).val);
            Assert.AreEqual(IndexedColors.BlueGrey.Index, (short)rPr.GetColorArray(0).indexed);
            Assert.AreEqual("Tahoma", rPr.GetRFontArray(0).val);

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        [Test]
        public void Author()
        {
            CommentsTable sheetComments = new CommentsTable();
            CT_Comment ctComment = sheetComments.NewComment(CellAddress.A1);

            Assert.AreEqual(1, sheetComments.GetNumberOfAuthors());
            XSSFComment comment = new XSSFComment(sheetComments, ctComment, null);
            Assert.AreEqual("", comment.Author);
            comment.Author = ("Apache POI");
            Assert.AreEqual("Apache POI", comment.Author);
            Assert.AreEqual(2, sheetComments.GetNumberOfAuthors());
            comment.Author = ("Apache POI");
            Assert.AreEqual(2, sheetComments.GetNumberOfAuthors());
            comment.Author = ("");
            Assert.AreEqual("", comment.Author);
            Assert.AreEqual(2, sheetComments.GetNumberOfAuthors());
        }

        [Test]
        public void TestBug58175()
        {
            IWorkbook wb = new SXSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(1);
                ICell cell = row.CreateCell(3);
                cell.SetCellValue("F4");
                ICreationHelper factory = wb.GetCreationHelper();
                // When the comment box is visible, have it show in a 1x3 space
                IClientAnchor anchor = factory.CreateClientAnchor();
                anchor.Col1 = (cell.ColumnIndex);
                anchor.Col2 = (cell.ColumnIndex + 1);
                anchor.Row1 = (row.RowNum);
                anchor.Row2 = (row.RowNum + 3);
                XSSFClientAnchor ca = (XSSFClientAnchor)anchor;
                // create comments and vmlDrawing parts if they don't exist
                CommentsTable comments = (((SXSSFWorkbook)wb).XssfWorkbook
                        .GetSheetAt(0) as XSSFSheet).GetCommentsTable(true);
                XSSFVMLDrawing vml = (((SXSSFWorkbook)wb).XssfWorkbook
                        .GetSheetAt(0) as XSSFSheet).GetVMLDrawing(true);
                CT_Shape vmlShape1 = vml.newCommentShape();
                if (ca.IsSet())
                {
                    String position = ca.Col1 + ", 0, " + ca.Row1
                            + ", 0, " + ca.Col2 + ", 0, " + ca.Row2
                            + ", 0";
                    vmlShape1.GetClientDataArray(0).SetAnchorArray(0, position);
                }
                // create the comment in two different ways and verify that there is no difference

                XSSFComment shape1 = new XSSFComment(comments, comments.NewComment(CellAddress.A1), vmlShape1);
                shape1.Column = (ca.Col1);
                shape1.Row = (ca.Row1);
                CT_Shape vmlShape2 = vml.newCommentShape();
                if (ca.IsSet())
                {
                    String position = ca.Col1 + ", 0, " + ca.Row1
                            + ", 0, " + ca.Col2 + ", 0, " + ca.Row2
                            + ", 0";
                    vmlShape2.GetClientDataArray(0).SetAnchorArray(0, position);
                }

                CellAddress ref1 = new CellAddress(ca.Row1, ca.Col1);
                XSSFComment shape2 = new XSSFComment(comments, comments.NewComment(ref1), vmlShape2);

                Assert.AreEqual(shape1.Author, shape2.Author);
                Assert.AreEqual(shape1.ClientAnchor, shape2.ClientAnchor);
                Assert.AreEqual(shape1.Column, shape2.Column);
                Assert.AreEqual(shape1.Row, shape2.Row);
                Assert.AreEqual(shape1.GetCTComment().ToString(), shape2.GetCTComment().ToString());
                Assert.AreEqual(shape1.GetCTComment().@ref, shape2.GetCTComment().@ref);

                /*CommentsTable table1 = shape1.CommentsTable;
                CommentsTable table2 = shape2.CommentsTable;
                Assert.AreEqual(table1.CTComments.toString(), table2.CTComments.toString());
                Assert.AreEqual(table1.NumberOfComments, table2.NumberOfComments);
                Assert.AreEqual(table1.Relations, table2.Relations);*/

                Assert.AreEqual(vmlShape1.ToString().Replace("_x0000_s\\d+", "_x0000_s0000"), 
                    vmlShape2.ToString().Replace("_x0000_s\\d+", "_x0000_s0000"),
                    "The vmlShapes should have equal content afterwards");
            }
            finally
            {
                wb.Close();
            }
        }
        [Ignore("Used for manual testing with opening the resulting Workbook in Excel")]
        [Test]
        public void TestBug58175a()
        {
            IWorkbook wb = new SXSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(1);
                ICell cell = row.CreateCell(3);
                cell.SetCellValue("F4");
                IDrawing drawing = sheet.CreateDrawingPatriarch();
                ICreationHelper factory = wb.GetCreationHelper();
                // When the comment box is visible, have it show in a 1x3 space
                IClientAnchor anchor = factory.CreateClientAnchor();
                anchor.Col1 = (cell.ColumnIndex);
                anchor.Col2 = (cell.ColumnIndex + 1);
                anchor.Row1 = (row.RowNum);
                anchor.Row2 = (row.RowNum + 3);
                // Create the comment and set the text+author
                IComment comment = drawing.CreateCellComment(anchor);
                IRichTextString str = factory.CreateRichTextString("Hello, World!");
                comment.String = (str);
                comment.Author = ("Apache POI");
                /* fixed the problem as well 
                 * comment.setColumn(cell.ColumnIndex);
                 * comment.setRow(cell.RowIndex);
                 */
                // Assign the comment to the cell
                cell.CellComment = (comment);
                FileStream out1 = new FileStream("C:\\temp\\58175.xlsx", FileMode.CreateNew, FileAccess.ReadWrite);
                try
                {
                    wb.Write(out1, false);
                }
                finally
                {
                    out1.Close();
                }
            }
            finally
            {
                wb.Close();
            }
        }
        [Test]
        public void Bug57838DeleteRowsWthCommentsBug()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57838.xlsx");
            ISheet sheet = wb.GetSheetAt(0);
            IComment comment1 = sheet.GetCellComment(new CellAddress(2, 1));
            Assert.IsNotNull(comment1);
            IComment comment2 = sheet.GetCellComment(new CellAddress(2, 2));
            Assert.IsNotNull(comment2);
            IRow row = sheet.GetRow(2);
            Assert.IsNotNull(row);
            sheet.RemoveRow(row); // Remove row from index 2
            row = sheet.GetRow(2);
            Assert.IsNull(row); // Row is null since we deleted it.
            comment1 = sheet.GetCellComment(new CellAddress(2, 1));
            Assert.IsNull(comment1); // comment should be null but will Assert.Fail due to bug
            comment2 = sheet.GetCellComment(new CellAddress(2, 2));
            Assert.IsNull(comment2); // comment should be null but will Assert.Fail due to bug
            wb.Close();
        }

        [Test]
        public void TestRemoveXSSFCellComment()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(1);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue("test");

                IDrawing drawing = sheet.CreateDrawingPatriarch();
                ICreationHelper factory = wb.GetCreationHelper();
                // When the comment box is visible, have it show in a 1x3 space
                IClientAnchor anchor = factory.CreateClientAnchor();
                anchor.Col1 = cell.ColumnIndex;
                anchor.Col2 = cell.ColumnIndex + 1;
                anchor.Row1 = row.RowNum;
                anchor.Row2 = row.RowNum + 3;
                // Create the comment and set the text+author
                IComment comment = drawing.CreateCellComment(anchor);
                IRichTextString str = factory.CreateRichTextString("Hello, World!");
                comment.String = str;
                comment.Author = "Apache POI";

                cell.CellComment = comment;

                var exCellComment = sheet.GetCellComment(new CellAddress(1, 0));
                Assert.IsNotNull(exCellComment);
                Assert.IsTrue(exCellComment.String.String.Equals("Hello, World!"));
                Assert.IsTrue(exCellComment.Author.Equals("Apache POI"));

                cell.RemoveCellComment();
                exCellComment = sheet.GetCellComment(new CellAddress(1, 0));
                Assert.IsNull(exCellComment);

                IComment newComment = drawing.CreateCellComment(anchor);
                newComment.String = str;
                newComment.Author = "Apache POI";
                cell.CellComment = newComment;

                exCellComment = sheet.GetCellComment(new CellAddress(1, 0));
                Assert.NotNull(exCellComment);
                Assert.IsTrue(exCellComment.String.String.Equals("Hello, World!"));
                Assert.IsTrue(exCellComment.Author.Equals("Apache POI"));
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestRemoveSXSSFCellComment()
        {
            IWorkbook wb = new SXSSFWorkbook();
            try
            {
                ISheet sheet = wb.CreateSheet();
                IRow row = sheet.CreateRow(1);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue("test");

                IDrawing drawing = sheet.CreateDrawingPatriarch();
                ICreationHelper factory = wb.GetCreationHelper();
                // When the comment box is visible, have it show in a 1x3 space
                IClientAnchor anchor = factory.CreateClientAnchor();
                anchor.Col1 = cell.ColumnIndex;
                anchor.Col2 = cell.ColumnIndex + 1;
                anchor.Row1 = row.RowNum;
                anchor.Row2 = row.RowNum + 3;
                // Create the comment and set the text+author
                IComment comment = drawing.CreateCellComment(anchor);
                IRichTextString str = factory.CreateRichTextString("Hello, World!");
                comment.String = str;
                comment.Author = "Apache POI";

                cell.CellComment = comment;

                var exCellComment = sheet.GetCellComment(new CellAddress(1, 0));
                Assert.IsNotNull(exCellComment);
                Assert.IsTrue(exCellComment.String.String.Equals("Hello, World!"));
                Assert.IsTrue(exCellComment.Author.Equals("Apache POI"));

                cell.RemoveCellComment();
                exCellComment = sheet.GetCellComment(new CellAddress(1, 0));
                Assert.IsNull(exCellComment);

                IComment newComment = drawing.CreateCellComment(anchor);
                newComment.String = str;
                newComment.Author = "Apache POI";
                cell.CellComment = newComment;

                exCellComment = sheet.GetCellComment(new CellAddress(1, 0));
                Assert.NotNull(exCellComment);
                Assert.IsTrue(exCellComment.String.String.Equals("Hello, World!"));
                Assert.IsTrue(exCellComment.Author.Equals("Apache POI"));
            }
            finally
            {
                wb.Close();
            }
        }
    }

}
