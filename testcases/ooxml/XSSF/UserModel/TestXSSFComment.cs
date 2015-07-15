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
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Model;
using NUnit.Framework;
using TestCases.SS.UserModel;
namespace NPOI.XSSF.UserModel
{
    /**
     * @author Yegor Kozlov
     */
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

            CT_Comment ctComment = sheetComments.NewComment("A1");
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
            CT_Comment ctComment = sheetComments.NewComment("A1");
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
            CT_Comment ctComment = sheetComments.NewComment("A1");
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
            //        "declare namespace w='http://schemas.Openxmlformats.org/spreadsheetml/2006/main' .//w:text");
            //Assert.AreEqual(1, obj.Length);
            Assert.AreEqual(TEST_RICHTEXTSTRING, comment.String.String);

            //sequential call of comment.String should return the same XSSFRichTextString object
            Assert.AreSame(comment.String, comment.String);

            XSSFRichTextString richText = new XSSFRichTextString(TEST_RICHTEXTSTRING);
            XSSFFont font1 = (XSSFFont)wb.CreateFont();
            font1.FontName = ("Tahoma");
            font1.FontHeight = 8.5;
            font1.IsItalic = true;
            font1.Color = IndexedColors.BlueGrey.Index;
            richText.ApplyFont(0, 5, font1);

            //check the low-level stuff
            comment.String = richText;
            //obj = ctComment.selectPath(
            //        "declare namespace w='http://schemas.Openxmlformats.org/spreadsheetml/2006/main' .//w:text");
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
            CT_Comment ctComment = sheetComments.NewComment("A1");

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
    }



}