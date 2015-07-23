/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Model;
using NUnit.Framework;
using TestCases.HSSF.Model;
using NPOI.HSSF.Record;
using NPOI.Util;
using NPOI.SS.UserModel;
using NPOI.DDF;
using System.Diagnostics.CodeAnalysis;

namespace TestCases.HSSF.UserModel
{

    /**
     * @author Evgeniy Berlog
     * @date 26.06.12
     */
    [TestFixture]
    public class TestComment
    {
        [Test]
        public void TestResultEqualsToAbstractShape()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;
            HSSFRow row = sh.CreateRow(0) as HSSFRow;
            HSSFCell cell = row.CreateCell(0) as HSSFCell;
            cell.CellComment = (comment);

            CommentShape commentShape = HSSFTestModelHelper.CreateCommentShape(1025, comment);

            Assert.AreEqual(comment.GetEscherContainer().ChildRecords.Count, 5);
            Assert.AreEqual(commentShape.SpContainer.ChildRecords.Count, 5);

            //sp record
            byte[] expected = commentShape.SpContainer.GetChild(0).Serialize();
            byte[] actual = comment.GetEscherContainer().GetChild(0).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = commentShape.SpContainer.GetChild(2).Serialize();
            actual = comment.GetEscherContainer().GetChild(2).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = commentShape.SpContainer.GetChild(3).Serialize();
            actual = comment.GetEscherContainer().GetChild(3).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = commentShape.SpContainer.GetChild(4).Serialize();
            actual = comment.GetEscherContainer().GetChild(4).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            ObjRecord obj = comment.GetObjRecord();
            ObjRecord objShape = commentShape.ObjRecord;

            expected = obj.Serialize();
            actual = objShape.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            //assertArrayEquals(expected, actual);


            TextObjectRecord tor = comment.GetTextObjectRecord();
            TextObjectRecord torShape = commentShape.TextObjectRecord;

            expected = tor.Serialize();
            actual = torShape.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            NoteRecord note = comment.NoteRecord;
            NoteRecord noteShape = commentShape.NoteRecord;

            expected = note.Serialize();
            actual = noteShape.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            wb.Close();
        }
        [Test]
        public void TestAddToExistingFile()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;
            int idx = wb.AddPicture(new byte[] { 1, 2, 3 }, PictureType.PNG);

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;
            comment.Column = (5);
            comment.String = new HSSFRichTextString("comment1");
            comment = patriarch.CreateCellComment(new HSSFClientAnchor(0, 0, 100, 100, (short)0, 0, (short)10, 10)) as HSSFComment;
            comment.Row = (5);
            comment.String = new HSSFRichTextString("comment2");
            comment.SetBackgroundImage(idx);
            Assert.AreEqual(comment.GetBackgroundImageId(), idx);

            Assert.AreEqual(patriarch.Children.Count, 2);

            HSSFWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wbBack.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            comment = (HSSFComment)patriarch.Children[(1)];
            Assert.AreEqual(comment.GetBackgroundImageId(), idx);
            comment.ResetBackgroundImage();
            Assert.AreEqual(comment.GetBackgroundImageId(), 0);

            Assert.AreEqual(patriarch.Children.Count, 2);
            comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;
            comment.String = new HSSFRichTextString("comment3");

            Assert.AreEqual(patriarch.Children.Count, 3);
            HSSFWorkbook wbBack2 = HSSFTestDataSamples.WriteOutAndReadBack(wbBack);
            sh = wbBack2.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;
            comment = (HSSFComment)patriarch.Children[1];
            Assert.AreEqual(comment.GetBackgroundImageId(), 0);
            Assert.AreEqual(patriarch.Children.Count, 3);
            Assert.AreEqual(((HSSFComment)patriarch.Children[0]).String.String, "comment1");
            Assert.AreEqual(((HSSFComment)patriarch.Children[1]).String.String, "comment2");
            Assert.AreEqual(((HSSFComment)patriarch.Children[2]).String.String, "comment3");

            wb.Close();
            wbBack.Close();
            wbBack2.Close();
        }
        [Test]
        public void TestSetGetProperties()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;
            comment.String = new HSSFRichTextString("comment1");
            Assert.AreEqual(comment.String.String, "comment1");

            comment.Author = ("poi");
            Assert.AreEqual(comment.Author, "poi");

            comment.Column = (3);
            Assert.AreEqual(comment.Column, 3);

            comment.Row = (4);
            Assert.AreEqual(comment.Row, 4);

            comment.Visible = (false);
            Assert.AreEqual(comment.Visible, false);

            HSSFWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wbBack.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            comment = (HSSFComment)patriarch.Children[0];

            Assert.AreEqual(comment.String.String, "comment1");
            Assert.AreEqual("poi", comment.Author);
            Assert.AreEqual(comment.Column, 3);
            Assert.AreEqual(comment.Row, 4);
            Assert.AreEqual(comment.Visible, false);

            comment.String = new HSSFRichTextString("comment12");
            comment.Author = ("poi2");
            comment.Column = (32);
            comment.Row = (42);
            comment.Visible = (true);

            HSSFWorkbook wbBack2 = HSSFTestDataSamples.WriteOutAndReadBack(wbBack);
            sh = wbBack2.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;
            comment = (HSSFComment)patriarch.Children[0];

            Assert.AreEqual(comment.String.String, "comment12");
            Assert.AreEqual("poi2", comment.Author);
            Assert.AreEqual(comment.Column, 32);
            Assert.AreEqual(comment.Row, 42);
            Assert.AreEqual(comment.Visible, true);

            wb.Close();
            wbBack.Close();
            wbBack2.Close();
        }
        [Test]
        public void TestExistingFileWithComment()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("comments") as HSSFSheet;
            HSSFPatriarch Drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, Drawing.Children.Count);
            HSSFComment comment = (HSSFComment)Drawing.Children[(0)];
            Assert.AreEqual(comment.Author, "evgeniy");
            Assert.AreEqual(comment.String.String, "evgeniy:\npoi test");
            Assert.AreEqual(comment.Column, 1);
            Assert.AreEqual(comment.Row, 2);
        }
        [Test]
        public void TestFindComments()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;
            HSSFRow row = sh.CreateRow(5) as HSSFRow;
            HSSFCell cell = row.CreateCell(4) as HSSFCell;
            cell.CellComment = (comment);

            HSSFTestModelHelper.CreateCommentShape(0, comment);

            Assert.IsNotNull(sh.FindCellComment(5, 4));
            Assert.IsNull(sh.FindCellComment(5, 5));

            HSSFWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wbBack.GetSheetAt(0) as HSSFSheet;

            Assert.IsNotNull(sh.FindCellComment(5, 4));
            Assert.IsNull(sh.FindCellComment(5, 5));

            wb.Close();
            wbBack.Close();
        }
        [Test]
        public void TestInitState()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);
            Assert.AreEqual(agg.TailRecords.Count, 0);

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;
            Assert.AreEqual(agg.TailRecords.Count, 1);

            HSSFSimpleShape shape = patriarch.CreateSimpleShape(new HSSFClientAnchor());
            Assert.IsNotNull(shape);

            Assert.AreEqual(comment.GetOptRecord().EscherProperties.Count, 10);
            wb.Close();
        }
        [Test]
        public void TestShapeId()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;

            comment.ShapeId = 2024;

            Assert.AreEqual(comment.ShapeId, 2024);

            CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)comment.GetObjRecord().SubRecords[0];
            Assert.AreEqual(2024, cod.ObjectId);
            EscherSpRecord spRecord = (EscherSpRecord)comment.GetEscherContainer().GetChild(0);
            Assert.AreEqual(2024, spRecord.ShapeId);
            Assert.AreEqual(2024, comment.ShapeId);
            Assert.AreEqual(2024, comment.NoteRecord.ShapeId);

            wb.Close();
        }
        [Test]
        public void TestAttemptToSave2CommentsWithSameCoordinates()
        {
            Object err = null;

            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;
            patriarch.CreateCellComment(new HSSFClientAnchor());
            patriarch.CreateCellComment(new HSSFClientAnchor());

            try
            {
                HSSFTestDataSamples.WriteOutAndReadBack(wb);
            }
            catch (InvalidOperationException e)
            {
                err = 1;
                Assert.AreEqual(e.Message, "found multiple cell comments for cell A1");
            }
            Assert.IsNotNull(err);
        }
    }

}
