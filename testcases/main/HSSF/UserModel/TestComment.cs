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

namespace TestCases.HSSF.UserModel
{

    /**
     * @author Evgeniy Berlog
     * @date 26.06.12
     */
    public class TestComment
    {

        public void TestResultEqualsToAbstractShape()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor());
            HSSFRow row = sh.CreateRow(0);
            HSSFCell cell = row.CreateCell(0);
            cell.SetCellComment(comment);

            CommentShape commentShape = HSSFTestModelHelper.CreateCommentShape(1025, comment);

            Assert.AreEqual(comment.GetEscherContainer().GetChildRecords().Count, 5);
            Assert.AreEqual(commentShape.GetSpContainer().GetChildRecords().Count, 5);

            //sp record
            byte[] expected = commentShape.GetSpContainer().GetChild(0).Serialize();
            byte[] actual = comment.GetEscherContainer().GetChild(0).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = commentShape.GetSpContainer().GetChild(2).Serialize();
            actual = comment.GetEscherContainer().GetChild(2).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = commentShape.GetSpContainer().GetChild(3).Serialize();
            actual = comment.GetEscherContainer().GetChild(3).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = commentShape.GetSpContainer().GetChild(4).Serialize();
            actual = comment.GetEscherContainer().GetChild(4).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            ObjRecord obj = comment.GetObjRecord();
            ObjRecord objShape = commentShape.GetObjRecord();
            /**shapeId = 1025 % 1024**/
            ((CommonObjectDataSubRecord)objShape.GetSubRecords().Get(0)).SetObjectId(1);

            expected = obj.Serialize();
            actual = objShape.Serialize();

            TextObjectRecord tor = comment.GetTextObjectRecord();
            TextObjectRecord torShape = commentShape.GetTextObjectRecord();

            expected = tor.Serialize();
            actual = torShape.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            NoteRecord note = comment.GetNoteRecord();
            NoteRecord noteShape = commentShape.GetNoteRecord();
            noteShape.SetShapeId(1);

            expected = note.Serialize();
            actual = noteShape.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));
        }

        public void TestAddToExistingFile()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();
            int idx = wb.AddPicture(new byte[] { 1, 2, 3 }, HSSFWorkbook.PICTURE_TYPE_PNG);

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor());
            comment.SetColumn(5);
            comment.SetString(new HSSFRichTextString("comment1"));
            comment = patriarch.CreateCellComment(new HSSFClientAnchor(0, 0, 100, 100, (short)0, 0, (short)10, 10));
            comment.SetRow(5);
            comment.SetString(new HSSFRichTextString("comment2"));
            comment.SetBackgroundImage(idx);
            Assert.AreEqual(comment.GetBackgroundImageId(), idx);

            Assert.AreEqual(patriarch.Children().Count, 2);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            comment = (HSSFComment)patriarch.Children().Get(1);
            Assert.AreEqual(comment.GetBackgroundImageId(), idx);
            comment.ResetBackgroundImage();
            Assert.AreEqual(comment.GetBackgroundImageId(), 0);

            Assert.AreEqual(patriarch.Children().Count, 2);
            comment = patriarch.CreateCellComment(new HSSFClientAnchor());
            comment.SetString(new HSSFRichTextString("comment3"));

            Assert.AreEqual(patriarch.Children().Count, 3);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();
            comment = (HSSFComment)patriarch.Children().Get(1);
            Assert.AreEqual(comment.GetBackgroundImageId(), 0);
            Assert.AreEqual(patriarch.Children().Count, 3);
            Assert.AreEqual(((HSSFComment)patriarch.Children().Get(0)).GetString().GetString(), "comment1");
            Assert.AreEqual(((HSSFComment)patriarch.Children().Get(1)).GetString().GetString(), "comment2");
            Assert.AreEqual(((HSSFComment)patriarch.Children().Get(2)).GetString().GetString(), "comment3");
        }

        public void TestSetGetProperties()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor());
            comment.SetString(new HSSFRichTextString("comment1"));
            Assert.AreEqual(comment.GetString().GetString(), "comment1");

            comment.SetAuthor("poi");
            Assert.AreEqual(comment.Author, "poi");

            comment.SetColumn(3);
            Assert.AreEqual(comment.Column, 3);

            comment.SetRow(4);
            Assert.AreEqual(comment.Row, 4);

            comment.SetVisible(false);
            Assert.AreEqual(comment.IsVisible(), false);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            comment = (HSSFComment)patriarch.Children().Get(0);

            Assert.AreEqual(comment.GetString().GetString(), "comment1");
            Assert.AreEqual("poi", comment.Author);
            Assert.AreEqual(comment.Column, 3);
            Assert.AreEqual(comment.Row, 4);
            Assert.AreEqual(comment.IsVisible(), false);

            comment.SetString(new HSSFRichTextString("comment12"));
            comment.SetAuthor("poi2");
            comment.SetColumn(32);
            comment.SetRow(42);
            comment.SetVisible(true);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();
            comment = (HSSFComment)patriarch.Children().Get(0);

            Assert.AreEqual(comment.GetString().GetString(), "comment12");
            Assert.AreEqual("poi2", comment.Author);
            Assert.AreEqual(comment.Column, 32);
            Assert.AreEqual(comment.Row, 42);
            Assert.AreEqual(comment.IsVisible(), true);
        }

        public void TestExistingFileWithComment()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("comments");
            HSSFPatriarch Drawing = sheet.DrawingPatriarch();
            Assert.AreEqual(1, Drawing.Children().Count);
            HSSFComment comment = (HSSFComment)Drawing.Children().Get(0);
            Assert.AreEqual(comment.Author, "evgeniy");
            Assert.AreEqual(comment.GetString().GetString(), "evgeniy:\npoi Test");
            Assert.AreEqual(comment.Column, 1);
            Assert.AreEqual(comment.Row, 2);
        }

        public void TestFindComments()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor());
            HSSFRow row = sh.CreateRow(5);
            HSSFCell cell = row.CreateCell(4);
            cell.SetCellComment(comment);

            HSSFTestModelHelper.CreateCommentShape(0, comment);

            Assert.IsNotNull(sh.FindCellComment(5, 4));
            Assert.IsNull(sh.FindCellComment(5, 5));

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);

            Assert.IsNotNull(sh.FindCellComment(5, 4));
            Assert.IsNull(sh.FindCellComment(5, 5));
        }

        public void TestInitState()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);
            Assert.AreEqual(agg.GetTailRecords().Count, 0);

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor());
            Assert.AreEqual(agg.GetTailRecords().Count, 1);

            HSSFSimpleShape shape = patriarch.CreateSimpleShape(new HSSFClientAnchor());

            Assert.AreEqual(comment.GetOptRecord().GetEscherProperties().Count, 10);
        }

        public void TestShapeId()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor());

            comment.SetShapeId(2024);
            /**
             * SpRecord.id == shapeId
             * ObjRecord.id == shapeId % 1024
             * NoteRecord.id == ObjectRecord.id == shapeId % 1024
             */

            Assert.AreEqual(comment.GetShapeId(), 2024);

            CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)comment.GetObjRecord().GetSubRecords().Get(0);
            Assert.AreEqual(cod.GetObjectId(), 1000);
            EscherSpRecord spRecord = (EscherSpRecord)comment.GetEscherContainer().GetChild(0);
            Assert.AreEqual(spRecord.GetShapeId(), 2024);
            Assert.AreEqual(comment.GetShapeId(), 2024);
            Assert.AreEqual(comment.GetNoteRecord().GetShapeId(), 1000);
        }

        public void TestAttemptToSave2CommentsWithSameCoordinates()
        {
            Object err = null;

            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();
            patriarch.CreateCellComment(new HSSFClientAnchor());
            patriarch.CreateCellComment(new HSSFClientAnchor());

            try
            {
                HSSFTestDataSamples.WriteOutAndReadBack(wb);
            }
            catch (InvalidOperationException e)
            {
                err = 1;
                Assert.AreEqual(e.GetMessage(), "found multiple cell comments for cell $A$1");
            }
            Assert.IsNotNull(err);
        }
    }

}
