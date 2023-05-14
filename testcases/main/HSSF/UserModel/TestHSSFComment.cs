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
namespace TestCases.HSSF.UserModel
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using TestCases.HSSF;
    using TestCases.SS.UserModel;
    using System;
    using TestCases.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.DDF;
    using NPOI.Util;
    using NPOI.HSSF.Model;
    using static TestCases.POIFS.Storage.RawDataUtil;

    /**
     * Tests TestHSSFCellComment.
     *
     * @author  Yegor Kozlov
     */
    [TestFixture]
    public class TestHSSFComment:BaseTestCellComment
    {
        public TestHSSFComment(): base(HSSFITestDataProvider.Instance)
        {

        }

        [Test]
        public void DefaultShapeType()
        {
            HSSFComment comment = new HSSFComment((HSSFShape)null, new HSSFClientAnchor());
            Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_COMMENT, comment.ShapeType);
        }
        /**
         *  HSSFCell#findCellComment should NOT rely on the order of records
         * when matching cells and their cell comments. The correct algorithm is to map
         */
        [Test]
        public void Bug47924()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("47924.xls");
            ISheet sheet = wb.GetSheetAt(0);
            ICell cell;
            IComment comment;

            cell = sheet.GetRow(0).GetCell(0);
            comment = cell.CellComment;
            Assert.AreEqual("a1", comment.String.String);

            cell = sheet.GetRow(1).GetCell(0);
            comment = cell.CellComment;
            Assert.AreEqual("a2", comment.String.String);

            cell = sheet.GetRow(2).GetCell(0);
            comment = cell.CellComment;
            Assert.AreEqual("a3", comment.String.String);

            cell = sheet.GetRow(2).GetCell(2);
            comment = cell.CellComment;
            Assert.AreEqual("c3", comment.String.String);

            cell = sheet.GetRow(4).GetCell(1);
            comment = cell.CellComment;
            Assert.AreEqual("b5", comment.String.String);

            cell = sheet.GetRow(5).GetCell(2);
            comment = cell.CellComment;
            Assert.AreEqual("c6", comment.String.String);

            wb.Close();
        }

        [Test]
        public void TestBug56380InsertComments()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            int noOfRows = 1025;
            String comment = "c";

            for (int i = 0; i < noOfRows; i++)
            {
                IRow row = sheet.CreateRow(i);
                ICell cell = row.CreateCell(0);
                insertComment(drawing, cell, comment + i);
            }

            // assert that the comments are Created properly before writing
            CheckComments(sheet, noOfRows, comment);

            /*// store in temp-file
            OutputStream fs = new FileOutputStream("/tmp/56380.xls");
            try {
                sheet.Workbook.Write(fs);
            } finally {
                fs.Close();
            }*/

            // save and recreate the workbook from the saved file
            HSSFWorkbook workbookBack = HSSFTestDataSamples.WriteOutAndReadBack(workbook);
            sheet = workbookBack.GetSheetAt(0);

            // assert that the comments are Created properly After Reading back in
            CheckComments(sheet, noOfRows, comment);

            workbook.Close();
            workbookBack.Close();
        }

        [Test]
        public void TestBug56380InsertTooManyComments()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            try
            {
                ISheet sheet = workbook.CreateSheet();
                IDrawing drawing = sheet.CreateDrawingPatriarch();
                String comment = "c";

                for (int rowNum = 0; rowNum < 258; rowNum++)
                {
                    sheet.CreateRow(rowNum);
                }

                // should still work, for some reason DrawingManager2.AllocateShapeId() skips the first 1024...
                for (int count = 1025; count < 65535; count++)
                {
                    int rowNum = count / 255;
                    int cellNum = count % 255;
                    ICell cell = sheet.GetRow(rowNum).CreateCell(cellNum);
                    try
                    {
                        IComment commentObj = insertComment(drawing, cell, comment + cellNum);

                        Assert.AreEqual(count, ((HSSFComment)commentObj).NoteRecord.ShapeId);
                    }
                    catch (ArgumentException e)
                    {
                        throw new ArgumentException("While Adding shape number " + count, e);
                    }
                }

                // this should now fail to insert
                IRow row = sheet.CreateRow(257);
                ICell cell2 = row.CreateCell(0);
                insertComment(drawing, cell2, comment + 0);
            }
            finally
            {
                workbook.Close();
            }
        }

        private void CheckComments(ISheet sheet, int noOfRows, String comment)
        {
            for (int i = 0; i < noOfRows; i++)
            {
                Assert.IsNotNull(sheet.GetRow(i));
                Assert.IsNotNull(sheet.GetRow(i).GetCell(0));
                Assert.IsNotNull(sheet.GetRow(i).GetCell(0).CellComment, "Did not Get a Cell Comment for row " + i);
                Assert.IsNotNull(sheet.GetRow(i).GetCell(0).CellComment.String);
                Assert.AreEqual(comment + i, sheet.GetRow(i).GetCell(0).CellComment.String.String);
            }
        }

        private IComment insertComment(IDrawing Drawing, ICell cell, String message)
        {
            ICreationHelper factory = cell.Sheet.Workbook.GetCreationHelper();

            IClientAnchor anchor = factory.CreateClientAnchor();
            anchor.Col1 = (/*setter*/cell.ColumnIndex);
            anchor.Col2 = (/*setter*/cell.ColumnIndex + 1);
            anchor.Row1 = (/*setter*/cell.RowIndex);
            anchor.Row2 = (/*setter*/cell.RowIndex + 1);
            anchor.Dx1 = (/*setter*/100);
            anchor.Dx2 = (/*setter*/100);
            anchor.Dy1 = (/*setter*/100);
            anchor.Dy2 = (/*setter*/100);

            IComment comment = Drawing.CreateCellComment(anchor);

            IRichTextString str = factory.CreateRichTextString(message);
            comment.String = (/*setter*/str);
            comment.Author = (/*setter*/"fanfy");
            cell.CellComment = (/*setter*/comment);

            return comment;
        }
        [Test]
        public void ResultEqualsToNonExistingAbstractShape()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;
            HSSFRow row = sh.CreateRow(0) as HSSFRow;
            HSSFCell cell = row.CreateCell(0) as HSSFCell;
            cell.CellComment = (comment);

            Assert.AreEqual(comment.GetEscherContainer().ChildRecords.Count, 5);

            //sp record
            byte[] expected = Decompress("H4sIAAAAAAAAAFvEw/WBg4GBgZEFSHAxMAAA9gX7nhAAAAA=");
            byte[] actual = comment.GetEscherContainer().GetChild(0).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = Decompress("H4sIAAAAAAAAAGNgEPggxIANAABK4+laGgAAAA==");
            actual = comment.GetEscherContainer().GetChild(2).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = Decompress("H4sIAAAAAAAAAGNgEPzAAAQACl6c5QgAAAA=");
            actual = comment.GetEscherContainer().GetChild(3).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = Decompress("H4sIAAAAAAAAAGNg4P3AAAQA6pyIkQgAAAA=");
            actual = comment.GetEscherContainer().GetChild(4).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            ObjRecord obj = comment.GetObjRecord();

            expected = Decompress("H4sIAAAAAAAAAItlMGEQZRBikGRgZBF0YEACvAxiDLgBAJZsuoU4AAAA");
            actual = obj.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            //assertArrayEquals(expected, actual);

            TextObjectRecord tor = comment.GetTextObjectRecord();

            expected = Decompress("H4sIAAAAAAAAANvGKMQgxMSABgBGi8T+FgAAAA==");
            actual = tor.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            NoteRecord note = comment.NoteRecord;

            expected = Decompress("H4sIAAAAAAAAAJNh4GGAAEYWEAkAS0KXuRAAAAA=");
            actual = note.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            wb.Close();
        }
        [Test]
        public void AddToExistingFile()
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
        public void SetGetProperties()
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
        public void ExistingFileWithComment()
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

            wb.Close();
        }
        [Test]
        public void FindComments()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFComment comment = patriarch.CreateCellComment(new HSSFClientAnchor()) as HSSFComment;
            HSSFRow row = sh.CreateRow(5) as HSSFRow;
            HSSFCell cell = row.CreateCell(4) as HSSFCell;
            cell.CellComment = (comment);

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
        public void InitState()
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
        public void ShapeId()
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
    }
}
