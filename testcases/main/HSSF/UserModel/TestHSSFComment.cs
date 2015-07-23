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

    }
}
