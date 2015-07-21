
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
    using System;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Util;
    using NUnit.Framework;
    using NPOI.SS.UserModel;
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    /**
     * Test the ability to clone a sheet. 
     *  If Adding new records that belong to a sheet (as opposed to a book)
     *  Add that record to the sheet in the TestCloneSheetBasic method. 
     * @author  avik
     */
    [TestFixture]
    public class TestCloneSheet
    {
        [Test]
        public void TestCloneSheetBasic()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            ISheet s = b.CreateSheet("Test");
            s.AddMergedRegion(new CellRangeAddress(0, 1, 0, 1));
            ISheet clonedSheet = b.CloneSheet(0);
            Assert.AreEqual(1, clonedSheet.NumMergedRegions, "One merged area");

            b.Close();
        }

        /**
         * Ensures that pagebreak cloning works properly
         *
         */
        [Test]
        public void TestPageBreakClones()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            ISheet s = b.CreateSheet("Test");
            s.SetRowBreak(3);
            s.SetColumnBreak((short)6);

            ISheet clone = b.CloneSheet(0);
            Assert.IsTrue(clone.IsRowBroken(3), "Row 3 not broken");
            Assert.IsTrue(clone.IsColumnBroken((short)6), "Column 6 not broken");

            s.RemoveRowBreak(3);

            Assert.IsTrue(clone.IsRowBroken(3), "Row 3 still should be broken");

            b.Close();
        }
        [Test]
        public void TestCloneSheetWithoutDrawings()
        {

            HSSFWorkbook b = new HSSFWorkbook();
            HSSFSheet s = b.CreateSheet("Test") as HSSFSheet;
            HSSFSheet s2 = s.CloneSheet(b) as HSSFSheet;

            Assert.IsNull(s.DrawingPatriarch);
            Assert.IsNull(s2.DrawingPatriarch);
            Assert.AreEqual(HSSFTestHelper.GetSheetForTest(s).Records.Count, HSSFTestHelper.GetSheetForTest(s2).Records.Count);

        }
        [Test]
        public void TestCloneSheetWithEmptyDrawingAggregate()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            HSSFSheet s = b.CreateSheet("Test") as HSSFSheet;
            HSSFPatriarch patriarch = s.CreateDrawingPatriarch() as HSSFPatriarch;

            EscherAggregate agg1 = patriarch.GetBoundAggregate();

            HSSFSheet s2 = s.CloneSheet(b) as HSSFSheet;

            patriarch = s2.DrawingPatriarch as HSSFPatriarch;

            EscherAggregate agg2 = patriarch.GetBoundAggregate();

            EscherSpRecord sp1 = (EscherSpRecord)agg1.GetEscherContainer().GetChild(1).GetChild(0).GetChild(1);
            EscherSpRecord sp2 = (EscherSpRecord)agg2.GetEscherContainer().GetChild(1).GetChild(0).GetChild(1);

            Assert.AreEqual(sp1.ShapeId, 1024);
            Assert.AreEqual(sp2.ShapeId, 2048);

            EscherDgRecord dg = (EscherDgRecord)agg2.GetEscherContainer().GetChild(0);

            Assert.AreEqual(dg.LastMSOSPID, 2048);
            Assert.AreEqual(dg.Instance, 0x2);

            //everything except id and DgRecord.lastMSOSPID and DgRecord.Instance must be the same

            sp2.ShapeId = (1024);
            dg.LastMSOSPID = (1024);
            dg.Instance =((short)0x1);

            Assert.AreEqual(agg1.Serialize().Length, agg2.Serialize().Length);
            Assert.AreEqual(agg1.ToXml(""), agg2.ToXml(""));
            Assert.IsTrue(Arrays.Equals(agg1.Serialize(), agg2.Serialize()));
        }
        [Test]
        public void TestCloneComment()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch p = sh.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFComment c = p.CreateComment(new HSSFClientAnchor(0, 0, 100, 100, (short)0, 0, (short)5, 5));
            c.Column=(1);
            c.Row=(2);
            c.String=(new HSSFRichTextString("qwertyuio"));

            HSSFSheet sh2 = wb.CloneSheet(0) as HSSFSheet;
            HSSFPatriarch p2 = sh2.DrawingPatriarch as HSSFPatriarch;
            HSSFComment c2 = (HSSFComment)p2.Children[0];

            Assert.AreEqual(c.String, c2.String);
            Assert.AreEqual(c.Row, c2.Row);
            Assert.AreEqual(c.Column, c2.Column);

            // The ShapeId is not equal? 
            // assertEquals(c.getNoteRecord().getShapeId(), c2.getNoteRecord().getShapeId());

            Assert.IsTrue(Arrays.Equals(c2.GetTextObjectRecord().Serialize(), c.GetTextObjectRecord().Serialize()));

            // ShapeId is different
            CommonObjectDataSubRecord subRecord = (CommonObjectDataSubRecord)c2.GetObjRecord().SubRecords[0];
            subRecord.ObjectId = (1025);

            Assert.IsTrue(Arrays.Equals(c2.GetObjRecord().Serialize(), c.GetObjRecord().Serialize()));

            // ShapeId is different
            c2.NoteRecord.ShapeId = (1025);

            Assert.IsTrue(Arrays.Equals(c2.NoteRecord.Serialize(), c.NoteRecord.Serialize()));


            //everything except spRecord.shapeId must be the same
            Assert.IsFalse(Arrays.Equals(c2.GetEscherContainer().Serialize(), c.GetEscherContainer().Serialize()));
            EscherSpRecord sp = (EscherSpRecord)c2.GetEscherContainer().GetChild(0);
            sp.ShapeId=(1025);
            Assert.IsTrue(Arrays.Equals(c2.GetEscherContainer().Serialize(), c.GetEscherContainer().Serialize()));

            wb.Close();
        }
    }
}