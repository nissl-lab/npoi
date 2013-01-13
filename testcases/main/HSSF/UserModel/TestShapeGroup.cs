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
using NUnit.Framework;
using NPOI.HSSF.UserModel;
using NPOI.DDF;
using System.Reflection;
using System.Diagnostics;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;

namespace TestCases.HSSF.UserModel
{

    /**
     * @author Evgeniy Berlog
     * @date 29.06.12
     */
    [TestFixture]
    public class TestShapeGroup
    {
        [Test]
        public void TestSetGetCoordinates()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
            Assert.AreEqual(group.X1, 0);
            Assert.AreEqual(group.Y1, 0);
            Assert.AreEqual(group.X2, 1023);
            Assert.AreEqual(group.Y2, 255);

            group.SetCoordinates(1, 2, 3, 4);

            Assert.AreEqual(group.X1, 1);
            Assert.AreEqual(group.Y1, 2);
            Assert.AreEqual(group.X2, 3);
            Assert.AreEqual(group.Y2, 4);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            group = (HSSFShapeGroup)patriarch.Children[(0)];
            Assert.AreEqual(group.X1, 1);
            Assert.AreEqual(group.Y1, 2);
            Assert.AreEqual(group.X2, 3);
            Assert.AreEqual(group.Y2, 4);
        }
        [Test]
        public void TestAddToExistingFile()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFShapeGroup group1 = patriarch.CreateGroup(new HSSFClientAnchor());
            HSSFShapeGroup group2 = patriarch.CreateGroup(new HSSFClientAnchor());

            group1.SetCoordinates(1, 2, 3, 4);
            group2.SetCoordinates(5, 6, 7, 8);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(patriarch.Children.Count, 2);

            HSSFShapeGroup group3 = patriarch.CreateGroup(new HSSFClientAnchor());
            group3.SetCoordinates(9, 10, 11, 12);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sh.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(patriarch.Children.Count, 3);
        }
        [Test]
        public void TestModify()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // create a sheet with a text box
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFShapeGroup group1 = patriarch.CreateGroup(new
                    HSSFClientAnchor(0, 0, 0, 0,
                    (short)0, 0, (short)15, 25));
            group1.SetCoordinates(0, 0, 792, 612);

            HSSFTextbox textbox1 = group1.CreateTextbox(new
                    HSSFChildAnchor(100, 100, 300, 300));
            HSSFRichTextString rt1 = new HSSFRichTextString("Hello, World!");
            textbox1.String = rt1;

            // Write, read back and check that our text box is there
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, patriarch.Children.Count);

            group1 = (HSSFShapeGroup)patriarch.Children[(0)];
            Assert.AreEqual(1, group1.Children.Count);
            textbox1 = (HSSFTextbox)group1.Children[(0)];
            Assert.AreEqual("Hello, World!", textbox1.String.String);

            // modify anchor
            Assert.AreEqual(new HSSFChildAnchor(100, 100, 300, 300),
                    textbox1.Anchor);
            HSSFChildAnchor newAnchor = new HSSFChildAnchor(200, 200, 400, 400);
            textbox1.Anchor = newAnchor;
            // modify text
            textbox1.String = new HSSFRichTextString("Hello, World! (modified)");

            // add a new text box
            HSSFTextbox textbox2 = group1.CreateTextbox(new
                    HSSFChildAnchor(400, 400, 600, 600));
            HSSFRichTextString rt2 = new HSSFRichTextString("Hello, World-2");
            textbox2.String = rt2;
            Assert.AreEqual(2, group1.Children.Count);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, patriarch.Children.Count);

            group1 = (HSSFShapeGroup)patriarch.Children[(0)];
            Assert.AreEqual(2, group1.Children.Count);
            textbox1 = (HSSFTextbox)group1.Children[(0)];
            Assert.AreEqual("Hello, World! (modified)",
                    textbox1.String.String);
            Assert.AreEqual(new HSSFChildAnchor(200, 200, 400, 400),
                    textbox1.Anchor);

            textbox2 = (HSSFTextbox)group1.Children[(1)];
            Assert.AreEqual("Hello, World-2", textbox2.String.String);
            Assert.AreEqual(new HSSFChildAnchor(400, 400, 600, 600),
                    textbox2.Anchor);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            group1 = (HSSFShapeGroup)patriarch.Children[(0)];
            textbox1 = (HSSFTextbox)group1.Children[(0)];
            textbox2 = (HSSFTextbox)group1.Children[(1)];
            HSSFTextbox textbox3 = group1.CreateTextbox(new
                    HSSFChildAnchor(400, 200, 600, 400));
            HSSFRichTextString rt3 = new HSSFRichTextString("Hello, World-3");
            textbox3.String = rt3;
        }
        [Test]
        public void TestAddShapesToGroup()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // create a sheet with a text box
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
            int index = wb.AddPicture(new byte[] { 1, 2, 3 }, PictureType.JPEG);
            group.CreatePicture(new HSSFChildAnchor(), index);
            HSSFPolygon polygon = group.CreatePolygon(new HSSFChildAnchor());
            polygon.SetPoints(new int[] { 1, 100, 1 }, new int[] { 1, 50, 100 });
            group.CreateTextbox(new HSSFChildAnchor());
            group.CreateShape(new HSSFChildAnchor());

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, patriarch.Children.Count);

            Assert.IsTrue(patriarch.Children[0] is HSSFShapeGroup);
            group = (HSSFShapeGroup)patriarch.Children[0];

            Assert.AreEqual(group.Children.Count, 4);

            Assert.IsTrue(group.Children[0] is HSSFPicture);
            Assert.IsTrue(group.Children[1] is HSSFPolygon);
            Assert.IsTrue(group.Children[2] is HSSFTextbox);
            Assert.IsTrue(group.Children[3] is HSSFSimpleShape);

            HSSFShapeGroup group2 = patriarch.CreateGroup(new HSSFClientAnchor());

            index = wb.AddPicture(new byte[] { 2, 2, 2 }, PictureType.JPEG);
            group2.CreatePicture(new HSSFChildAnchor(), index);
            polygon = group2.CreatePolygon(new HSSFChildAnchor());
            polygon.SetPoints(new int[] { 1, 100, 1 }, new int[] { 1, 50, 100 });
            group2.CreateTextbox(new HSSFChildAnchor());
            group2.CreateShape(new HSSFChildAnchor());
            group2.CreateShape(new HSSFChildAnchor());

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(2, patriarch.Children.Count);

            group = (HSSFShapeGroup)patriarch.Children[1];

            Assert.AreEqual(group.Children.Count, 5);

            Assert.IsTrue(group.Children[0] is HSSFPicture);
            Assert.IsTrue(group.Children[1] is HSSFPolygon);
            Assert.IsTrue(group.Children[2] is HSSFTextbox);
            Assert.IsTrue(group.Children[3] is HSSFSimpleShape);
            Assert.IsTrue(group.Children[4] is HSSFSimpleShape);

            int shapeid = group.ShapeId;
        }
        [Test]
        public void TestSpgrRecord()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            // create a sheet with a text box
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
            Assert.AreSame(((EscherContainerRecord)group.GetEscherContainer().GetChild(0)).GetChildById(EscherSpgrRecord.RECORD_ID), 
                GetSpgrRecord(group));
        }

        private static EscherSpgrRecord GetSpgrRecord(HSSFShapeGroup group)
        {
            FieldInfo spgrField = null;
            try
            {
                spgrField = group.GetType().GetField("_spgrRecord",BindingFlags.NonPublic| BindingFlags.Instance);
                //spgrField.SetAccessible(true);
                return (EscherSpgrRecord)spgrField.GetValue(group);
            }
            catch (NullReferenceException e)
            {
                Debug.Write(e.Message);
            }
            catch (FieldAccessException e)
            {
                Debug.Write(e.Message);
            }
            return null;
        }
        [Test]
        public void TestClearShapes()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());

            group.CreateShape(new HSSFChildAnchor());
            group.CreateShape(new HSSFChildAnchor());

            EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);

            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 5);
            Assert.AreEqual(agg.TailRecords.Count, 0);
            Assert.AreEqual(group.Children.Count, 2);

            group.Clear();

            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 1);
            Assert.AreEqual(agg.TailRecords.Count, 0);
            Assert.AreEqual(group.Children.Count, 0);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            group = (HSSFShapeGroup)patriarch.Children[(0)];

            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 1);
            Assert.AreEqual(agg.TailRecords.Count, 0);
            Assert.AreEqual(group.Children.Count, 0);
        }
    }


}
