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
using NPOI.HSSF.Model;
using NPOI.DDF;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using TestCases.HSSF.UserModel;
using NPOI.Util;

namespace TestCases.HSSF.Model
{
    /**
     * @author Evgeniy Berlog
     * date: 12.06.12
     */
    [TestFixture]
    public class TestDrawingShapes
    {

        /**
         * HSSFShape tree bust be built correctly
         * Check file with such records structure:
         * -patriarch
         * --shape
         * --group
         * ---group
         * ----shape
         * ----shape
         * ---shape
         * ---group
         * ----shape
         * ----shape
         */
        [Test]
        public void TestDrawingGroups()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("groups") as HSSFSheet;
            HSSFPatriarch patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(patriarch.Children.Count, 2);
            HSSFShapeGroup group = (HSSFShapeGroup)patriarch.Children[1];
            Assert.AreEqual(3, group.Children.Count);
            HSSFShapeGroup group1 = (HSSFShapeGroup)group.Children[0];
            Assert.AreEqual(2, group1.Children.Count);
            group1 = (HSSFShapeGroup)group.Children[2];
            Assert.AreEqual(2, group1.Children.Count);
        }

        public void TestHSSFShapeCompatibility()
        {
            HSSFSimpleShape shape = new HSSFSimpleShape(null, new HSSFClientAnchor());
            shape.ShapeType=(HSSFSimpleShape.OBJECT_TYPE_LINE);
            Assert.AreEqual(0x08000040, shape.LineStyleColor);
            Assert.AreEqual(0x08000009, shape.FillColor);
            Assert.AreEqual(HSSFShape.LINEWIDTH_DEFAULT, shape.LineWidth);
            Assert.AreEqual(HSSFShape.LINESTYLE_SOLID, shape.LineStyle);
            Assert.IsFalse(shape.IsNoFill);

            AbstractShape sp = AbstractShape.CreateShape(shape, 1);
            EscherContainerRecord spContainer = sp.SpContainer;
            EscherOptRecord opt = spContainer.GetChildById(EscherOptRecord.RECORD_ID) as EscherOptRecord;

            Assert.AreEqual(7, opt.EscherProperties.Count);
            Assert.AreEqual(true,
                    ((EscherBoolProperty)opt.Lookup(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE)).IsTrue);
            Assert.AreEqual(0x00000004,
                    ((EscherSimpleProperty)opt.Lookup(EscherProperties.GEOMETRY__SHAPEPATH)).PropertyValue);
            Assert.AreEqual(0x08000009,
                    ((EscherSimpleProperty)opt.Lookup(EscherProperties.FILL__FILLCOLOR)).PropertyValue);
            Assert.AreEqual(true,
                    ((EscherBoolProperty)opt.Lookup(EscherProperties.FILL__NOFILLHITTEST)).IsTrue);
            Assert.AreEqual(0x08000040,
                    ((EscherSimpleProperty)opt.Lookup(EscherProperties.LINESTYLE__COLOR)).PropertyValue);
            Assert.AreEqual(true,
                    ((EscherBoolProperty)opt.Lookup(EscherProperties.LINESTYLE__NOLINEDRAWDASH)).IsTrue);
            Assert.AreEqual(true,
                    ((EscherBoolProperty)opt.Lookup(EscherProperties.GROUPSHAPE__PRINT)).IsTrue);
        }

        public void TestDefaultPictureSettings()
        {
            HSSFPicture picture = new HSSFPicture(null, new HSSFClientAnchor());
            Assert.AreEqual(picture.LineWidth, HSSFShape.LINEWIDTH_DEFAULT);
            Assert.AreEqual(picture.FillColor, HSSFShape.FILL__FILLCOLOR_DEFAULT);
            Assert.AreEqual(picture.LineStyle, HSSFShape.LINESTYLE_NONE);
            Assert.AreEqual(picture.LineStyleColor, HSSFShape.LINESTYLE__COLOR_DEFAULT);
            Assert.AreEqual(picture.IsNoFill, false);
            Assert.AreEqual(picture.PictureIndex, -1);//not Set yet
        }

        /**
         * No NullPointerException should appear
         */
        public void TestDefaultSettingsWithEmptyContainer()
        {
            EscherContainerRecord Container = new EscherContainerRecord();
            EscherOptRecord opt = new EscherOptRecord();
            opt.RecordId=(EscherOptRecord.RECORD_ID);
            Container.AddChildRecord(opt);
            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord cod = new CommonObjectDataSubRecord();
            cod.ObjectType= (CommonObjectType) (HSSFSimpleShape.OBJECT_TYPE_PICTURE);
            obj.AddSubRecord(cod);
            HSSFPicture picture = new HSSFPicture(Container, obj);

            Assert.AreEqual(picture.LineWidth, HSSFShape.LINEWIDTH_DEFAULT);
            Assert.AreEqual(picture.FillColor, HSSFShape.FILL__FILLCOLOR_DEFAULT);
            Assert.AreEqual(picture.LineStyle, HSSFShape.LINESTYLE_DEFAULT);
            Assert.AreEqual(picture.LineStyleColor, HSSFShape.LINESTYLE__COLOR_DEFAULT);
            Assert.AreEqual(picture.IsNoFill, HSSFShape.NO_FILL_DEFAULT);
            Assert.AreEqual(picture.PictureIndex, -1);//not Set yet
        }

        /**
         * create a rectangle, save the workbook, read back and verify that all shape properties are there
         */
        public void TestReadWriteRectangle()
        {

            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            HSSFPatriarch drawing = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFClientAnchor anchor = new HSSFClientAnchor(10, 10, 50, 50, (short)2, 2, (short)4, 4);
            anchor.AnchorType = (AnchorType)(2);
            Assert.AreEqual(anchor.AnchorType, 2);

            HSSFSimpleShape rectangle = drawing.CreateSimpleShape(anchor);
            rectangle.ShapeType=(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
            rectangle.LineWidth=(10000);
            rectangle.FillColor=(777);
            Assert.AreEqual(rectangle.FillColor, 777);
            Assert.AreEqual(10000, rectangle.LineWidth);
            rectangle.LineStyle= (LineStyle)(10);
            Assert.AreEqual(10, rectangle.LineStyle);
            Assert.AreEqual(rectangle.WrapText, HSSFSimpleShape.WRAP_SQUARE);
            rectangle.LineStyleColor=(1111);
            rectangle.IsNoFill=(true);
            rectangle.WrapText=(HSSFSimpleShape.WRAP_NONE);
            rectangle.String=(new HSSFRichTextString("teeeest"));
            Assert.AreEqual(rectangle.LineStyleColor, 1111);
            Assert.AreEqual(((EscherSimpleProperty)((EscherOptRecord)HSSFTestHelper.GetEscherContainer(rectangle).GetChildById(EscherOptRecord.RECORD_ID))
                    .Lookup(EscherProperties.TEXT__TEXTID)).PropertyValue, "teeeest".GetHashCode());
            Assert.AreEqual(rectangle.IsNoFill, true);
            Assert.AreEqual(rectangle.WrapText, HSSFSimpleShape.WRAP_NONE);
            Assert.AreEqual(rectangle.String.String, "teeeest");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, drawing.Children.Count);

            HSSFSimpleShape rectangle2 =
                    (HSSFSimpleShape)drawing.Children[0];
            Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE,
                    rectangle2.ShapeType);
            Assert.AreEqual(10000, rectangle2.LineWidth);
            Assert.AreEqual(10, (int)rectangle2.LineStyle);
            Assert.AreEqual(anchor, rectangle2.Anchor);
            Assert.AreEqual(rectangle2.LineStyleColor, 1111);
            Assert.AreEqual(rectangle2.FillColor, 777);
            Assert.AreEqual(rectangle2.IsNoFill, true);
            Assert.AreEqual(rectangle2.String.String, "teeeest");
            Assert.AreEqual(rectangle.WrapText, HSSFSimpleShape.WRAP_NONE);

            rectangle2.FillColor=(3333);
            rectangle2.LineStyle = (LineStyle)(9);
            rectangle2.LineStyleColor=(4444);
            rectangle2.IsNoFill=(false);
            rectangle2.LineWidth=(77);
            rectangle2.Anchor.Dx1=2;
            rectangle2.Anchor.Dx2=3;
            rectangle2.Anchor.Dy1=(4);
            rectangle2.Anchor.Dy2=(5);
            rectangle.WrapText=(HSSFSimpleShape.WRAP_BY_POINTS);
            rectangle2.String=(new HSSFRichTextString("test22"));

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, drawing.Children.Count);
            rectangle2 = (HSSFSimpleShape)drawing.Children[0];
            Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE, rectangle2.ShapeType);
            Assert.AreEqual(rectangle.WrapText, HSSFSimpleShape.WRAP_BY_POINTS);
            Assert.AreEqual(77, rectangle2.LineWidth);
            Assert.AreEqual(9, rectangle2.LineStyle);
            Assert.AreEqual(rectangle2.LineStyleColor, 4444);
            Assert.AreEqual(rectangle2.FillColor, 3333);
            Assert.AreEqual(rectangle2.Anchor.Dx1, 2);
            Assert.AreEqual(rectangle2.Anchor.Dx2, 3);
            Assert.AreEqual(rectangle2.Anchor.Dy1, 4);
            Assert.AreEqual(rectangle2.Anchor.Dy2, 5);
            Assert.AreEqual(rectangle2.IsNoFill, false);
            Assert.AreEqual(rectangle2.String.String, "test22");

            HSSFSimpleShape rect3 = drawing.CreateSimpleShape(new HSSFClientAnchor());
            rect3.ShapeType=(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            drawing = (wb.GetSheetAt(0) as HSSFSheet).DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(drawing.Children.Count, 2);
        }

        public void TestReadExistingImage()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("pictures") as HSSFSheet;
            HSSFPatriarch Drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, Drawing.Children.Count);
            HSSFPicture picture = (HSSFPicture)Drawing.Children[0];

            Assert.AreEqual(picture.PictureIndex, 2);
            Assert.AreEqual(picture.LineStyleColor, HSSFShape.LINESTYLE__COLOR_DEFAULT);
            Assert.AreEqual(picture.FillColor, 0x5DC943);
            Assert.AreEqual(picture.LineWidth, HSSFShape.LINEWIDTH_DEFAULT);
            Assert.AreEqual(picture.LineStyle, HSSFShape.LINESTYLE_DEFAULT);
            Assert.AreEqual(picture.IsNoFill, false);

            picture.PictureIndex=(2);
            Assert.AreEqual(picture.PictureIndex, 2);
        }


        /* assert shape properties when Reading shapes from a existing workbook */
        [Test]
        public void TestReadExistingRectangle()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("rectangles") as HSSFSheet;
            HSSFPatriarch Drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, Drawing.Children.Count);

            HSSFSimpleShape shape = (HSSFSimpleShape)Drawing.Children[0];
            Assert.AreEqual(shape.IsNoFill, false);
            Assert.AreEqual((int)shape.LineStyle, HSSFShape.LINESTYLE_DASHDOTGEL);
            Assert.AreEqual(shape.LineStyleColor, 0x616161);
            Assert.AreEqual(shape.FillColor, 0x2CE03D, HexDump.ToHex(shape.FillColor));
            Assert.AreEqual(shape.LineWidth, HSSFShape.LINEWIDTH_ONE_PT * 2);
            Assert.AreEqual(shape.String.String, "POItest");
            Assert.AreEqual(shape.RotationDegree, 27);
        }
        [Test]
        public void TestShapeIds()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet1 = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch1 = sheet1.CreateDrawingPatriarch() as HSSFPatriarch;
            for (int i = 0; i < 2; i++)
            {
                patriarch1.CreateSimpleShape(new HSSFClientAnchor());
            }

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet1 = wb.GetSheetAt(0) as HSSFSheet;
            patriarch1 = sheet1.DrawingPatriarch as HSSFPatriarch;

            EscherAggregate agg1 = HSSFTestHelper.GetEscherAggregate(patriarch1);
            // last shape ID cached in EscherDgRecord
            EscherDgRecord dg1 =
                    agg1.GetEscherContainer().GetChildById(EscherDgRecord.RECORD_ID) as EscherDgRecord;
            Assert.AreEqual(1026, dg1.LastMSOSPID);

            // iterate over shapes and check shapeId
            EscherContainerRecord spgrContainer =
                    agg1.GetEscherContainer().ChildContainers[0] as EscherContainerRecord;
            // root spContainer + 2 spContainers for shapes
            Assert.AreEqual(3, spgrContainer.ChildRecords.Count);

            EscherSpRecord sp0 =
                    ((EscherContainerRecord)spgrContainer.GetChild(0)).GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;
            Assert.AreEqual(1024, sp0.ShapeId);

            EscherSpRecord sp1 =
                    ((EscherContainerRecord)spgrContainer.GetChild(1)).GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;
            Assert.AreEqual(1025, sp1.ShapeId);

            EscherSpRecord sp2 =
                    ((EscherContainerRecord)spgrContainer.GetChild(2)).GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;
            Assert.AreEqual(1026, sp2.ShapeId);
        }

        /**
         * Test Get new id for shapes from existing file
         * File already have for 1 shape on each sheet, because document must contain EscherDgRecord for each sheet
         */
        [Test]
        public void TestAllocateNewIds()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("empty.xls");
            HSSFSheet sheet = wb.GetSheetAt(0) as HSSFSheet;
            HSSFPatriarch patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            /**
             * 2048 - main SpContainer id
             * 2049 - existing shape id
             */
            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 2050);
            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 2051);
            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 2052);

            sheet = wb.GetSheetAt(1) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            /**
             * 3072 - main SpContainer id
             * 3073 - existing shape id
             */
            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 3074);
            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 3075);
            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 3076);


            sheet = wb.GetSheetAt(2) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 1026);
            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 1027);
            Assert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 1028);
        }
        [Test]
        public void TestOpt() {
            HSSFWorkbook wb = new HSSFWorkbook();

            // create a sheet with a text box
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFTextbox textbox = patriarch.CreateTextbox(new HSSFClientAnchor()) as HSSFTextbox;
            EscherOptRecord opt1 = HSSFTestHelper.GetOptRecord(textbox);
            EscherOptRecord opt2 = HSSFTestHelper.GetEscherContainer(textbox).GetChildById(EscherOptRecord.RECORD_ID) as EscherOptRecord;
            Assert.AreSame(opt1, opt2);
        }
        [Test]
        public void TestCorrectOrderInOptRecord()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFTextbox textbox = patriarch.CreateTextbox(new HSSFClientAnchor()) as HSSFTextbox;
            EscherOptRecord opt = HSSFTestHelper.GetOptRecord(textbox);

            String opt1Str = opt.ToXml();

            textbox.FillColor = textbox.FillColor;
            EscherContainerRecord Container = HSSFTestHelper.GetEscherContainer(textbox);
            EscherOptRecord optRecord = Container.GetChildById(EscherOptRecord.RECORD_ID) as EscherOptRecord;
            Assert.AreEqual(opt1Str, optRecord.ToXml());
            textbox.LineStyle = textbox.LineStyle;
            Assert.AreEqual(opt1Str, optRecord.ToXml());
            textbox.LineWidth = textbox.LineWidth;
            Assert.AreEqual(opt1Str, optRecord.ToXml());
            textbox.LineStyleColor = textbox.LineStyleColor;
            Assert.AreEqual(opt1Str, optRecord.ToXml());
        }
        [Test]
        public void TestDgRecordNumShapes()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            EscherAggregate aggregate = HSSFTestHelper.GetEscherAggregate(patriarch);
            EscherDgRecord dgRecord = (EscherDgRecord)aggregate.GetEscherRecord(0).GetChild(0) as EscherDgRecord;
            Assert.AreEqual(dgRecord.NumShapes, 1);
        }

        public void TestTextForSimpleShape()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape shape = patriarch.CreateSimpleShape(new HSSFClientAnchor());
            shape.ShapeType = HSSFSimpleShape.OBJECT_TYPE_RECTANGLE;

            EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);
            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            shape = (HSSFSimpleShape)patriarch.Children[0];

            agg = HSSFTestHelper.GetEscherAggregate(patriarch);
            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

            shape.String = new HSSFRichTextString("string1");
            Assert.AreEqual(shape.String.String, "string1");

            Assert.IsNotNull(HSSFTestHelper.GetEscherContainer(shape).GetChildById(EscherTextboxRecord.RECORD_ID));
            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            shape = (HSSFSimpleShape)patriarch.Children[0];

            Assert.IsNotNull(HSSFTestHelper.GetTextObjRecord(shape));
            Assert.AreEqual(shape.String.String, "string1");
            Assert.IsNotNull(HSSFTestHelper.GetEscherContainer(shape).GetChildById(EscherTextboxRecord.RECORD_ID));
            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 2);
        }
        [Test]
        public void TestRemoveShapes()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor());
            rectangle.ShapeType = HSSFSimpleShape.OBJECT_TYPE_RECTANGLE;

            int idx = wb.AddPicture(new byte[] { 1, 2, 3 }, PictureType.JPEG);
            patriarch.CreatePicture(new HSSFClientAnchor(), idx);

            patriarch.CreateCellComment(new HSSFClientAnchor());

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPoints(new int[] { 1, 2 }, new int[] { 2, 3 });

            patriarch.CreateTextbox(new HSSFClientAnchor());

            HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
            group.CreateTextbox(new HSSFChildAnchor());
            group.CreatePicture(new HSSFChildAnchor(), idx);

            Assert.AreEqual(patriarch.Children.Count, 6);
            Assert.AreEqual(group.Children.Count, 2);

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 12);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 12);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            Assert.AreEqual(patriarch.Children.Count, 6);

            group = (HSSFShapeGroup)patriarch.Children[5];
            group.RemoveShape(group.Children[0]);

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 10);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 10);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            group = (HSSFShapeGroup)patriarch.Children[(5)];
            patriarch.RemoveShape(group);

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 8);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 8);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            Assert.AreEqual(patriarch.Children.Count, 5);

            HSSFShape shape = patriarch.Children[0];
            patriarch.RemoveShape(shape);

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 6);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            Assert.AreEqual(patriarch.Children.Count, 4);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 6);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            Assert.AreEqual(patriarch.Children.Count, 4);

            HSSFPicture picture = (HSSFPicture)patriarch.Children[0];
            patriarch.RemoveShape(picture);

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 5);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            Assert.AreEqual(patriarch.Children.Count, 3);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 5);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            Assert.AreEqual(patriarch.Children.Count, 3);

            HSSFComment comment = (HSSFComment)patriarch.Children[0];
            patriarch.RemoveShape(comment);

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 3);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            Assert.AreEqual(patriarch.Children.Count, 2);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 3);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            Assert.AreEqual(patriarch.Children.Count, 2);

            polygon = (HSSFPolygon)patriarch.Children[0];
            patriarch.RemoveShape(polygon);

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 2);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            Assert.AreEqual(patriarch.Children.Count, 1);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 2);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            Assert.AreEqual(patriarch.Children.Count, 1);

            HSSFTextbox textbox = (HSSFTextbox)patriarch.Children[0];
            patriarch.RemoveShape(textbox);

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 0);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            Assert.AreEqual(patriarch.Children.Count, 0);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 0);
            Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            Assert.AreEqual(patriarch.Children.Count, 0);
        }
        [Test]
        public void TestShapeFlip()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor());
            rectangle.ShapeType = HSSFSimpleShape.OBJECT_TYPE_RECTANGLE;

            Assert.AreEqual(rectangle.IsFlipVertical, false);
            Assert.AreEqual(rectangle.IsFlipHorizontal, false);

            rectangle.IsFlipVertical = true;
            Assert.AreEqual(rectangle.IsFlipVertical, true);
            rectangle.IsFlipHorizontal = true;
            Assert.AreEqual(rectangle.IsFlipHorizontal, true);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            rectangle = (HSSFSimpleShape)patriarch.Children[0];

            Assert.AreEqual(rectangle.IsFlipHorizontal, true);
            rectangle.IsFlipHorizontal = false;
            Assert.AreEqual(rectangle.IsFlipHorizontal, false);

            Assert.AreEqual(rectangle.IsFlipVertical, true);
            rectangle.IsFlipVertical = false;
            Assert.AreEqual(rectangle.IsFlipVertical, false);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            rectangle = (HSSFSimpleShape)patriarch.Children[0];

            Assert.AreEqual(rectangle.IsFlipVertical, false);
            Assert.AreEqual(rectangle.IsFlipHorizontal, false);
        }
        [Test]
        public void TestRotation()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor(0, 0, 100, 100, (short)0, 0, (short)5, 5));
            rectangle.ShapeType = HSSFSimpleShape.OBJECT_TYPE_RECTANGLE;

            Assert.AreEqual(rectangle.RotationDegree, 0);
            rectangle.RotationDegree = (short)45;
            Assert.AreEqual(rectangle.RotationDegree, 45);
            rectangle.IsFlipHorizontal = true;

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            rectangle = (HSSFSimpleShape)patriarch.Children[0];
            Assert.AreEqual(rectangle.RotationDegree, 45);
            rectangle.RotationDegree = (short)30;
            Assert.AreEqual(rectangle.RotationDegree, 30);

            patriarch.SetCoordinates(0, 0, 10, 10);
            rectangle.String = new HSSFRichTextString("1234");
        }
        [Test]
        public void TestShapeContainerImplementsIterable(){
            HSSFWorkbook wb = new HSSFWorkbook();
            try
            {
                HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
                HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

                patriarch.CreateSimpleShape(new HSSFClientAnchor());
                patriarch.CreateSimpleShape(new HSSFClientAnchor());

                int i = 2;

                foreach (HSSFShape shape in patriarch)
                {
                    i--;
                }
                Assert.AreEqual(i, 0);
            }
            finally
            {
                //wb.Close();
            }
        }
        [Test]
        public void TestClearShapesForPatriarch()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            patriarch.CreateSimpleShape(new HSSFClientAnchor());
            patriarch.CreateSimpleShape(new HSSFClientAnchor());
            patriarch.CreateCellComment(new HSSFClientAnchor());

            EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);

            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 6);
            Assert.AreEqual(agg.TailRecords.Count, 1);
            Assert.AreEqual(patriarch.Children.Count, 3);

            patriarch.Clear();

            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 0);
            Assert.AreEqual(agg.TailRecords.Count, 0);
            Assert.AreEqual(patriarch.Children.Count, 0);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            Assert.AreEqual(agg.GetShapeToObjMapping().Count, 0);
            Assert.AreEqual(agg.TailRecords.Count, 0);
            Assert.AreEqual(patriarch.Children.Count, 0);
        }

        [Test]
        public void TestBug45312()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            try
            {
                HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
                HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

                {
                    HSSFClientAnchor a1 = new HSSFClientAnchor();
                    a1.SetAnchor((short)1, 1, 0, 0, (short)1, 1, 512, 100);
                    HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                    shape1.ShapeType = (/*setter*/HSSFSimpleShape.OBJECT_TYPE_LINE);
                }

                {
                    HSSFClientAnchor a1 = new HSSFClientAnchor();
                    //setAnchor method is wrong??
                    a1.SetAnchor((short)1, 1, 512, 0, (short)1, 1, 1023, 100);
                    HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                    shape1.FlipVertical=(true);
                    shape1.ShapeType = (/*setter*/HSSFSimpleShape.OBJECT_TYPE_LINE);
                }

                {
                    HSSFClientAnchor a1 = new HSSFClientAnchor();
                    a1.SetAnchor((short)2, 2, 0, 0, (short)2, 2, 512, 100);
                    HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                    shape1.ShapeType = (/*setter*/HSSFSimpleShape.OBJECT_TYPE_LINE);
                }
                {
                    HSSFClientAnchor a1 = new HSSFClientAnchor();
                    a1.SetAnchor((short)2, 2, 0, 100, (short)2, 2, 512, 200);
                    HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                    shape1.FlipHorizontal = (/*setter*/true);
                    shape1.ShapeType = (/*setter*/HSSFSimpleShape.OBJECT_TYPE_LINE);
                }

                /*OutputStream stream = new FileOutputStream("/tmp/45312.xls");
                try {
                    wb.Write(stream);
                } finally {
                    stream.Close();
                }*/

                CheckWorkbookBack(wb);
            }
            finally
            {
                //wb.Close();
            }
        }

        private void CheckWorkbookBack(HSSFWorkbook wb)
        {
            HSSFWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);

            HSSFSheet sheetBack = wbBack.GetSheetAt(0) as HSSFSheet;
            Assert.IsNotNull(sheetBack);

            HSSFPatriarch patriarchBack = sheetBack.DrawingPatriarch as HSSFPatriarch;
            Assert.IsNotNull(patriarchBack);

            IList<HSSFShape> children = patriarchBack.Children;
            Assert.AreEqual(4, children.Count);
            HSSFShape hssfShape = children[(0)];
            Assert.IsTrue(hssfShape is HSSFSimpleShape);
            HSSFAnchor anchor = hssfShape.Anchor;
            Assert.IsTrue(anchor is HSSFClientAnchor);
            Assert.AreEqual(0, anchor.Dx1);
            Assert.AreEqual(512, anchor.Dx2);
            Assert.AreEqual(0, anchor.Dy1);
            Assert.AreEqual(100, anchor.Dy2);
            HSSFClientAnchor cAnchor = (HSSFClientAnchor)anchor;
            Assert.AreEqual(1, cAnchor.Col1);
            Assert.AreEqual(1, cAnchor.Col2);
            Assert.AreEqual(1, cAnchor.Row1);
            Assert.AreEqual(1, cAnchor.Row2);

            hssfShape = children[(1)];
            Assert.IsTrue(hssfShape is HSSFSimpleShape);
            anchor = hssfShape.Anchor;
            Assert.IsTrue(anchor is HSSFClientAnchor);
            Assert.AreEqual(512, anchor.Dx1);
            Assert.AreEqual(1023, anchor.Dx2);
            Assert.AreEqual(0, anchor.Dy1);
            Assert.AreEqual(100, anchor.Dy2);
            cAnchor = (HSSFClientAnchor)anchor;
            Assert.AreEqual(1, cAnchor.Col1);
            Assert.AreEqual(1, cAnchor.Col2);
            Assert.AreEqual(1, cAnchor.Row1);
            Assert.AreEqual(1, cAnchor.Row2);

            hssfShape = children[(2)];
            Assert.IsTrue(hssfShape is HSSFSimpleShape);
            anchor = hssfShape.Anchor;
            Assert.IsTrue(anchor is HSSFClientAnchor);
            Assert.AreEqual(0, anchor.Dx1);
            Assert.AreEqual(512, anchor.Dx2);
            Assert.AreEqual(0, anchor.Dy1);
            Assert.AreEqual(100, anchor.Dy2);
            cAnchor = (HSSFClientAnchor)anchor;
            Assert.AreEqual(2, cAnchor.Col1);
            Assert.AreEqual(2, cAnchor.Col2);
            Assert.AreEqual(2, cAnchor.Row1);
            Assert.AreEqual(2, cAnchor.Row2);

            hssfShape = children[(3)];
            Assert.IsTrue(hssfShape is HSSFSimpleShape);
            anchor = hssfShape.Anchor;
            Assert.IsTrue(anchor is HSSFClientAnchor);
            Assert.AreEqual(0, anchor.Dx1);
            Assert.AreEqual(512, anchor.Dx2);
            Assert.AreEqual(100, anchor.Dy1);
            Assert.AreEqual(200, anchor.Dy2);
            cAnchor = (HSSFClientAnchor)anchor;
            Assert.AreEqual(2, cAnchor.Col1);
            Assert.AreEqual(2, cAnchor.Col2);
            Assert.AreEqual(2, cAnchor.Row1);
            Assert.AreEqual(2, cAnchor.Row2);
        }

    }
}
