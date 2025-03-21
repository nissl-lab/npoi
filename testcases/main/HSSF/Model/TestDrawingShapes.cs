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
using NPOI.DDF;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using NUnit.Framework;using NUnit.Framework.Legacy;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using TestCases.HSSF.UserModel;

namespace TestCases.HSSF.Model
{
    /**
     * Test escher drawing
     * 
     * optionally the system setting "poi.deserialize.escher" can be set to {@code true}
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
            ClassicAssert.AreEqual(patriarch.Children.Count, 2);
            HSSFShapeGroup group = (HSSFShapeGroup)patriarch.Children[1];
            ClassicAssert.AreEqual(3, group.Children.Count);
            HSSFShapeGroup group1 = (HSSFShapeGroup)group.Children[0];
            ClassicAssert.AreEqual(2, group1.Children.Count);
            group1 = (HSSFShapeGroup)group.Children[2];
            ClassicAssert.AreEqual(2, group1.Children.Count);

            wb.Close();
        }

        public void TestHSSFShapeCompatibility()
        {
            HSSFSimpleShape shape = new HSSFSimpleShape(null, new HSSFClientAnchor());
            shape.ShapeType=(HSSFSimpleShape.OBJECT_TYPE_LINE);
            ClassicAssert.AreEqual(0x08000040, shape.LineStyleColor);
            ClassicAssert.AreEqual(0x08000009, shape.FillColor);
            ClassicAssert.AreEqual(HSSFShape.LINEWIDTH_DEFAULT, shape.LineWidth);
            ClassicAssert.AreEqual(HSSFShape.LINESTYLE_SOLID, shape.LineStyle);
            ClassicAssert.IsFalse(shape.IsNoFill);

            EscherOptRecord opt = shape.GetOptRecord();
            ClassicAssert.AreEqual(7, opt.EscherProperties.Count);
            ClassicAssert.IsTrue(((EscherBoolProperty)opt.Lookup(EscherProperties.GROUPSHAPE__PRINT)).IsTrue);
            ClassicAssert.IsTrue(((EscherBoolProperty)opt.Lookup(EscherProperties.LINESTYLE__NOLINEDRAWDASH)).IsTrue);
            ClassicAssert.AreEqual(0x00000004, ((EscherSimpleProperty)opt.Lookup(EscherProperties.GEOMETRY__SHAPEPATH)).PropertyValue);
            ClassicAssert.IsNull(opt.Lookup(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE));
        }

        public void TestDefaultPictureSettings()
        {
            HSSFPicture picture = new HSSFPicture(null, new HSSFClientAnchor());
            ClassicAssert.AreEqual(picture.LineWidth, HSSFShape.LINEWIDTH_DEFAULT);
            ClassicAssert.AreEqual(picture.FillColor, HSSFShape.FILL__FILLCOLOR_DEFAULT);
            ClassicAssert.AreEqual(picture.LineStyle, HSSFShape.LINESTYLE_NONE);
            ClassicAssert.AreEqual(picture.LineStyleColor, HSSFShape.LINESTYLE__COLOR_DEFAULT);
            ClassicAssert.IsFalse(picture.IsNoFill);
            ClassicAssert.AreEqual(picture.PictureIndex, -1);//not Set yet
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

            ClassicAssert.AreEqual(picture.LineWidth, HSSFShape.LINEWIDTH_DEFAULT);
            ClassicAssert.AreEqual(picture.FillColor, HSSFShape.FILL__FILLCOLOR_DEFAULT);
            ClassicAssert.AreEqual(picture.LineStyle, HSSFShape.LINESTYLE_DEFAULT);
            ClassicAssert.AreEqual(picture.LineStyleColor, HSSFShape.LINESTYLE__COLOR_DEFAULT);
            ClassicAssert.AreEqual(picture.IsNoFill, HSSFShape.NO_FILL_DEFAULT);
            ClassicAssert.AreEqual(picture.PictureIndex, -1);//not Set yet
        }

        /**
         * create a rectangle, save the workbook, read back and verify that all shape properties are there
         */
        public void TestReadWriteRectangle()
        {

            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sheet = wb1.CreateSheet() as HSSFSheet;

            HSSFPatriarch drawing = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFClientAnchor anchor = new HSSFClientAnchor(10, 10, 50, 50, (short)2, 2, (short)4, 4);
            anchor.AnchorType = AnchorType.MoveDontResize;
            ClassicAssert.AreEqual(AnchorType.MoveDontResize, anchor.AnchorType);

            //noinspection deprecation
            //anchor.AnchorType = (AnchorType.MoveDontResize.value);
            //ClassicAssert.AreEqual(AnchorType.MoveDontResize, anchor.AnchorType);

            HSSFSimpleShape rectangle = drawing.CreateSimpleShape(anchor);
            rectangle.ShapeType=(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
            rectangle.LineWidth=(10000);
            rectangle.FillColor=(777);
            ClassicAssert.AreEqual(rectangle.FillColor, 777);
            ClassicAssert.AreEqual(10000, rectangle.LineWidth);
            rectangle.LineStyle= (LineStyle)(10);
            ClassicAssert.AreEqual(10, rectangle.LineStyle);
            ClassicAssert.AreEqual(rectangle.WrapText, HSSFSimpleShape.WRAP_SQUARE);
            rectangle.LineStyleColor=(1111);
            rectangle.IsNoFill=(true);
            rectangle.WrapText=(HSSFSimpleShape.WRAP_NONE);
            rectangle.String=(new HSSFRichTextString("teeeest"));
            ClassicAssert.AreEqual(rectangle.LineStyleColor, 1111);
            //ClassicAssert.AreEqual(((EscherSimpleProperty)((EscherOptRecord)HSSFTestHelper.GetEscherContainer(rectangle).GetChildById(EscherOptRecord.RECORD_ID))
            //        .Lookup(EscherProperties.TEXT__TEXTID)).PropertyValue, "teeeest".GetHashCode());
            EscherContainerRecord escherContainer = HSSFTestHelper.GetEscherContainer(rectangle);
            ClassicAssert.IsNotNull(escherContainer);
            EscherRecord childById = escherContainer.GetChildById(EscherOptRecord.RECORD_ID);
            ClassicAssert.IsNotNull(childById);
            EscherProperty lookup = ((EscherOptRecord)childById).Lookup(EscherProperties.TEXT__TEXTID);
            ClassicAssert.IsNotNull(lookup);
            ClassicAssert.AreEqual("teeeest".GetHashCode(), ((EscherSimpleProperty)lookup).PropertyValue);

            ClassicAssert.AreEqual(rectangle.IsNoFill, true);
            ClassicAssert.AreEqual(rectangle.WrapText, HSSFSimpleShape.WRAP_NONE);
            ClassicAssert.AreEqual(rectangle.String.String, "teeeest");

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as HSSFSheet;
            drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            ClassicAssert.AreEqual(1, drawing.Children.Count);

            HSSFSimpleShape rectangle2 =
                    (HSSFSimpleShape)drawing.Children[0];
            ClassicAssert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE,
                    rectangle2.ShapeType);
            ClassicAssert.AreEqual(10000, rectangle2.LineWidth);
            ClassicAssert.AreEqual(10, (int)rectangle2.LineStyle);
            ClassicAssert.AreEqual(anchor, rectangle2.Anchor);
            ClassicAssert.AreEqual(rectangle2.LineStyleColor, 1111);
            ClassicAssert.AreEqual(rectangle2.FillColor, 777);
            ClassicAssert.AreEqual(rectangle2.IsNoFill, true);
            ClassicAssert.AreEqual(rectangle2.String.String, "teeeest");
            ClassicAssert.AreEqual(rectangle.WrapText, HSSFSimpleShape.WRAP_NONE);

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

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();
            sheet = wb3.GetSheetAt(0) as HSSFSheet;
            drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            ClassicAssert.AreEqual(1, drawing.Children.Count);
            rectangle2 = (HSSFSimpleShape)drawing.Children[0];
            ClassicAssert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE, rectangle2.ShapeType);
            ClassicAssert.AreEqual(rectangle.WrapText, HSSFSimpleShape.WRAP_BY_POINTS);
            ClassicAssert.AreEqual(77, rectangle2.LineWidth);
            ClassicAssert.AreEqual(9, rectangle2.LineStyle);
            ClassicAssert.AreEqual(rectangle2.LineStyleColor, 4444);
            ClassicAssert.AreEqual(rectangle2.FillColor, 3333);
            ClassicAssert.AreEqual(rectangle2.Anchor.Dx1, 2);
            ClassicAssert.AreEqual(rectangle2.Anchor.Dx2, 3);
            ClassicAssert.AreEqual(rectangle2.Anchor.Dy1, 4);
            ClassicAssert.AreEqual(rectangle2.Anchor.Dy2, 5);
            ClassicAssert.AreEqual(rectangle2.IsNoFill, false);
            ClassicAssert.AreEqual(rectangle2.String.String, "test22");

            HSSFSimpleShape rect3 = drawing.CreateSimpleShape(new HSSFClientAnchor());
            rect3.ShapeType=(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
            HSSFWorkbook wb4 = HSSFTestDataSamples.WriteOutAndReadBack(wb3);
            wb3.Close();

            drawing = (wb4.GetSheetAt(0) as HSSFSheet).DrawingPatriarch as HSSFPatriarch;
            ClassicAssert.AreEqual(drawing.Children.Count, 2);
            wb4.Close();
        }

        public void TestReadExistingImage()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("pictures") as HSSFSheet;
            HSSFPatriarch Drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            ClassicAssert.AreEqual(1, Drawing.Children.Count);
            HSSFPicture picture = (HSSFPicture)Drawing.Children[0];

            ClassicAssert.AreEqual(picture.PictureIndex, 2);
            ClassicAssert.AreEqual(picture.LineStyleColor, HSSFShape.LINESTYLE__COLOR_DEFAULT);
            ClassicAssert.AreEqual(picture.FillColor, 0x5DC943);
            ClassicAssert.AreEqual(picture.LineWidth, HSSFShape.LINEWIDTH_DEFAULT);
            ClassicAssert.AreEqual(picture.LineStyle, HSSFShape.LINESTYLE_DEFAULT);
            ClassicAssert.AreEqual(picture.IsNoFill, false);

            picture.PictureIndex=(2);
            ClassicAssert.AreEqual(picture.PictureIndex, 2);
            wb.Close();

        }


        /* assert shape properties when Reading shapes from a existing workbook */
        [Test]
        public void TestReadExistingRectangle()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("rectangles") as HSSFSheet;
            HSSFPatriarch Drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            ClassicAssert.AreEqual(1, Drawing.Children.Count);

            HSSFSimpleShape shape = (HSSFSimpleShape)Drawing.Children[0];
            ClassicAssert.AreEqual(shape.IsNoFill, false);
            ClassicAssert.AreEqual((int)shape.LineStyle, HSSFShape.LINESTYLE_DASHDOTGEL);
            ClassicAssert.AreEqual(shape.LineStyleColor, 0x616161);
            ClassicAssert.AreEqual(shape.FillColor, 0x2CE03D, HexDump.ToHex(shape.FillColor));
            ClassicAssert.AreEqual(shape.LineWidth, HSSFShape.LINEWIDTH_ONE_PT * 2);
            ClassicAssert.AreEqual(shape.String.String, "POItest");
            ClassicAssert.AreEqual(shape.RotationDegree, 27);
            wb.Close();
        }
        [Test]
        public void TestShapeIds()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sheet1 = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch1 = sheet1.CreateDrawingPatriarch() as HSSFPatriarch;
            for (int i = 0; i < 2; i++)
            {
                patriarch1.CreateSimpleShape(new HSSFClientAnchor());
            }

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet1 = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch1 = sheet1.DrawingPatriarch as HSSFPatriarch;

            EscherAggregate agg1 = HSSFTestHelper.GetEscherAggregate(patriarch1);
            // last shape ID cached in EscherDgRecord
            EscherDgRecord dg1 =
                    agg1.GetEscherContainer().GetChildById(EscherDgRecord.RECORD_ID) as EscherDgRecord;
            ClassicAssert.AreEqual(1026, dg1.LastMSOSPID);

            // iterate over shapes and check shapeId
            EscherContainerRecord spgrContainer =
                    agg1.GetEscherContainer().ChildContainers[0] as EscherContainerRecord;
            // root spContainer + 2 spContainers for shapes
            ClassicAssert.AreEqual(3, spgrContainer.ChildRecords.Count);

            EscherSpRecord sp0 =
                    ((EscherContainerRecord)spgrContainer.GetChild(0)).GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;
            ClassicAssert.AreEqual(1024, sp0.ShapeId);

            EscherSpRecord sp1 =
                    ((EscherContainerRecord)spgrContainer.GetChild(1)).GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;
            ClassicAssert.AreEqual(1025, sp1.ShapeId);

            EscherSpRecord sp2 =
                    ((EscherContainerRecord)spgrContainer.GetChild(2)).GetChildById(EscherSpRecord.RECORD_ID) as EscherSpRecord;
            ClassicAssert.AreEqual(1026, sp2.ShapeId);
            wb2.Close();
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
            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 2050);
            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 2051);
            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 2052);

            sheet = wb.GetSheetAt(1) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            /**
             * 3072 - main SpContainer id
             * 3073 - existing shape id
             */
            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 3074);
            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 3075);
            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 3076);


            sheet = wb.GetSheetAt(2) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 1026);
            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 1027);
            ClassicAssert.AreEqual(HSSFTestHelper.AllocateNewShapeId(patriarch), 1028);
            wb.Close();
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
            ClassicAssert.AreSame(opt1, opt2);
            wb.Close();
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
            ClassicAssert.AreEqual(opt1Str, optRecord.ToXml());
            textbox.LineStyle = textbox.LineStyle;
            ClassicAssert.AreEqual(opt1Str, optRecord.ToXml());
            textbox.LineWidth = textbox.LineWidth;
            ClassicAssert.AreEqual(opt1Str, optRecord.ToXml());
            textbox.LineStyleColor = textbox.LineStyleColor;
            ClassicAssert.AreEqual(opt1Str, optRecord.ToXml());

            wb.Close();
        }
        [Test]
        public void TestDgRecordNumShapes()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            EscherAggregate aggregate = HSSFTestHelper.GetEscherAggregate(patriarch);
            EscherDgRecord dgRecord = (EscherDgRecord)aggregate.GetEscherRecord(0).GetChild(0) as EscherDgRecord;
            ClassicAssert.AreEqual(dgRecord.NumShapes, 1);

            wb.Close();
        }

        public void TestTextForSimpleShape()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sheet = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape shape = patriarch.CreateSimpleShape(new HSSFClientAnchor());
            shape.ShapeType = HSSFSimpleShape.OBJECT_TYPE_RECTANGLE;

            EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);
            ClassicAssert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            shape = (HSSFSimpleShape)patriarch.Children[0];

            agg = HSSFTestHelper.GetEscherAggregate(patriarch);
            ClassicAssert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

            shape.String = new HSSFRichTextString("string1");
            ClassicAssert.AreEqual(shape.String.String, "string1");

            ClassicAssert.IsNotNull(HSSFTestHelper.GetEscherContainer(shape).GetChildById(EscherTextboxRecord.RECORD_ID));
            ClassicAssert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();

            HSSFWorkbook wb4 = HSSFTestDataSamples.WriteOutAndReadBack(wb3);
            wb3.Close();

            sheet = wb4.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            shape = (HSSFSimpleShape)patriarch.Children[0];

            ClassicAssert.IsNotNull(HSSFTestHelper.GetTextObjRecord(shape));
            ClassicAssert.AreEqual(shape.String.String, "string1");
            ClassicAssert.IsNotNull(HSSFTestHelper.GetEscherContainer(shape).GetChildById(EscherTextboxRecord.RECORD_ID));
            ClassicAssert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

            wb4.Close();
        }
        [Test]
        public void TestRemoveShapes()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sheet = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor());
            rectangle.ShapeType = HSSFSimpleShape.OBJECT_TYPE_RECTANGLE;

            int idx = wb1.AddPicture(new byte[] { 1, 2, 3 }, PictureType.JPEG);
            patriarch.CreatePicture(new HSSFClientAnchor(), idx);

            patriarch.CreateCellComment(new HSSFClientAnchor());

            HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
            polygon.SetPoints(new int[] { 1, 2 }, new int[] { 2, 3 });

            patriarch.CreateTextbox(new HSSFClientAnchor());

            HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
            group.CreateTextbox(new HSSFChildAnchor());
            group.CreatePicture(new HSSFChildAnchor(), idx);

            ClassicAssert.AreEqual(patriarch.Children.Count, 6);
            ClassicAssert.AreEqual(group.Children.Count, 2);

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 12);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 12);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            ClassicAssert.AreEqual(patriarch.Children.Count, 6);

            group = (HSSFShapeGroup)patriarch.Children[5];
            group.RemoveShape(group.Children[0]);

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 10);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();
            sheet = wb3.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 10);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            group = (HSSFShapeGroup)patriarch.Children[(5)];
            patriarch.RemoveShape(group);

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 8);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);

            HSSFWorkbook wb4 = HSSFTestDataSamples.WriteOutAndReadBack(wb3);
            wb3.Close();
            sheet = wb4.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 8);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            ClassicAssert.AreEqual(patriarch.Children.Count, 5);

            HSSFShape shape = patriarch.Children[0];
            patriarch.RemoveShape(shape);

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 6);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            ClassicAssert.AreEqual(patriarch.Children.Count, 4);

            HSSFWorkbook wb5 = HSSFTestDataSamples.WriteOutAndReadBack(wb4);
            wb4.Close();
            sheet = wb5.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 6);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            ClassicAssert.AreEqual(patriarch.Children.Count, 4);

            HSSFPicture picture = (HSSFPicture)patriarch.Children[0];
            patriarch.RemoveShape(picture);

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 5);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            ClassicAssert.AreEqual(patriarch.Children.Count, 3);

            HSSFWorkbook wb6 = HSSFTestDataSamples.WriteOutAndReadBack(wb5);
            wb5.Close();
            sheet = wb6.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 5);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 1);
            ClassicAssert.AreEqual(patriarch.Children.Count, 3);

            HSSFComment comment = (HSSFComment)patriarch.Children[0];
            patriarch.RemoveShape(comment);

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 3);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            ClassicAssert.AreEqual(patriarch.Children.Count, 2);

            HSSFWorkbook wb7 = HSSFTestDataSamples.WriteOutAndReadBack(wb6);
            wb6.Close();
            sheet = wb7.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 3);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            ClassicAssert.AreEqual(patriarch.Children.Count, 2);

            polygon = (HSSFPolygon)patriarch.Children[0];
            patriarch.RemoveShape(polygon);

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 2);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            ClassicAssert.AreEqual(patriarch.Children.Count, 1);

            HSSFWorkbook wb8 = HSSFTestDataSamples.WriteOutAndReadBack(wb7);
            wb7.Close();
            sheet = wb8.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 2);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            ClassicAssert.AreEqual(patriarch.Children.Count, 1);

            HSSFTextbox textbox = (HSSFTextbox)patriarch.Children[0];
            patriarch.RemoveShape(textbox);

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 0);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            ClassicAssert.AreEqual(patriarch.Children.Count, 0);

            HSSFWorkbook wb9 = HSSFTestDataSamples.WriteOutAndReadBack(wb8);
            wb8.Close();
            sheet = wb9.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 0);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).TailRecords.Count, 0);
            ClassicAssert.AreEqual(patriarch.Children.Count, 0);
            
            wb9.Close();
        }
        [Test]
        public void TestShapeFlip()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sheet = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor());
            rectangle.ShapeType = HSSFSimpleShape.OBJECT_TYPE_RECTANGLE;

            ClassicAssert.AreEqual(rectangle.IsFlipVertical, false);
            ClassicAssert.AreEqual(rectangle.IsFlipHorizontal, false);

            rectangle.IsFlipVertical = true;
            ClassicAssert.AreEqual(rectangle.IsFlipVertical, true);
            rectangle.IsFlipHorizontal = true;
            ClassicAssert.AreEqual(rectangle.IsFlipHorizontal, true);

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            rectangle = (HSSFSimpleShape)patriarch.Children[0];

            ClassicAssert.AreEqual(rectangle.IsFlipHorizontal, true);
            rectangle.IsFlipHorizontal = false;
            ClassicAssert.AreEqual(rectangle.IsFlipHorizontal, false);

            ClassicAssert.AreEqual(rectangle.IsFlipVertical, true);
            rectangle.IsFlipVertical = false;
            ClassicAssert.AreEqual(rectangle.IsFlipVertical, false);

            HSSFWorkbook wb3 = HSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();
            sheet = wb3.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            rectangle = (HSSFSimpleShape)patriarch.Children[0];

            ClassicAssert.AreEqual(rectangle.IsFlipVertical, false);
            ClassicAssert.AreEqual(rectangle.IsFlipHorizontal, false);

            wb3.Close();
        }
        [Test]
        public void TestRotation()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sheet = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor(0, 0, 100, 100, (short)0, 0, (short)5, 5));
            rectangle.ShapeType = HSSFSimpleShape.OBJECT_TYPE_RECTANGLE;

            ClassicAssert.AreEqual(rectangle.RotationDegree, 0);
            rectangle.RotationDegree = (short)45;
            ClassicAssert.AreEqual(rectangle.RotationDegree, 45);
            rectangle.IsFlipHorizontal = true;

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            sheet = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;
            rectangle = (HSSFSimpleShape)patriarch.Children[0];
            ClassicAssert.AreEqual(rectangle.RotationDegree, 45);
            rectangle.RotationDegree = (short)30;
            ClassicAssert.AreEqual(rectangle.RotationDegree, 30);

            patriarch.SetCoordinates(0, 0, 10, 10);
            rectangle.String = new HSSFRichTextString("1234");
            wb2.Close();
        }
        [Test]
        public void TestShapeContainerImplementsIterable(){
            HSSFWorkbook wb = new HSSFWorkbook();

            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            patriarch.CreateSimpleShape(new HSSFClientAnchor());
            patriarch.CreateSimpleShape(new HSSFClientAnchor());

            int i = 2;

            foreach (HSSFShape shape in patriarch)
            {
                i--;
            }
            ClassicAssert.AreEqual(i, 0);
            wb.Close();
        }
        [Test]
        public void TestClearShapesForPatriarch()
        {
            HSSFWorkbook wb1 = new HSSFWorkbook();
            HSSFSheet sheet = wb1.CreateSheet() as HSSFSheet;
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            patriarch.CreateSimpleShape(new HSSFClientAnchor());
            patriarch.CreateSimpleShape(new HSSFClientAnchor());
            patriarch.CreateCellComment(new HSSFClientAnchor());

            EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);

            ClassicAssert.AreEqual(agg.GetShapeToObjMapping().Count, 6);
            ClassicAssert.AreEqual(agg.TailRecords.Count, 1);
            ClassicAssert.AreEqual(patriarch.Children.Count, 3);

            patriarch.Clear();

            ClassicAssert.AreEqual(agg.GetShapeToObjMapping().Count, 0);
            ClassicAssert.AreEqual(agg.TailRecords.Count, 0);
            ClassicAssert.AreEqual(patriarch.Children.Count, 0);

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as HSSFSheet;
            patriarch = sheet.DrawingPatriarch as HSSFPatriarch;

            ClassicAssert.AreEqual(agg.GetShapeToObjMapping().Count, 0);
            ClassicAssert.AreEqual(agg.TailRecords.Count, 0);
            ClassicAssert.AreEqual(patriarch.Children.Count, 0);
            wb2.Close();
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
            ClassicAssert.IsNotNull(wbBack);

            HSSFSheet sheetBack = wbBack.GetSheetAt(0) as HSSFSheet;
            ClassicAssert.IsNotNull(sheetBack);

            HSSFPatriarch patriarchBack = sheetBack.DrawingPatriarch as HSSFPatriarch;
            ClassicAssert.IsNotNull(patriarchBack);

            IList<HSSFShape> children = patriarchBack.Children;
            ClassicAssert.AreEqual(4, children.Count);
            HSSFShape hssfShape = children[(0)];
            ClassicAssert.IsTrue(hssfShape is HSSFSimpleShape);
            HSSFAnchor anchor = hssfShape.Anchor as HSSFAnchor;
            ClassicAssert.IsTrue(anchor is HSSFClientAnchor);
            ClassicAssert.AreEqual(0, anchor.Dx1);
            ClassicAssert.AreEqual(512, anchor.Dx2);
            ClassicAssert.AreEqual(0, anchor.Dy1);
            ClassicAssert.AreEqual(100, anchor.Dy2);
            HSSFClientAnchor cAnchor = (HSSFClientAnchor)anchor;
            ClassicAssert.AreEqual(1, cAnchor.Col1);
            ClassicAssert.AreEqual(1, cAnchor.Col2);
            ClassicAssert.AreEqual(1, cAnchor.Row1);
            ClassicAssert.AreEqual(1, cAnchor.Row2);

            hssfShape = children[(1)];
            ClassicAssert.IsTrue(hssfShape is HSSFSimpleShape);
            anchor = hssfShape.Anchor as HSSFAnchor;
            ClassicAssert.IsTrue(anchor is HSSFClientAnchor);
            ClassicAssert.AreEqual(512, anchor.Dx1);
            ClassicAssert.AreEqual(1023, anchor.Dx2);
            ClassicAssert.AreEqual(0, anchor.Dy1);
            ClassicAssert.AreEqual(100, anchor.Dy2);
            cAnchor = (HSSFClientAnchor)anchor;
            ClassicAssert.AreEqual(1, cAnchor.Col1);
            ClassicAssert.AreEqual(1, cAnchor.Col2);
            ClassicAssert.AreEqual(1, cAnchor.Row1);
            ClassicAssert.AreEqual(1, cAnchor.Row2);

            hssfShape = children[(2)];
            ClassicAssert.IsTrue(hssfShape is HSSFSimpleShape);
            anchor = hssfShape.Anchor as HSSFAnchor;
            ClassicAssert.IsTrue(anchor is HSSFClientAnchor);
            ClassicAssert.AreEqual(0, anchor.Dx1);
            ClassicAssert.AreEqual(512, anchor.Dx2);
            ClassicAssert.AreEqual(0, anchor.Dy1);
            ClassicAssert.AreEqual(100, anchor.Dy2);
            cAnchor = (HSSFClientAnchor)anchor;
            ClassicAssert.AreEqual(2, cAnchor.Col1);
            ClassicAssert.AreEqual(2, cAnchor.Col2);
            ClassicAssert.AreEqual(2, cAnchor.Row1);
            ClassicAssert.AreEqual(2, cAnchor.Row2);

            hssfShape = children[(3)];
            ClassicAssert.IsTrue(hssfShape is HSSFSimpleShape);
            anchor = hssfShape.Anchor as HSSFAnchor;
            ClassicAssert.IsTrue(anchor is HSSFClientAnchor);
            ClassicAssert.AreEqual(0, anchor.Dx1);
            ClassicAssert.AreEqual(512, anchor.Dx2);
            ClassicAssert.AreEqual(100, anchor.Dy1);
            ClassicAssert.AreEqual(200, anchor.Dy2);
            cAnchor = (HSSFClientAnchor)anchor;
            ClassicAssert.AreEqual(2, cAnchor.Col1);
            ClassicAssert.AreEqual(2, cAnchor.Col2);
            ClassicAssert.AreEqual(2, cAnchor.Row1);
            ClassicAssert.AreEqual(2, cAnchor.Row2);

            wbBack.Close();
        }

    }
}
