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

namespace TestCases.HSSF.Model
{/**
 * @author Evgeniy Berlog
 * date: 12.06.12
 */
public class TestDrawingShapes  {

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
    public void TestDrawingGroups() {
        HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
        HSSFSheet sheet = wb.GetSheet("groups");
        HSSFPatriarch patriarch = sheet.DrawingPatriarch();
        Assert.AreEqual(patriarch.Children().Count, 2);
        HSSFShapeGroup group = (HSSFShapeGroup) patriarch.Children().Get(1);
        Assert.AreEqual(3, group.Children().Count);
        HSSFShapeGroup group1 = (HSSFShapeGroup) group.Children().Get(0);
        Assert.AreEqual(2, group1.Children().Count);
        group1 = (HSSFShapeGroup) group.Children().Get(2);
        Assert.AreEqual(2, group1.Children().Count);
    }

    public void TestHSSFShapeCompatibility() {
        HSSFSimpleShape shape = new HSSFSimpleShape(null, new HSSFClientAnchor());
        shape.SetShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
        Assert.AreEqual(0x08000040, shape.GetLineStyleColor());
        Assert.AreEqual(0x08000009, shape.GetFillColor());
        Assert.AreEqual(HSSFShape.LINEWIDTH_DEFAULT, shape.GetLineWidth());
        Assert.AreEqual(HSSFShape.LINESTYLE_SOLID, shape.GetLineStyle());
        Assert.IsFalse(shape.IsNoFill());

        AbstractShape sp = AbstractShape.CreateShape(shape, 1);
        EscherContainerRecord spContainer = sp.GetSpContainer();
        EscherOptRecord opt =
                spContainer.GetChildById(EscherOptRecord.RECORD_ID);

        Assert.AreEqual(7, opt.GetEscherProperties().Count);
        Assert.AreEqual(true,
                ((EscherBoolProperty) opt.Lookup(EscherProperties.TEXT__SIZE_TEXT_TO_FIT_SHAPE)).IsTrue());
        Assert.AreEqual(0x00000004,
                ((EscherSimpleProperty) opt.Lookup(EscherProperties.GEOMETRY__SHAPEPATH)).GetPropertyValue());
        Assert.AreEqual(0x08000009,
                ((EscherSimpleProperty) opt.Lookup(EscherProperties.FILL__FILLCOLOR)).GetPropertyValue());
        Assert.AreEqual(true,
                ((EscherBoolProperty) opt.Lookup(EscherProperties.FILL__NOFILLHITTEST)).IsTrue());
        Assert.AreEqual(0x08000040,
                ((EscherSimpleProperty) opt.Lookup(EscherProperties.LINESTYLE__COLOR)).GetPropertyValue());
        Assert.AreEqual(true,
                ((EscherBoolProperty) opt.Lookup(EscherProperties.LINESTYLE__NOLINEDRAWDASH)).IsTrue());
        Assert.AreEqual(true,
                ((EscherBoolProperty) opt.Lookup(EscherProperties.GROUPSHAPE__PRINT)).IsTrue());
    }

    public void TestDefaultPictureSettings() {
        HSSFPicture picture = new HSSFPicture(null, new HSSFClientAnchor());
        Assert.AreEqual(picture.GetLineWidth(), HSSFShape.LINEWIDTH_DEFAULT);
        Assert.AreEqual(picture.GetFillColor(), HSSFShape.FILL__FILLCOLOR_DEFAULT);
        Assert.AreEqual(picture.GetLineStyle(), HSSFShape.LINESTYLE_NONE);
        Assert.AreEqual(picture.GetLineStyleColor(), HSSFShape.LINESTYLE__COLOR_DEFAULT);
        Assert.AreEqual(picture.IsNoFill(), false);
        Assert.AreEqual(picture.GetPictureIndex(), -1);//not Set yet
    }

    /**
     * No NullPointerException should appear
     */
    public void TestDefaultSettingsWithEmptyContainer() {
        EscherContainerRecord Container = new EscherContainerRecord();
        EscherOptRecord opt = new EscherOptRecord();
        opt.SetRecordId(EscherOptRecord.RECORD_ID);
        Container.addChildRecord(opt);
        ObjRecord obj = new ObjRecord();
        CommonObjectDataSubRecord cod = new CommonObjectDataSubRecord();
        cod.SetObjectType(HSSFSimpleShape.OBJECT_TYPE_PICTURE);
        obj.AddSubRecord(cod);
        HSSFPicture picture = new HSSFPicture(Container, obj);

        Assert.AreEqual(picture.GetLineWidth(), HSSFShape.LINEWIDTH_DEFAULT);
        Assert.AreEqual(picture.GetFillColor(), HSSFShape.FILL__FILLCOLOR_DEFAULT);
        Assert.AreEqual(picture.GetLineStyle(), HSSFShape.LINESTYLE_DEFAULT);
        Assert.AreEqual(picture.GetLineStyleColor(), HSSFShape.LINESTYLE__COLOR_DEFAULT);
        Assert.AreEqual(picture.IsNoFill(), HSSFShape.NO_FILL_DEFAULT);
        Assert.AreEqual(picture.GetPictureIndex(), -1);//not Set yet
    }

    /**
     * create a rectangle, save the workbook, read back and verify that all shape properties are there
     */
    public void TestReadWriteRectangle()  {

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();

        HSSFPatriarch Drawing = sheet.CreateDrawingPatriarch();
        HSSFClientAnchor anchor = new HSSFClientAnchor(10, 10, 50, 50, (short) 2, 2, (short) 4, 4);
        anchor.SetAnchorType(2);
        Assert.AreEqual(anchor.GetAnchorType(), 2);

        HSSFSimpleShape rectangle = Drawing.CreateSimpleShape(anchor);
        rectangle.SetShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
        rectangle.SetLineWidth(10000);
        rectangle.SetFillColor(777);
        Assert.AreEqual(rectangle.GetFillColor(), 777);
        Assert.AreEqual(10000, rectangle.GetLineWidth());
        rectangle.SetLineStyle(10);
        Assert.AreEqual(10, rectangle.GetLineStyle());
        Assert.AreEqual(rectangle.GetWrapText(), HSSFSimpleShape.WRAP_SQUARE);
        rectangle.SetLineStyleColor(1111);
        rectangle.SetNoFill(true);
        rectangle.SetWrapText(HSSFSimpleShape.WRAP_NONE);
        rectangle.SetString(new HSSFRichTextString("teeeest"));
        Assert.AreEqual(rectangle.GetLineStyleColor(), 1111);
        Assert.AreEqual(((EscherSimpleProperty)((EscherOptRecord)HSSFTestHelper.GetEscherContainer(rectangle).GetChildById(EscherOptRecord.RECORD_ID))
                .Lookup(EscherProperties.TEXT__TEXTID)).GetPropertyValue(), "teeeest".HashCode());
        Assert.AreEqual(rectangle.IsNoFill(), true);
        Assert.AreEqual(rectangle.GetWrapText(), HSSFSimpleShape.WRAP_NONE);
        Assert.AreEqual(rectangle.GetString().GetString(), "teeeest");

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        Drawing = sheet.DrawingPatriarch();
        Assert.AreEqual(1, Drawing.Children().Count);

        HSSFSimpleShape rectangle2 =
                (HSSFSimpleShape) Drawing.Children().Get(0);
        Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE,
                rectangle2.GetShapeType());
        Assert.AreEqual(10000, rectangle2.GetLineWidth());
        Assert.AreEqual(10, rectangle2.GetLineStyle());
        Assert.AreEqual(anchor, rectangle2.GetAnchor());
        Assert.AreEqual(rectangle2.GetLineStyleColor(), 1111);
        Assert.AreEqual(rectangle2.GetFillColor(), 777);
        Assert.AreEqual(rectangle2.IsNoFill(), true);
        Assert.AreEqual(rectangle2.GetString().GetString(), "teeeest");
        Assert.AreEqual(rectangle.GetWrapText(), HSSFSimpleShape.WRAP_NONE);

        rectangle2.SetFillColor(3333);
        rectangle2.SetLineStyle(9);
        rectangle2.SetLineStyleColor(4444);
        rectangle2.SetNoFill(false);
        rectangle2.SetLineWidth(77);
        rectangle2.GetAnchor().SetDx1(2);
        rectangle2.GetAnchor().SetDx2(3);
        rectangle2.GetAnchor().SetDy1(4);
        rectangle2.GetAnchor().SetDy2(5);
        rectangle.SetWrapText(HSSFSimpleShape.WRAP_BY_POINTS);
        rectangle2.SetString(new HSSFRichTextString("test22"));

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        Drawing = sheet.DrawingPatriarch();
        Assert.AreEqual(1, Drawing.Children().Count);
        rectangle2 = (HSSFSimpleShape) Drawing.Children().Get(0);
        Assert.AreEqual(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE, rectangle2.GetShapeType());
        Assert.AreEqual(rectangle.GetWrapText(), HSSFSimpleShape.WRAP_BY_POINTS);
        Assert.AreEqual(77, rectangle2.GetLineWidth());
        Assert.AreEqual(9, rectangle2.GetLineStyle());
        Assert.AreEqual(rectangle2.GetLineStyleColor(), 4444);
        Assert.AreEqual(rectangle2.GetFillColor(), 3333);
        Assert.AreEqual(rectangle2.GetAnchor().GetDx1(), 2);
        Assert.AreEqual(rectangle2.GetAnchor().GetDx2(), 3);
        Assert.AreEqual(rectangle2.GetAnchor().GetDy1(), 4);
        Assert.AreEqual(rectangle2.GetAnchor().GetDy2(), 5);
        Assert.AreEqual(rectangle2.IsNoFill(), false);
        Assert.AreEqual(rectangle2.GetString().GetString(), "test22");

        HSSFSimpleShape rect3 = Drawing.CreateSimpleShape(new HSSFClientAnchor());
        rect3.SetShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

        Drawing = wb.GetSheetAt(0).DrawingPatriarch();
        Assert.AreEqual(drawing.Children().Count, 2);
    }

    public void TestReadExistingImage() {
        HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
        HSSFSheet sheet = wb.GetSheet("pictures");
        HSSFPatriarch Drawing = sheet.DrawingPatriarch();
        Assert.AreEqual(1, Drawing.Children().Count);
        HSSFPicture picture = (HSSFPicture) Drawing.Children().Get(0);

        Assert.AreEqual(picture.GetPictureIndex(), 2);
        Assert.AreEqual(picture.GetLineStyleColor(), HSSFShape.LINESTYLE__COLOR_DEFAULT);
        Assert.AreEqual(picture.GetFillColor(), 0x5DC943);
        Assert.AreEqual(picture.GetLineWidth(), HSSFShape.LINEWIDTH_DEFAULT);
        Assert.AreEqual(picture.GetLineStyle(), HSSFShape.LINESTYLE_DEFAULT);
        Assert.AreEqual(picture.IsNoFill(), false);

        picture.SetPictureIndex(2);
        Assert.AreEqual(picture.GetPictureIndex(), 2);
    }


    /* assert shape properties when Reading shapes from a existing workbook */
    public void TestReadExistingRectangle() {
        HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
        HSSFSheet sheet = wb.GetSheet("rectangles");
        HSSFPatriarch Drawing = sheet.DrawingPatriarch();
        Assert.AreEqual(1, Drawing.Children().Count);

        HSSFSimpleShape shape = (HSSFSimpleShape) Drawing.Children().Get(0);
        Assert.AreEqual(shape.IsNoFill(), false);
        Assert.AreEqual(shape.GetLineStyle(), HSSFShape.LINESTYLE_DASHDOTGEL);
        Assert.AreEqual(shape.GetLineStyleColor(), 0x616161);
        Assert.AreEqual(HexDump.ToHex(shape.GetFillColor()), shape.GetFillColor(), 0x2CE03D);
        Assert.AreEqual(shape.GetLineWidth(), HSSFShape.LINEWIDTH_ONE_PT * 2);
        Assert.AreEqual(shape.GetString().GetString(), "POItest");
        Assert.AreEqual(shape.GetRotationDegree(), 27);
    }

    public void TestShapeIds() {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet1 = wb.CreateSheet();
        HSSFPatriarch patriarch1 = sheet1.CreateDrawingPatriarch();
        for (int i = 0; i < 2; i++) {
            patriarch1.CreateSimpleShape(new HSSFClientAnchor());
        }

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet1 = wb.GetSheetAt(0);
        patriarch1 = sheet1.DrawingPatriarch();

        EscherAggregate agg1 = HSSFTestHelper.GetEscherAggregate(patriarch1);
        // last shape ID cached in EscherDgRecord
        EscherDgRecord dg1 =
                agg1.GetEscherContainer().GetChildById(EscherDgRecord.RECORD_ID);
        Assert.AreEqual(1026, dg1.GetLastMSOSPID());

        // iterate over shapes and check shapeId
        EscherContainerRecord spgrContainer =
                agg1.GetEscherContainer().GetChildContainers().Get(0);
        // root spContainer + 2 spContainers for shapes
        Assert.AreEqual(3, spgrContainer.GetChildRecords().Count);

        EscherSpRecord sp0 =
                ((EscherContainerRecord) spgrContainer.GetChild(0)).GetChildById(EscherSpRecord.RECORD_ID);
        Assert.AreEqual(1024, sp0.GetShapeId());

        EscherSpRecord sp1 =
                ((EscherContainerRecord) spgrContainer.GetChild(1)).GetChildById(EscherSpRecord.RECORD_ID);
        Assert.AreEqual(1025, sp1.GetShapeId());

        EscherSpRecord sp2 =
                ((EscherContainerRecord) spgrContainer.GetChild(2)).GetChildById(EscherSpRecord.RECORD_ID);
        Assert.AreEqual(1026, sp2.GetShapeId());
    }

    /**
     * Test Get new id for shapes from existing file
     * File already have for 1 shape on each sheet, because document must contain EscherDgRecord for each sheet
     */
    public void TestAllocateNewIds() {
        HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("empty.xls");
        HSSFSheet sheet = wb.GetSheetAt(0);
        HSSFPatriarch patriarch = sheet.DrawingPatriarch();

        /**
         * 2048 - main SpContainer id
         * 2049 - existing shape id
         */
        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 2050);
        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 2051);
        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 2052);

        sheet = wb.GetSheetAt(1);
        patriarch = sheet.DrawingPatriarch();

        /**
         * 3072 - main SpContainer id
         * 3073 - existing shape id
         */
        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 3074);
        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 3075);
        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 3076);


        sheet = wb.GetSheetAt(2);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 1026);
        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 1027);
        Assert.AreEqual(HSSFTestHelper.allocateNewShapeId(patriarch), 1028);
    }

    public void TestOpt() throws Exception {
        HSSFWorkbook wb = new HSSFWorkbook();

        // create a sheet with a text box
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFTextbox textbox = patriarch.CreateTextbox(new HSSFClientAnchor());
        EscherOptRecord opt1 = HSSFTestHelper.GetOptRecord(textbox);
        EscherOptRecord opt2 = HSSFTestHelper.GetEscherContainer(textbox).GetChildById(EscherOptRecord.RECORD_ID);
        Assert.AreSame(opt1, opt2);
    }
    
    public void TestCorrectOrderInOptRecord(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFTextbox textbox = patriarch.CreateTextbox(new HSSFClientAnchor());
        EscherOptRecord opt = HSSFTestHelper.GetOptRecord(textbox);    
        
        String opt1Str = opt.ToXml();

        textbox.SetFillColor(textbox.GetFillColor());
        EscherContainerRecord Container = HSSFTestHelper.GetEscherContainer(textbox);
        EscherOptRecord optRecord = Container.GetChildById(EscherOptRecord.RECORD_ID);
        Assert.AreEqual(opt1Str, optRecord.ToXml());
        textbox.SetLineStyle(textbox.GetLineStyle());
        Assert.AreEqual(opt1Str, optRecord.ToXml());
        textbox.SetLineWidth(textbox.GetLineWidth());
        Assert.AreEqual(opt1Str, optRecord.ToXml());
        textbox.SetLineStyleColor(textbox.GetLineStyleColor());
        Assert.AreEqual(opt1Str, optRecord.ToXml());
    }

    public void TestDgRecordNumShapes(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        EscherAggregate aggregate = HSSFTestHelper.GetEscherAggregate(patriarch);
        EscherDgRecord dgRecord = (EscherDgRecord) aggregate.GetEscherRecord(0).GetChild(0);
        Assert.AreEqual(dgRecord.GetNumShapes(), 1);
    }

    public void TestTextForSimpleShape(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFSimpleShape shape = patriarch.CreateSimpleShape(new HSSFClientAnchor());
        shape.SetShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);

        EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);
        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        shape = (HSSFSimpleShape) patriarch.Children().Get(0);

        agg = HSSFTestHelper.GetEscherAggregate(patriarch);
        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

        shape.SetString(new HSSFRichTextString("string1"));
        Assert.AreEqual(shape.GetString().GetString(), "string1");

        Assert.IsNotNull(HSSFTestHelper.GetEscherContainer(shape).GetChildById(EscherTextboxRecord.RECORD_ID));
        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 2);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        shape = (HSSFSimpleShape) patriarch.Children().Get(0);

        Assert.IsNotNull(HSSFTestHelper.GetTextObjRecord(shape));
        Assert.AreEqual(shape.GetString().GetString(), "string1");
        Assert.IsNotNull(HSSFTestHelper.GetEscherContainer(shape).GetChildById(EscherTextboxRecord.RECORD_ID));
        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 2);
    }

    public void TestRemoveShapes(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor());
        rectangle.SetShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);

        int idx = wb.AddPicture(new byte[]{1,2,3}, Workbook.PICTURE_TYPE_JPEG);
        patriarch.CreatePicture(new HSSFClientAnchor(), idx);

        patriarch.CreateCellComment(new HSSFClientAnchor());

        HSSFPolygon polygon = patriarch.CreatePolygon(new HSSFClientAnchor());
        polygon.SetPoints(new int[]{1,2}, new int[]{2,3});

        patriarch.CreateTextbox(new HSSFClientAnchor());

        HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
        group.CreateTextbox(new HSSFChildAnchor());
        group.CreatePicture(new HSSFChildAnchor(), idx);

        Assert.AreEqual(patriarch.Children().Count, 6);
        Assert.AreEqual(group.Children().Count, 2);

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 12);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 12);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);

        Assert.AreEqual(patriarch.Children().Count, 6);

        group = (HSSFShapeGroup) patriarch.Children().Get(5);
        group.RemoveShape(group.Children().Get(0));

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 10);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 10);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);

        group = (HSSFShapeGroup) patriarch.Children().Get(5);
        patriarch.RemoveShape(group);

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 8);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 8);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);
        Assert.AreEqual(patriarch.Children().Count, 5);

        HSSFShape shape = patriarch.Children().Get(0);
        patriarch.RemoveShape(shape);

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 6);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);
        Assert.AreEqual(patriarch.Children().Count, 4);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 6);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);
        Assert.AreEqual(patriarch.Children().Count, 4);

        HSSFPicture picture = (HSSFPicture) patriarch.Children().Get(0);
        patriarch.RemoveShape(picture);

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 5);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);
        Assert.AreEqual(patriarch.Children().Count, 3);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 5);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 1);
        Assert.AreEqual(patriarch.Children().Count, 3);

        HSSFComment comment = (HSSFComment) patriarch.Children().Get(0);
        patriarch.RemoveShape(comment);

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 3);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 0);
        Assert.AreEqual(patriarch.Children().Count, 2);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 3);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 0);
        Assert.AreEqual(patriarch.Children().Count, 2);

        polygon = (HSSFPolygon) patriarch.Children().Get(0);
        patriarch.RemoveShape(polygon);

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 2);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 0);
        Assert.AreEqual(patriarch.Children().Count, 1);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 2);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 0);
        Assert.AreEqual(patriarch.Children().Count, 1);

        HSSFTextbox textbox = (HSSFTextbox) patriarch.Children().Get(0);
        patriarch.RemoveShape(textbox);

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 0);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 0);
        Assert.AreEqual(patriarch.Children().Count, 0);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetShapeToObjMapping().Count, 0);
        Assert.AreEqual(HSSFTestHelper.GetEscherAggregate(patriarch).GetTailRecords().Count, 0);
        Assert.AreEqual(patriarch.Children().Count, 0);
    }

    public void TestShapeFlip(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor());
        rectangle.SetShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);

        Assert.AreEqual(rectangle.IsFlipVertical(), false);
        Assert.AreEqual(rectangle.IsFlipHorizontal(), false);

        rectangle.SetFlipVertical(true);
        Assert.AreEqual(rectangle.IsFlipVertical(), true);
        rectangle.SetFlipHorizontal(true);
        Assert.AreEqual(rectangle.IsFlipHorizontal(), true);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        rectangle = (HSSFSimpleShape) patriarch.Children().Get(0);

        Assert.AreEqual(rectangle.IsFlipHorizontal(), true);
        rectangle.SetFlipHorizontal(false);
        Assert.AreEqual(rectangle.IsFlipHorizontal(), false);

        Assert.AreEqual(rectangle.IsFlipVertical(), true);
        rectangle.SetFlipVertical(false);
        Assert.AreEqual(rectangle.IsFlipVertical(), false);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        rectangle = (HSSFSimpleShape) patriarch.Children().Get(0);

        Assert.AreEqual(rectangle.IsFlipVertical(), false);
        Assert.AreEqual(rectangle.IsFlipHorizontal(), false);
    }

    public void TestRotation() {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFSimpleShape rectangle = patriarch.CreateSimpleShape(new HSSFClientAnchor(0,0,100,100, (short) 0,0,(short)5,5));
        rectangle.SetShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);

        Assert.AreEqual(rectangle.GetRotationDegree(), 0);
        rectangle.SetRotationDegree((short) 45);
        Assert.AreEqual(rectangle.GetRotationDegree(), 45);
        rectangle.SetFlipHorizontal(true);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();
        rectangle = (HSSFSimpleShape) patriarch.Children().Get(0);
        Assert.AreEqual(rectangle.GetRotationDegree(), 45);
        rectangle.SetRotationDegree((short) 30);
        Assert.AreEqual(rectangle.GetRotationDegree(), 30);

        patriarch.SetCoordinates(0, 0, 10, 10);
        rectangle.SetString(new HSSFRichTextString("1234"));
    }

    public void TestShapeContainerImplementsIterable(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        patriarch.CreateSimpleShape(new HSSFClientAnchor());
        patriarch.CreateSimpleShape(new HSSFClientAnchor());

        int i=2;

        foreach (HSSFShape shape: patriarch){
            i--;
        }
        Assert.AreEqual(i, 0);
    }

    public void TestClearShapesForPatriarch(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        patriarch.CreateSimpleShape(new HSSFClientAnchor());
        patriarch.CreateSimpleShape(new HSSFClientAnchor());
        patriarch.CreateCellComment(new HSSFClientAnchor());

        EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);

        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 6);
        Assert.AreEqual(agg.GetTailRecords().Count, 1);
        Assert.AreEqual(patriarch.Children().Count, 3);

        patriarch.Clear();

        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 0);
        Assert.AreEqual(agg.GetTailRecords().Count, 0);
        Assert.AreEqual(patriarch.Children().Count, 0);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 0);
        Assert.AreEqual(agg.GetTailRecords().Count, 0);
        Assert.AreEqual(patriarch.Children().Count, 0);
    }
}

}
