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
 * @date 29.06.12
 */
public class TestShapeGroup {

    public void TestSetGetCoordinates(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sh = wb.CreateSheet();
        HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();
        HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
        Assert.AreEqual(group.GetX1(), 0);
        Assert.AreEqual(group.GetY1(), 0);
        Assert.AreEqual(group.GetX2(), 1023);
        Assert.AreEqual(group.GetY2(), 255);

        group.SetCoordinates(1,2,3,4);

        Assert.AreEqual(group.GetX1(), 1);
        Assert.AreEqual(group.GetY1(), 2);
        Assert.AreEqual(group.GetX2(), 3);
        Assert.AreEqual(group.GetY2(), 4);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sh = wb.GetSheetAt(0);
        patriarch = sh.DrawingPatriarch();

        group = (HSSFShapeGroup) patriarch.Children().Get(0);
        Assert.AreEqual(group.GetX1(), 1);
        Assert.AreEqual(group.GetY1(), 2);
        Assert.AreEqual(group.GetX2(), 3);
        Assert.AreEqual(group.GetY2(), 4);
    }

    public void TestAddToExistingFile(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sh = wb.CreateSheet();
        HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();
        HSSFShapeGroup group1 = patriarch.CreateGroup(new HSSFClientAnchor());
        HSSFShapeGroup group2 = patriarch.CreateGroup(new HSSFClientAnchor());

        group1.SetCoordinates(1,2,3,4);
        group2.SetCoordinates(5,6,7,8);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sh = wb.GetSheetAt(0);
        patriarch = sh.DrawingPatriarch();

        Assert.AreEqual(patriarch.Children().Count, 2);

        HSSFShapeGroup group3 = patriarch.CreateGroup(new HSSFClientAnchor());
        group3.SetCoordinates(9,10,11,12);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sh = wb.GetSheetAt(0);
        patriarch = sh.DrawingPatriarch();

        Assert.AreEqual(patriarch.Children().Count, 3);
    }

    public void TestModify() throws Exception {
        HSSFWorkbook wb = new HSSFWorkbook();

        // create a sheet with a text box
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFShapeGroup group1 = patriarch.CreateGroup(new
                HSSFClientAnchor(0,0,0,0,
                (short)0, 0, (short)15, 25));
        group1.SetCoordinates(0, 0, 792, 612);

        HSSFTextbox textbox1 = group1.CreateTextbox(new
                HSSFChildAnchor(100, 100, 300, 300));
        HSSFRichTextString rt1 = new HSSFRichTextString("Hello, World!");
        textbox1.SetString(rt1);

        // Write, read back and check that our text box is there
        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();
        Assert.AreEqual(1, patriarch.Children().Count);

        group1 = (HSSFShapeGroup)patriarch.Children().Get(0);
        Assert.AreEqual(1, group1.Children().Count);
        textbox1 = (HSSFTextbox)group1.Children().Get(0);
        Assert.AreEqual("Hello, World!", textbox1.GetString().GetString());

        // modify anchor
        Assert.AreEqual(new HSSFChildAnchor(100, 100, 300, 300),
                textbox1.GetAnchor());
        HSSFChildAnchor newAnchor = new HSSFChildAnchor(200,200, 400, 400);
        textbox1.SetAnchor(newAnchor);
        // modify text
        textbox1.SetString(new HSSFRichTextString("Hello, World! (modified)"));

        // add a new text box
        HSSFTextbox textbox2 = group1.CreateTextbox(new
                HSSFChildAnchor(400, 400, 600, 600));
        HSSFRichTextString rt2 = new HSSFRichTextString("Hello, World-2");
        textbox2.SetString(rt2);
        Assert.AreEqual(2, group1.Children().Count);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();
        Assert.AreEqual(1, patriarch.Children().Count);

        group1 = (HSSFShapeGroup)patriarch.Children().Get(0);
        Assert.AreEqual(2, group1.Children().Count);
        textbox1 = (HSSFTextbox)group1.Children().Get(0);
        Assert.AreEqual("Hello, World! (modified)",
                textbox1.GetString().GetString());
        Assert.AreEqual(new HSSFChildAnchor(200,200, 400, 400),
                textbox1.GetAnchor());

        textbox2 = (HSSFTextbox)group1.Children().Get(1);
        Assert.AreEqual("Hello, World-2", textbox2.GetString().GetString());
        Assert.AreEqual(new HSSFChildAnchor(400, 400, 600, 600),
                textbox2.GetAnchor());

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();
        group1 = (HSSFShapeGroup)patriarch.Children().Get(0);
        textbox1 = (HSSFTextbox)group1.Children().Get(0);
        textbox2 = (HSSFTextbox)group1.Children().Get(1);
        HSSFTextbox textbox3 = group1.CreateTextbox(new
                HSSFChildAnchor(400,200, 600, 400));
        HSSFRichTextString rt3 = new HSSFRichTextString("Hello, World-3");
        textbox3.SetString(rt3);
    }

    public void TestAddShapesToGroup(){
        HSSFWorkbook wb = new HSSFWorkbook();

        // create a sheet with a text box
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
        int index = wb.AddPicture(new byte[]{1,2,3}, HSSFWorkbook.PICTURE_TYPE_JPEG);
        group.CreatePicture(new HSSFChildAnchor(), index);
        HSSFPolygon polygon = group.CreatePolygon(new HSSFChildAnchor());
        polygon.SetPoints(new int[]{1,100, 1}, new int[]{1, 50, 100});
        group.CreateTextbox(new HSSFChildAnchor());
        group.CreateShape(new HSSFChildAnchor());

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();
        Assert.AreEqual(1, patriarch.Children().Count);

        Assert.IsTrue(patriarch.Children().Get(0) is HSSFShapeGroup);
        group = (HSSFShapeGroup) patriarch.Children().Get(0);

        Assert.AreEqual(group.Children().Count, 4);

        Assert.IsTrue(group.Children().Get(0) is HSSFPicture);
        Assert.IsTrue(group.Children().Get(1) is HSSFPolygon);
        Assert.IsTrue(group.Children().Get(2) is HSSFTextbox);
        Assert.IsTrue(group.Children().Get(3) is HSSFSimpleShape);

        HSSFShapeGroup group2 = patriarch.CreateGroup(new HSSFClientAnchor());

        index = wb.AddPicture(new byte[]{2,2,2}, HSSFWorkbook.PICTURE_TYPE_JPEG);
        group2.CreatePicture(new HSSFChildAnchor(), index);
        polygon = group2.CreatePolygon(new HSSFChildAnchor());
        polygon.SetPoints(new int[]{1,100, 1}, new int[]{1, 50, 100});
        group2.CreateTextbox(new HSSFChildAnchor());
        group2.CreateShape(new HSSFChildAnchor());
        group2.CreateShape(new HSSFChildAnchor());

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();
        Assert.AreEqual(2, patriarch.Children().Count);

        group = (HSSFShapeGroup) patriarch.Children().Get(1);

        Assert.AreEqual(group.Children().Count, 5);

        Assert.IsTrue(group.Children().Get(0) is HSSFPicture);
        Assert.IsTrue(group.Children().Get(1) is HSSFPolygon);
        Assert.IsTrue(group.Children().Get(2) is HSSFTextbox);
        Assert.IsTrue(group.Children().Get(3) is HSSFSimpleShape);
        Assert.IsTrue(group.Children().Get(4) is HSSFSimpleShape);

        group.GetShapeId();
    }

    public void TestSpgrRecord(){
        HSSFWorkbook wb = new HSSFWorkbook();

        // create a sheet with a text box
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());
        Assert.AreSame(((EscherContainerRecord)group.GetEscherContainer().GetChild(0)).GetChildById(EscherSpgrRecord.RECORD_ID), GetSpgrRecord(group));
    }

    private static EscherSpgrRecord GetSpgrRecord(HSSFShapeGroup group) {
        Field spgrField = null;
        try {
            spgrField = group.GetType().GetDeclaredField("_spgrRecord");
            spgrField.SetAccessible(true);
            return (EscherSpgrRecord) spgrField.Get(group);
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return null;
    }

    public void TestClearShapes(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet();
        HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();
        HSSFShapeGroup group = patriarch.CreateGroup(new HSSFClientAnchor());

        group.CreateShape(new HSSFChildAnchor());
        group.CreateShape(new HSSFChildAnchor());

        EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(patriarch);

        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 5);
        Assert.AreEqual(agg.GetTailRecords().Count, 0);
        Assert.AreEqual(group.Children().Count, 2);

        group.Clear();

        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 1);
        Assert.AreEqual(agg.GetTailRecords().Count, 0);
        Assert.AreEqual(group.Children().Count, 0);

        wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        sheet = wb.GetSheetAt(0);
        patriarch = sheet.DrawingPatriarch();

        group = (HSSFShapeGroup) patriarch.Children().Get(0);

        Assert.AreEqual(agg.GetShapeToObjMapping().Count, 1);
        Assert.AreEqual(agg.GetTailRecords().Count, 0);
        Assert.AreEqual(group.Children().Count, 0);
    }
}


}
