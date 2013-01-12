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
 * @date 25.06.12
 */
    public class TestText
    {

        public void TestResultEqualsToAbstractShape()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();
            HSSFTextbox textbox = patriarch.CreateTextbox(new HSSFClientAnchor());
            TextboxShape textboxShape = HSSFTestModelHelper.CreateTextboxShape(1025, textbox);

            Assert.AreEqual(textbox.GetEscherContainer().GetChildRecords().Count, 5);
            Assert.AreEqual(textboxShape.GetSpContainer().GetChildRecords().Count, 5);

            //sp record
            byte[] expected = textboxShape.GetSpContainer().GetChild(0).Serialize();
            byte[] actual = textbox.GetEscherContainer().GetChild(0).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = textboxShape.GetSpContainer().GetChild(2).Serialize();
            actual = textbox.GetEscherContainer().GetChild(2).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = textboxShape.GetSpContainer().GetChild(3).Serialize();
            actual = textbox.GetEscherContainer().GetChild(3).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            expected = textboxShape.GetSpContainer().GetChild(4).Serialize();
            actual = textbox.GetEscherContainer().GetChild(4).Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));

            ObjRecord obj = textbox.GetObjRecord();
            ObjRecord objShape = textboxShape.GetObjRecord();

            expected = obj.Serialize();
            actual = objShape.Serialize();

            TextObjectRecord tor = textbox.GetTextObjectRecord();
            TextObjectRecord torShape = textboxShape.GetTextObjectRecord();

            expected = tor.Serialize();
            actual = torShape.Serialize();

            Assert.AreEqual(expected.Length, actual.Length);
            Assert.IsTrue(Arrays.Equals(expected, actual));
        }

        public void TestAddTextToExistingFile()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();
            HSSFTextbox textbox = patriarch.CreateTextbox(new HSSFClientAnchor());
            textbox.SetString(new HSSFRichTextString("just for Test"));
            HSSFTextbox textbox2 = patriarch.CreateTextbox(new HSSFClientAnchor());
            textbox2.SetString(new HSSFRichTextString("just for Test2"));

            Assert.AreEqual(patriarch.Children().Count, 2);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            Assert.AreEqual(patriarch.Children().Count, 2);
            HSSFTextbox text3 = patriarch.CreateTextbox(new HSSFClientAnchor());
            text3.SetString(new HSSFRichTextString("text3"));
            Assert.AreEqual(patriarch.Children().Count, 3);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();

            Assert.AreEqual(patriarch.Children().Count, 3);
            Assert.AreEqual(((HSSFTextbox)patriarch.Children().Get(0)).GetString().GetString(), "just for Test");
            Assert.AreEqual(((HSSFTextbox)patriarch.Children().Get(1)).GetString().GetString(), "just for Test2");
            Assert.AreEqual(((HSSFTextbox)patriarch.Children().Get(2)).GetString().GetString(), "text3");
        }

        public void TestSetGetProperties()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sh = wb.CreateSheet();
            HSSFPatriarch patriarch = sh.CreateDrawingPatriarch();
            HSSFTextbox textbox = patriarch.CreateTextbox(new HSSFClientAnchor());
            textbox.SetString(new HSSFRichTextString("test"));
            Assert.AreEqual(textbox.GetString().GetString(), "test");

            textbox.SetHorizontalAlignment((short)5);
            Assert.AreEqual(textbox.GetHorizontalAlignment(), 5);

            textbox.SetVerticalAlignment((short)6);
            Assert.AreEqual(textbox.GetVerticalAlignment(), (short)6);

            textbox.SetMarginBottom(7);
            Assert.AreEqual(textbox.GetMarginBottom(), 7);

            textbox.SetMarginLeft(8);
            Assert.AreEqual(textbox.GetMarginLeft(), 8);

            textbox.SetMarginRight(9);
            Assert.AreEqual(textbox.GetMarginRight(), 9);

            textbox.SetMarginTop(10);
            Assert.AreEqual(textbox.GetMarginTop(), 10);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();
            textbox = (HSSFTextbox)patriarch.Children().Get(0);
            Assert.AreEqual(textbox.GetString().GetString(), "test");
            Assert.AreEqual(textbox.GetHorizontalAlignment(), 5);
            Assert.AreEqual(textbox.GetVerticalAlignment(), (short)6);
            Assert.AreEqual(textbox.GetMarginBottom(), 7);
            Assert.AreEqual(textbox.GetMarginLeft(), 8);
            Assert.AreEqual(textbox.GetMarginRight(), 9);
            Assert.AreEqual(textbox.GetMarginTop(), 10);

            textbox.SetString(new HSSFRichTextString("test1"));
            textbox.SetHorizontalAlignment(HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
            textbox.SetVerticalAlignment(HSSFTextbox.VERTICAL_ALIGNMENT_TOP);
            textbox.SetMarginBottom(71);
            textbox.SetMarginLeft(81);
            textbox.SetMarginRight(91);
            textbox.SetMarginTop(101);

            Assert.AreEqual(textbox.GetString().GetString(), "test1");
            Assert.AreEqual(textbox.GetHorizontalAlignment(), HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
            Assert.AreEqual(textbox.GetVerticalAlignment(), HSSFTextbox.VERTICAL_ALIGNMENT_TOP);
            Assert.AreEqual(textbox.GetMarginBottom(), 71);
            Assert.AreEqual(textbox.GetMarginLeft(), 81);
            Assert.AreEqual(textbox.GetMarginRight(), 91);
            Assert.AreEqual(textbox.GetMarginTop(), 101);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0);
            patriarch = sh.DrawingPatriarch();
            textbox = (HSSFTextbox)patriarch.Children().Get(0);

            Assert.AreEqual(textbox.GetString().GetString(), "test1");
            Assert.AreEqual(textbox.GetHorizontalAlignment(), HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
            Assert.AreEqual(textbox.GetVerticalAlignment(), HSSFTextbox.VERTICAL_ALIGNMENT_TOP);
            Assert.AreEqual(textbox.GetMarginBottom(), 71);
            Assert.AreEqual(textbox.GetMarginLeft(), 81);
            Assert.AreEqual(textbox.GetMarginRight(), 91);
            Assert.AreEqual(textbox.GetMarginTop(), 101);
        }

        public void TestExistingFileWithText()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("text");
            HSSFPatriarch Drawing = sheet.DrawingPatriarch();
            Assert.AreEqual(1, Drawing.Children().Count);
            HSSFTextbox textbox = (HSSFTextbox)Drawing.Children().Get(0);
            Assert.AreEqual(textbox.GetHorizontalAlignment(), HSSFTextbox.HORIZONTAL_ALIGNMENT_LEFT);
            Assert.AreEqual(textbox.GetVerticalAlignment(), HSSFTextbox.VERTICAL_ALIGNMENT_TOP);
            Assert.AreEqual(textbox.GetMarginTop(), 0);
            Assert.AreEqual(textbox.GetMarginBottom(), 3600000);
            Assert.AreEqual(textbox.GetMarginLeft(), 3600000);
            Assert.AreEqual(textbox.GetMarginRight(), 0);
            Assert.AreEqual(textbox.GetString().GetString(), "teeeeesssstttt");
        }
    }

}
