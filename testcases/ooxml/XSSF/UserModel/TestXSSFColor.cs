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

namespace TestCases.XSSF.UserModel
{
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    [TestFixture]
    public class TestXSSFColor
    {
        [Test]
        public void TestIndexedColour()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48779.xlsx");

            // Check the CTColor is as expected
            XSSFColor indexed = (XSSFColor)wb.GetCellStyleAt(1).FillBackgroundColorColor;
            ClassicAssert.AreEqual(true, indexed.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(64, indexed.GetCTColor().indexed);
            ClassicAssert.AreEqual(false, indexed.GetCTColor().IsSetRgb());
            ClassicAssert.AreEqual(null, indexed.GetCTColor().GetRgb());

            // Now check the XSSFColor
            // Note - 64 is a special "auto" one with no rgb equiv
            ClassicAssert.AreEqual(64, indexed.Indexed);
            ClassicAssert.AreEqual(null, indexed.RGB);
            ClassicAssert.AreEqual(null, indexed.GetRgbWithTint());
            ClassicAssert.AreEqual(null, indexed.ARGBHex);
            ClassicAssert.IsFalse(indexed.HasAlpha);
            ClassicAssert.IsFalse(indexed.HasTint);

            // Now Move to one with indexed rgb values
            indexed.Indexed = (59);
            ClassicAssert.AreEqual(true, indexed.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(59, indexed.GetCTColor().indexed);
            ClassicAssert.AreEqual(false, indexed.GetCTColor().IsSetRgb());
            ClassicAssert.AreEqual(null, indexed.GetCTColor().GetRgb());

            ClassicAssert.AreEqual(59, indexed.Indexed);
            ClassicAssert.AreEqual("FF333300", indexed.ARGBHex);

            ClassicAssert.AreEqual(3, indexed.RGB.Length);
            ClassicAssert.AreEqual(0x33, indexed.RGB[0]);
            ClassicAssert.AreEqual(0x33, indexed.RGB[1]);
            ClassicAssert.AreEqual(0x00, indexed.RGB[2]);

            ClassicAssert.AreEqual(4, indexed.ARGB.Length);
            ClassicAssert.AreEqual(255, indexed.ARGB[0]);
            ClassicAssert.AreEqual(0x33, indexed.ARGB[1]);
            ClassicAssert.AreEqual(0x33, indexed.ARGB[2]);
            ClassicAssert.AreEqual(0x00, indexed.ARGB[3]);

            // You don't Get tinted indexed colours, sorry...
            ClassicAssert.AreEqual(null, indexed.GetRgbWithTint());
        }
        [Test]
        public void TestRGBColour()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50299.xlsx");

            // Check the CTColor is as expected
            XSSFColor rgb3 = (XSSFColor)((XSSFCellStyle)wb.GetCellStyleAt((short)25)).FillForegroundXSSFColor;
            ClassicAssert.AreEqual(false, rgb3.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(0, rgb3.GetCTColor().indexed);
            ClassicAssert.AreEqual(true, rgb3.GetCTColor().IsSetTint());
            ClassicAssert.AreEqual(-0.34999, rgb3.GetCTColor().tint, 0.00001);
            ClassicAssert.AreEqual(true, rgb3.GetCTColor().IsSetRgb());
            ClassicAssert.AreEqual(3, rgb3.GetCTColor().GetRgb().Length);

            // Now check the XSSFColor
            ClassicAssert.AreEqual(0, rgb3.Indexed);
            ClassicAssert.AreEqual(-0.34999, rgb3.Tint, 0.00001);
            ClassicAssert.IsFalse(rgb3.HasAlpha);
            ClassicAssert.IsTrue(rgb3.HasTint);

            ClassicAssert.AreEqual("FFFFFFFF", rgb3.ARGBHex);
            ClassicAssert.AreEqual(3, rgb3.RGB.Length);
            ClassicAssert.AreEqual(255, rgb3.RGB[0]);
            ClassicAssert.AreEqual(255, rgb3.RGB[1]);
            ClassicAssert.AreEqual(255, rgb3.RGB[2]);

            ClassicAssert.AreEqual(4, rgb3.ARGB.Length);
            ClassicAssert.AreEqual(255, rgb3.ARGB[0]);
            ClassicAssert.AreEqual(255, rgb3.ARGB[1]);
            ClassicAssert.AreEqual(255, rgb3.ARGB[2]);
            ClassicAssert.AreEqual(255, rgb3.ARGB[3]);

            // Tint doesn't have the alpha
            // tint = -0.34999
            // 255 * (1 + tint) = 165 truncated
            // or (byte) -91 (which is 165 - 256)
            ClassicAssert.AreEqual(3, rgb3.GetRgbWithTint().Length);
            ClassicAssert.AreEqual(-91, (sbyte)rgb3.GetRgbWithTint()[0]);
            ClassicAssert.AreEqual(-91, (sbyte)rgb3.GetRgbWithTint()[1]);
            ClassicAssert.AreEqual(-91, (sbyte)rgb3.GetRgbWithTint()[2]);

            // Set the colour to black, will Get translated internally
            // (Excel stores 3 colour white and black wrong!)
            // Set the color to black (no theme).
            rgb3.SetRgb(new byte[] { 0, 0, 0 });
            ClassicAssert.AreEqual("FF000000", rgb3.ARGBHex);
            ClassicAssert.AreEqual(0, rgb3.GetCTColor().GetRgb()[0]);
            ClassicAssert.AreEqual(0, rgb3.GetCTColor().GetRgb()[1]);
            ClassicAssert.AreEqual(0, rgb3.GetCTColor().GetRgb()[2]);

            // Set another, is fine
            rgb3.SetRgb(new byte[] { 16, 17, 18 });
            ClassicAssert.IsFalse(rgb3.HasAlpha);
            ClassicAssert.AreEqual("FF101112", rgb3.ARGBHex);
            ClassicAssert.AreEqual(0x10, rgb3.GetCTColor().GetRgb()[0]);
            ClassicAssert.AreEqual(0x11, rgb3.GetCTColor().GetRgb()[1]);
            ClassicAssert.AreEqual(0x12, rgb3.GetCTColor().GetRgb()[2]);
        }
        [Test]
        public void TestARGBColour()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48779.xlsx");

            // Check the CTColor is as expected
            XSSFColor rgb4 = (XSSFColor)wb.GetCellStyleAt(1).FillForegroundColorColor;
            ClassicAssert.AreEqual(false, rgb4.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(0, rgb4.GetCTColor().indexed);
            ClassicAssert.AreEqual(true, rgb4.GetCTColor().IsSetRgb());
            ClassicAssert.AreEqual(4, rgb4.GetCTColor().GetRgb().Length);

            // Now check the XSSFColor
            ClassicAssert.AreEqual(0, rgb4.Indexed);
            ClassicAssert.AreEqual(0.0, rgb4.Tint);
            ClassicAssert.IsFalse(rgb4.HasTint);
            ClassicAssert.IsTrue(rgb4.HasAlpha);

            ClassicAssert.AreEqual("FFFF0000", rgb4.ARGBHex);
            ClassicAssert.AreEqual(3, rgb4.RGB.Length);
            ClassicAssert.AreEqual(255, rgb4.RGB[0]);
            ClassicAssert.AreEqual(0, rgb4.RGB[1]);
            ClassicAssert.AreEqual(0, rgb4.RGB[2]);

            ClassicAssert.AreEqual(4, rgb4.ARGB.Length);
            ClassicAssert.AreEqual(255, rgb4.ARGB[0]);
            ClassicAssert.AreEqual(255, rgb4.ARGB[1]);
            ClassicAssert.AreEqual(0, rgb4.ARGB[2]);
            ClassicAssert.AreEqual(0, rgb4.ARGB[3]);

            // Tint doesn't have the alpha
            ClassicAssert.AreEqual(3, rgb4.GetRgbWithTint().Length);
            ClassicAssert.AreEqual(255, rgb4.GetRgbWithTint()[0]);
            ClassicAssert.AreEqual(0, rgb4.GetRgbWithTint()[1]);
            ClassicAssert.AreEqual(0, rgb4.GetRgbWithTint()[2]);


            // Turn on tinting, and check it behaves
            // TODO These values are suspected to be wrong...
            rgb4.Tint = (0.4);
            ClassicAssert.IsTrue(rgb4.HasTint);
            ClassicAssert.AreEqual(0.4, rgb4.Tint);

            ClassicAssert.AreEqual(3, rgb4.GetRgbWithTint().Length);
            ClassicAssert.AreEqual(255, rgb4.GetRgbWithTint()[0]);
            ClassicAssert.AreEqual(102, rgb4.GetRgbWithTint()[1]);
            ClassicAssert.AreEqual(102, rgb4.GetRgbWithTint()[2]);

            wb.Close();
        }

        [Test]
        public void TestCustomIndexedColour()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("customIndexedColors.xlsx");
            XSSFCell cell = wb.GetSheetAt(1).GetRow(0).GetCell(0) as XSSFCell;
            XSSFColor color = cell.CellStyle.FillForegroundColorColor as XSSFColor;
            CT_Colors ctColors = wb.GetStylesSource().GetCTStylesheet().colors;

            CT_RgbColor ctRgbColor = ctColors.indexedColors[color.Index];

            string hexRgb = ctRgbColor.rgbHex;

            ClassicAssert.AreEqual(hexRgb, color.ARGBHex);
        }
    }

}
