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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.OpenXmlFormats.Spreadsheet;
namespace NPOI.XSSF.UserModel.Extensions
{

    [TestClass]
    public class TestXSSFCellFill
    {
        [TestMethod]
        public void TestGetFillBackgroundColor()
        {
            CT_Fill ctFill = new CT_Fill();
            XSSFCellFill cellFill = new XSSFCellFill(ctFill);
            CT_PatternFill ctPatternFill = ctFill.AddNewPatternFill();
            CT_Color bgColor = ctPatternFill.AddNewBgColor();
            Assert.IsNotNull(cellFill.GetFillBackgroundColor());
            bgColor.indexed = 2;
            Assert.AreEqual(2, cellFill.GetFillBackgroundColor().GetIndexed());
        }
        [TestMethod]
        public void TestGetFillForegroundColor()
        {
            CT_Fill ctFill = new CT_Fill();
            XSSFCellFill cellFill = new XSSFCellFill(ctFill);
            CT_PatternFill ctPatternFill = ctFill.AddNewPatternFill();
            CT_Color fgColor = ctPatternFill.AddNewFgColor();
            Assert.IsNotNull(cellFill.GetFillForegroundColor());
            fgColor.indexed = 8;
            Assert.AreEqual(8, cellFill.GetFillForegroundColor().GetIndexed());
        }
        [TestMethod]
        public void TestGetSetPatternType()
        {
            CT_Fill ctFill = new CT_Fill();
            XSSFCellFill cellFill = new XSSFCellFill(ctFill);
            CT_PatternFill ctPatternFill = ctFill.AddNewPatternFill();
            ctPatternFill.patternType = (ST_PatternType.solid);
            //Assert.AreEqual(FillPatternType.SOLID_FOREGROUND.ordinal(), cellFill.GetPatternType().ordinal());
        }
        [TestMethod]
        public void TestGetNotModifies()
        {
            CT_Fill ctFill = new CT_Fill();
            XSSFCellFill cellFill = new XSSFCellFill(ctFill);
            CT_PatternFill ctPatternFill = ctFill.AddNewPatternFill();
            ctPatternFill.patternType = (ST_PatternType.darkDown);
            Assert.AreEqual(ST_PatternType.darkDown, cellFill.GetPatternType());
        }
        [TestMethod]
        public void TestColorFromTheme()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("styles.xlsx");
            XSSFCell cellWithThemeColor = (XSSFCell)wb.GetSheetAt(0).GetRow(10).GetCell(0);
            //color RGB will be extracted from theme
            XSSFColor foregroundColor = (XSSFColor)((XSSFCellStyle)cellWithThemeColor.CellStyle).FillForegroundColorColor;
            byte[] rgb = foregroundColor.GetRgb();
            byte[] rgbWithTint = foregroundColor.GetRgbWithTint();
            Assert.AreEqual(rgb[0], -18);
            Assert.AreEqual(rgb[1], -20);
            Assert.AreEqual(rgb[2], -31);
            Assert.AreEqual(rgbWithTint[0], -12);
            Assert.AreEqual(rgbWithTint[1], -13);
            Assert.AreEqual(rgbWithTint[2], -20);
        }
    }

}