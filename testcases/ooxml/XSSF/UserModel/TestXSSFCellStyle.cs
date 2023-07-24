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

using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel.Extensions;
using NUnit.Framework;
using SixLabors.ImageSharp;

namespace TestCases.XSSF.UserModel
{

    [TestFixture]
    public class TestXSSFCellStyle
    {

        private StylesTable stylesTable;
        private CT_Border ctBorderA;
        private CT_Fill ctFill;
        private CT_Font ctFont;
        private CT_Xf cellStyleXf;
        private CT_Xf cellXf;
        private CT_CellXfs cellXfs;
        private XSSFCellStyle cellStyle;
        private CT_Stylesheet ctStylesheet;

        [SetUp]
        public void SetUp()
        {
            stylesTable = new StylesTable();

            ctStylesheet = stylesTable.GetCTStylesheet();

            ctBorderA = new CT_Border();
            XSSFCellBorder borderA = new XSSFCellBorder(ctBorderA);
            long borderId = stylesTable.PutBorder(borderA);
            Assert.AreEqual(1, borderId);

            XSSFCellBorder borderB = new XSSFCellBorder();
            Assert.AreEqual(1, stylesTable.PutBorder(borderB));

            ctFill = new CT_Fill();
            XSSFCellFill fill = new XSSFCellFill(ctFill);
            long fillId = stylesTable.PutFill(fill);
            Assert.AreEqual(2, fillId);

            ctFont = new CT_Font();
            XSSFFont font = new XSSFFont(ctFont);
            long fontId = stylesTable.PutFont(font);
            Assert.AreEqual(1, fontId);

            cellStyleXf = ctStylesheet.AddNewCellStyleXfs().AddNewXf();
            cellStyleXf.borderId = 1;
            cellStyleXf.fillId = 1;
            cellStyleXf.fontId = 1;

            cellXfs = ctStylesheet.AddNewCellXfs();
            cellXf = cellXfs.AddNewXf();
            cellXf.xfId = (1);
            cellXf.borderId = (1);
            cellXf.fillId = (1);
            cellXf.fontId = (1);
            stylesTable.PutCellStyleXf(cellStyleXf);
            stylesTable.PutCellXf(cellXf);
            cellStyle = new XSSFCellStyle(1, 1, stylesTable, null);

            Assert.IsNotNull(stylesTable.GetFillAt(1).GetCTFill().patternFill);
            Assert.AreEqual(ST_PatternType.darkGray, stylesTable.GetFillAt(1).GetCTFill().patternFill.patternType);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetBorderBottom()
        {
            //default values
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderBottom);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderBottom = (BorderStyle.Medium);
            Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderBottom);
            //a new border has been Added
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the Created border
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            Assert.AreEqual(ST_BorderStyle.medium, ctBorder.bottom.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderBottom = (BorderStyle.Medium);
                Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderBottom);
            }
            Assert.AreEqual((uint)borderId, cellStyle.GetCoreXf().borderId);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            Assert.AreSame(ctBorder, stylesTable.GetBorderAt(borderId).GetCTBorder());

            //setting border to none Removes the <bottom> element
            cellStyle.BorderBottom = (BorderStyle.None);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = (int)cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            Assert.IsFalse(ctBorder.IsSetBottom());
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestSetServeralBordersOnSameCell()
        {
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderRight);
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderLeft);
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderTop);
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderBottom);
            Assert.AreEqual(2, stylesTable.GetBorders().Count);

            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.BorderLeft = BorderStyle.DashDotDot;
            cellStyle.LeftBorderColor = HSSFColor.Green.Index;
            cellStyle.BorderRight = BorderStyle.Hair;
            cellStyle.RightBorderColor = HSSFColor.Blue.Index;
            cellStyle.BorderTop = BorderStyle.MediumDashed;
            cellStyle.TopBorderColor = HSSFColor.Orange.Index;
            //only one border style should be generated
            Assert.AreEqual(3, stylesTable.GetBorders().Count);

        }
        [Ignore("not found in poi")]
        [Test]
        public void TestGetSetBorderDiagonal()
        {
            Assert.AreEqual(BorderDiagonal.None, cellStyle.BorderDiagonal);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderDiagonalLineStyle = BorderStyle.Medium;
            cellStyle.BorderDiagonalColor = HSSFColor.Red.Index;
            cellStyle.BorderDiagonal = BorderDiagonal.Backward;

            Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderDiagonalLineStyle);
            //a new border has been added
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);

            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.AreEqual(ST_BorderStyle.medium, ctBorder.diagonal.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderDiagonal = BorderDiagonal.Backward;
                Assert.AreEqual(BorderDiagonal.Backward, cellStyle.BorderDiagonal);
            }
            Assert.AreEqual(borderId, cellStyle.GetCoreXf().borderId);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            Assert.AreSame(ctBorder, stylesTable.GetBorderAt((int)borderId).GetCTBorder());

            cellStyle.BorderDiagonal = (BorderDiagonal.None);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.IsFalse(ctBorder.IsSetDiagonal());
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetBorderRight()
        {
            //default values
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderRight);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderRight = (BorderStyle.Medium);
            Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderRight);
            //a new border has been Added
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the Created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.AreEqual(ST_BorderStyle.medium, ctBorder.right.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderRight = (BorderStyle.Medium);
                Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderRight);
            }
            Assert.AreEqual(borderId, cellStyle.GetCoreXf().borderId);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            Assert.AreSame(ctBorder, stylesTable.GetBorderAt((int)borderId).GetCTBorder());

            //setting border to none Removes the <right> element
            cellStyle.BorderRight = (BorderStyle.None);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.IsFalse(ctBorder.IsSetRight());
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetBorderLeft()
        {
            //default values
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderLeft);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderLeft = (BorderStyle.Medium);
            Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderLeft);
            //a new border has been Added
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the Created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.AreEqual(ST_BorderStyle.medium, ctBorder.left.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderLeft = (BorderStyle.Medium);
                Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderLeft);
            }
            Assert.AreEqual(borderId, cellStyle.GetCoreXf().borderId);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            Assert.AreSame(ctBorder, stylesTable.GetBorderAt((int)borderId).GetCTBorder());

            //setting border to none Removes the <left> element
            cellStyle.BorderLeft = (BorderStyle.None);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.IsFalse(ctBorder.IsSetLeft());
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetBorderTop()
        {
            //default values
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderTop);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderTop = BorderStyle.Medium;
            Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderTop);
            //a new border has been Added
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the Created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.AreEqual(ST_BorderStyle.medium, ctBorder.top.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderTop = BorderStyle.Medium;
                Assert.AreEqual(BorderStyle.Medium, cellStyle.BorderTop);
            }
            Assert.AreEqual((uint)borderId, cellStyle.GetCoreXf().borderId);
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            Assert.AreSame(ctBorder, stylesTable.GetBorderAt((int)borderId).GetCTBorder());

            //setting border to none Removes the <top> element
            cellStyle.BorderTop = BorderStyle.None;
            Assert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.IsFalse(ctBorder.IsSetTop());
        }

        private void TestGetSetBorderXMLBean(BorderStyle border, ST_BorderStyle expected)
        {
            cellStyle.BorderTop = border;
            Assert.AreEqual(border, cellStyle.BorderTop);
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            Assert.AreEqual(expected, ctBorder.top.style);
        }
        // Border Styles, in BorderStyle/ST_BorderStyle enum order
        [Test]
        public void TestGetSetBorderNone() {
            cellStyle.BorderTop = BorderStyle.None;
            Assert.AreEqual(BorderStyle.None, cellStyle.BorderTop);
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            Assert.IsNull(ctBorder.top);
            // no border style and ST_BorderStyle.NONE are equivalent
            // POI prefers to unset the border style than explicitly set it ST_BorderStyle.NONE
        }
        [Test]
        public void TestGetSetBorderThin()
        {
            TestGetSetBorderXMLBean(BorderStyle.Thin, ST_BorderStyle.thin);
        }
        [Test]
        public void TestGetSetBorderMedium()
        {
            TestGetSetBorderXMLBean(BorderStyle.Medium, ST_BorderStyle.medium);
        }
        [Test]
        public void TestGetSetBorderDashed()
        {
            TestGetSetBorderXMLBean(BorderStyle.Dashed, ST_BorderStyle.dashed);
        }
        [Test]
        public void TestGetSetBorderDotted()
        {
            TestGetSetBorderXMLBean(BorderStyle.Dotted, ST_BorderStyle.dotted);
        }
        [Test]
        public void TestGetSetBorderThick()
        {
            TestGetSetBorderXMLBean(BorderStyle.Thick, ST_BorderStyle.thick);
        }
        [Test]
        public void TestGetSetBorderDouble()
        {
            TestGetSetBorderXMLBean(BorderStyle.Double, ST_BorderStyle.@double);
        }
        [Test]
        public void TestGetSetBorderHair()
        {
            TestGetSetBorderXMLBean(BorderStyle.Hair, ST_BorderStyle.hair);
        }
        [Test]
        public void TestGetSetBorderMediumDashed()
        {
            TestGetSetBorderXMLBean(BorderStyle.MediumDashed, ST_BorderStyle.mediumDashed);
        }

        [Test]
        public void TestGetSetBorderDashDot()
        {
            TestGetSetBorderXMLBean(BorderStyle.DashDot, ST_BorderStyle.dashDot);
        }
        [Test]
        public void TestGetSetBorderMediumDashDot()
        {
            TestGetSetBorderXMLBean(BorderStyle.MediumDashDot, ST_BorderStyle.mediumDashDot);
        }
        [Test]
        public void TestGetSetBorderDashDotDot()
        {
            TestGetSetBorderXMLBean(BorderStyle.DashDotDot, ST_BorderStyle.dashDotDot);
        }
        
        [Test]
        public void TestGetSetBorderMediumDashDotDot()
        {
            TestGetSetBorderXMLBean(BorderStyle.MediumDashDotDot, ST_BorderStyle.mediumDashDotDot);
        }
        
        [Test]
        public void TestGetSetBorderSlantDashDot()
        {
            TestGetSetBorderXMLBean(BorderStyle.SlantedDashDot, ST_BorderStyle.slantDashDot);
        }
        

        [Test]
        public void TestGetSetBottomBorderColor()
        {
            //defaults
            Assert.AreEqual(IndexedColors.Black.Index, cellStyle.BottomBorderColor);
            Assert.IsNull(cellStyle.BottomBorderXSSFColor);

            int num = stylesTable.GetBorders().Count;

            XSSFColor clr;

            //setting indexed color
            cellStyle.BottomBorderColor = (IndexedColors.BlueGrey.Index);
            Assert.AreEqual(IndexedColors.BlueGrey.Index, cellStyle.BottomBorderColor);
            clr = cellStyle.BottomBorderXSSFColor;
            Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
            Assert.AreEqual(IndexedColors.BlueGrey.Index, clr.Indexed);
            //a new border was Added to the styles table
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

            //id of the Created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            Assert.AreEqual((uint)IndexedColors.BlueGrey.Index, ctBorder.bottom.color.indexed);

            //setting XSSFColor
            num = stylesTable.GetBorders().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetBottomBorderColor(clr);
            Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.BottomBorderXSSFColor.GetCTColor().ToString());
            byte[] rgb = cellStyle.BottomBorderXSSFColor.RGB;
            Assert.AreEqual(Color.Cyan, Color.FromRgb((byte)(rgb[0] & 0xFF), (byte)(rgb[1] & 0xFF), (byte)(rgb[2] & 0xFF)));
            //another border was Added to the styles table
            Assert.AreEqual(num+1, stylesTable.GetBorders().Count);

            //passing null unsets the color
            cellStyle.SetBottomBorderColor(null);
            Assert.IsNull(cellStyle.BottomBorderXSSFColor);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetTopBorderColor()
        {
            //defaults
            Assert.AreEqual(IndexedColors.Black.Index, cellStyle.TopBorderColor);
            Assert.IsNull(cellStyle.TopBorderXSSFColor);

            int num = stylesTable.GetBorders().Count;

            XSSFColor clr;

            //setting indexed color
            cellStyle.TopBorderColor = (IndexedColors.BlueGrey.Index);
            Assert.AreEqual(IndexedColors.BlueGrey.Index, cellStyle.TopBorderColor);
            clr = cellStyle.TopBorderXSSFColor;
            Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
            Assert.AreEqual(IndexedColors.BlueGrey.Index, clr.Indexed);
            //a new border was added to the styles table
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

            //id of the Created border
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            Assert.AreEqual((uint)IndexedColors.BlueGrey.Index, ctBorder.top.color.indexed);

            //setting XSSFColor
            num = stylesTable.GetBorders().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetTopBorderColor(clr);
            Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.TopBorderXSSFColor.GetCTColor().ToString());
            byte[] rgb = cellStyle.TopBorderXSSFColor.RGB;
            Assert.AreEqual(Color.Cyan, Color.FromRgb(rgb[0], rgb[1], rgb[2]));
            //another border was added to the styles table
            Assert.AreEqual(num, stylesTable.GetBorders().Count);

            //passing null unsets the color
            cellStyle.SetTopBorderColor(null);
            Assert.IsNull(cellStyle.TopBorderXSSFColor);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetLeftBorderColor()
        {
            //defaults
            Assert.AreEqual(IndexedColors.Black.Index, cellStyle.LeftBorderColor);
            Assert.IsNull(cellStyle.LeftBorderXSSFColor);

            int num = stylesTable.GetBorders().Count;

            XSSFColor clr;

            //setting indexed color
            cellStyle.LeftBorderColor = (IndexedColors.BlueGrey.Index);
            Assert.AreEqual(IndexedColors.BlueGrey.Index, cellStyle.LeftBorderColor);
            clr = cellStyle.LeftBorderXSSFColor;
            Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
            Assert.AreEqual(IndexedColors.BlueGrey.Index, clr.Indexed);
            //a new border was Added to the styles table
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

            //id of the Created border
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            Assert.AreEqual((uint)IndexedColors.BlueGrey.Index, ctBorder.left.color.indexed);

            //setting XSSFColor
            num = stylesTable.GetBorders().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetLeftBorderColor(clr);
            Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.LeftBorderXSSFColor.GetCTColor().ToString());
            byte[] rgb = cellStyle.LeftBorderXSSFColor.RGB;
            Assert.AreEqual(Color.Cyan, Color.FromRgb((byte)(rgb[0] & 0xFF), (byte)(rgb[1] & 0xFF), (byte)(rgb[2] & 0xFF)));
            //another border was Added to the styles table
            Assert.AreEqual(num, stylesTable.GetBorders().Count);

            //passing null unsets the color
            cellStyle.SetLeftBorderColor(null);
            Assert.IsNull(cellStyle.LeftBorderXSSFColor);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetRightBorderColor()
        {
            //defaults
            Assert.AreEqual(IndexedColors.Black.Index, cellStyle.RightBorderColor);
            Assert.IsNull(cellStyle.RightBorderXSSFColor);

            int num = stylesTable.GetBorders().Count;

            XSSFColor clr;

            //setting indexed color
            cellStyle.RightBorderColor = (IndexedColors.BlueGrey.Index);
            Assert.AreEqual(IndexedColors.BlueGrey.Index, cellStyle.RightBorderColor);
            clr = cellStyle.RightBorderXSSFColor;
            Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
            Assert.AreEqual(IndexedColors.BlueGrey.Index, clr.Indexed);
            //a new border was Added to the styles table
            Assert.AreEqual(num + 1, stylesTable.GetBorders().Count);

            //id of the Created border
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            Assert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            Assert.AreEqual((uint)IndexedColors.BlueGrey.Index, ctBorder.right.color.indexed);

            //setting XSSFColor
            num = stylesTable.GetBorders().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetRightBorderColor(clr);
            Assert.AreEqual(clr.GetCTColor().ToString(), cellStyle.RightBorderXSSFColor.GetCTColor().ToString());
            byte[] rgb = cellStyle.RightBorderXSSFColor.RGB;
            Assert.AreEqual(Color.Cyan, Color.FromRgb((byte)(rgb[0] & 0xFF), (byte)(rgb[1] & 0xFF), (byte)(rgb[2] & 0xFF)));
            //another border was Added to the styles table
            Assert.AreEqual(num, stylesTable.GetBorders().Count);

            //passing null unsets the color
            cellStyle.SetRightBorderColor(null);
            Assert.IsNull(cellStyle.RightBorderXSSFColor);
        }


        [Test]
        public void TestGetSetFillBackgroundColor()
        {
            Assert.AreEqual(IndexedColors.Automatic.Index, cellStyle.FillBackgroundColor);
            Assert.IsNull(cellStyle.FillBackgroundColorColor);

            XSSFColor clr;

            int num = stylesTable.GetFills().Count;

            //setting indexed color
            cellStyle.FillBackgroundColor = (IndexedColors.Red.Index);
            Assert.AreEqual(IndexedColors.Red.Index, cellStyle.FillBackgroundColor);
            clr = (XSSFColor)cellStyle.FillBackgroundColorColor;
            Assert.IsTrue(clr.GetCTColor().IsSetIndexed());
            Assert.AreEqual(IndexedColors.Red.Index, clr.Indexed);
            //a new fill was Added to the styles table
            Assert.AreEqual(num + 1, stylesTable.GetFills().Count);

            //id of the Created border
            int FillId = (int)cellStyle.GetCoreXf().fillId;
            Assert.IsTrue(FillId > 0);
            //check changes in the underlying xml bean
            CT_Fill ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
            Assert.AreEqual((uint)IndexedColors.Red.Index, ctFill.GetPatternFill().bgColor.indexed);

            //setting XSSFColor
            num = stylesTable.GetFills().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetFillBackgroundColor(clr); // TODO this testcase assumes that cellStyle creates a new CT_Fill, but the implementation changes the existing style. - do not know whats right 8-(
            Assert.AreEqual(clr.GetCTColor().ToString(), ((XSSFColor)cellStyle.FillBackgroundColorColor).GetCTColor().ToString());
            byte[] rgb = ((XSSFColor)cellStyle.FillBackgroundColorColor).RGB;
            Assert.AreEqual(Color.Cyan, Color.FromRgb((byte)(rgb[0] & 0xFF), (byte)(rgb[1] & 0xFF), (byte)(rgb[2] & 0xFF)));
            //another border was added to the styles table
            Assert.AreEqual(num + 1, stylesTable.GetFills().Count);

            //passing null unsets the color
            cellStyle.SetFillBackgroundColor(null);
            Assert.IsNull(cellStyle.FillBackgroundColorColor);
            Assert.AreEqual(IndexedColors.Automatic.Index, cellStyle.FillBackgroundColor);
        }
        [Test]
        public void TestDefaultStyles()
        {

            XSSFWorkbook wb1 = new XSSFWorkbook();

            XSSFCellStyle style1 = (XSSFCellStyle)wb1.CreateCellStyle();
            Assert.AreEqual(IndexedColors.Automatic.Index, style1.FillBackgroundColor);
            Assert.IsNull(style1.FillBackgroundColorColor);

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb1));

            //compatibility with HSSF
            HSSFWorkbook wb2 = new HSSFWorkbook();
            HSSFCellStyle style2 = (HSSFCellStyle)wb2.CreateCellStyle();
            Assert.AreEqual(style2.FillBackgroundColor, style1.FillBackgroundColor);
            Assert.AreEqual(style2.FillForegroundColor, style1.FillForegroundColor);
            Assert.AreEqual(style2.FillPattern, style1.FillPattern);

            Assert.AreEqual(style2.LeftBorderColor, style1.LeftBorderColor);
            Assert.AreEqual(style2.TopBorderColor, style1.TopBorderColor);
            Assert.AreEqual(style2.RightBorderColor, style1.RightBorderColor);
            Assert.AreEqual(style2.BottomBorderColor, style1.BottomBorderColor);

            Assert.AreEqual(style2.BorderBottom, style1.BorderBottom);
            Assert.AreEqual(style2.BorderLeft, style1.BorderLeft);
            Assert.AreEqual(style2.BorderRight, style1.BorderRight);
            Assert.AreEqual(style2.BorderTop, style1.BorderTop);

            wb2.Close();
        }

        [Ignore("test")]
        public void TestGetFillForegroundColor()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            StylesTable styles = wb.GetStylesSource();
            Assert.AreEqual(1, wb.NumCellStyles);
            Assert.AreEqual(2, styles.GetFills().Count);

            XSSFCellStyle defaultStyle = (XSSFCellStyle)wb.GetCellStyleAt((short)0);
            Assert.AreEqual(IndexedColors.Automatic.Index, defaultStyle.FillForegroundColor);
            Assert.AreEqual(null, defaultStyle.FillForegroundColorColor);
            Assert.AreEqual(FillPattern.NoFill, defaultStyle.FillPattern);

            XSSFCellStyle customStyle = (XSSFCellStyle)wb.CreateCellStyle();

            customStyle.FillPattern = (FillPattern.SolidForeground);
            Assert.AreEqual(FillPattern.SolidForeground, customStyle.FillPattern);
            Assert.AreEqual(3, styles.GetFills().Count);

            customStyle.FillForegroundColor = (IndexedColors.BrightGreen.Index);
            Assert.AreEqual(IndexedColors.BrightGreen.Index, customStyle.FillForegroundColor);
            Assert.AreEqual(4, styles.GetFills().Count);

            for (int i = 0; i < 3; i++)
            {
                XSSFCellStyle style = (XSSFCellStyle)wb.CreateCellStyle();

                style.FillPattern = (FillPattern.SolidForeground);
                Assert.AreEqual(FillPattern.SolidForeground, style.FillPattern);
                Assert.AreEqual(4, styles.GetFills().Count);

                style.FillForegroundColor = (IndexedColors.BrightGreen.Index);
                Assert.AreEqual(IndexedColors.BrightGreen.Index, style.FillForegroundColor);
                Assert.AreEqual(4, styles.GetFills().Count);
            }

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        [Test]
        public void TestGetFillPattern()
        {
            //???
            //assertEquals(STPatternType.INT_DARK_GRAY-1, cellStyle.getFillPattern());

            Assert.AreEqual((int)ST_PatternType.darkGray, (int)cellStyle.FillPattern);

            int num = stylesTable.GetFills().Count;
            cellStyle.FillPattern = (FillPattern.SolidForeground);
            Assert.AreEqual(FillPattern.SolidForeground, cellStyle.FillPattern);
            Assert.AreEqual(num + 1, stylesTable.GetFills().Count);
            int FillId = (int)cellStyle.GetCoreXf().fillId;
            Assert.IsTrue(FillId > 0);
            //check Changes in the underlying xml bean
            CT_Fill ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
            Assert.AreEqual(ST_PatternType.solid, ctFill.GetPatternFill().patternType);

            //setting the same fill multiple time does not update the styles table
            for (int i = 0; i < 3; i++)
            {
                cellStyle.FillPattern = (FillPattern.SolidForeground);
            }
            Assert.AreEqual(num + 1, stylesTable.GetFills().Count);

            cellStyle.FillPattern = (FillPattern.NoFill);
            Assert.AreEqual(FillPattern.NoFill, cellStyle.FillPattern);
            FillId = (int)cellStyle.GetCoreXf().fillId;
            ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
            Assert.IsFalse(ctFill.GetPatternFill().IsSetPatternType());

        }
        [Test]
        public void TestGetFont()
        {
            Assert.IsNotNull(cellStyle.GetFont());
        }
        [Test]
        public void TestGetSetHidden()
        {
            Assert.IsFalse(cellStyle.IsHidden);
            cellStyle.IsHidden = (true);
            Assert.IsTrue(cellStyle.IsHidden);
            cellStyle.IsHidden = (false);
            Assert.IsFalse(cellStyle.IsHidden);
        }
        [Test]
        public void TestGetSetLocked()
        {
            Assert.IsTrue(cellStyle.IsLocked);
            cellStyle.IsLocked = (true);
            Assert.IsTrue(cellStyle.IsLocked);
            cellStyle.IsLocked = (false);
            Assert.IsFalse(cellStyle.IsLocked);
        }
        [Test]
        public void TestBug738()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("738.xlsx");

            ISheet sheet = wb.GetSheet("Sheet1");
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);
            Assert.IsTrue(cell.CellStyle.IsLocked);
            cell.CellStyle.IsLocked = true;
            Assert.IsTrue(cell.CellStyle.IsLocked);
            Assert.IsTrue(cell.CellStyle.IsHidden);
        }

        [Test]
        public void TestGetSetIndent()
        {
            
            Assert.AreEqual((short)0, cellStyle.Indention);
            cellStyle.Indention = ((short)3);
            Assert.AreEqual((short)3, cellStyle.Indention);
            cellStyle.Indention = ((short)13);
            Assert.AreEqual((short)13, cellStyle.Indention);
        }
        [Test]
        public void TestGetSetAlignement()
        {
            Assert.IsTrue(!cellStyle.GetCellAlignment().GetCTCellAlignment().horizontalSpecified);
            Assert.AreEqual(HorizontalAlignment.General, cellStyle.Alignment);

            cellStyle.Alignment = HorizontalAlignment.Left;
            Assert.AreEqual(HorizontalAlignment.Left, cellStyle.Alignment);
            Assert.AreEqual(ST_HorizontalAlignment.left, cellStyle.GetCellAlignment().GetCTCellAlignment().horizontal);

            cellStyle.Alignment = (HorizontalAlignment.Justify);
            Assert.AreEqual(HorizontalAlignment.Justify, cellStyle.Alignment);
            Assert.AreEqual(ST_HorizontalAlignment.justify, cellStyle.GetCellAlignment().GetCTCellAlignment().horizontal);

            cellStyle.Alignment = (HorizontalAlignment.Center);
            Assert.AreEqual(HorizontalAlignment.Center, cellStyle.Alignment);
            Assert.AreEqual(ST_HorizontalAlignment.center, cellStyle.GetCellAlignment().GetCTCellAlignment().horizontal);
        }
        [Test]
        public void TestGetSetVerticalAlignment()
        {
            Assert.AreEqual(VerticalAlignment.Bottom, cellStyle.VerticalAlignment);
            Assert.IsTrue(!cellStyle.GetCellAlignment().GetCTCellAlignment().verticalSpecified);

            cellStyle.VerticalAlignment = (VerticalAlignment.Top);
            Assert.AreEqual(VerticalAlignment.Top, cellStyle.VerticalAlignment);
            Assert.AreEqual(ST_VerticalAlignment.top, cellStyle.GetCellAlignment().GetCTCellAlignment().vertical);

            cellStyle.VerticalAlignment = (VerticalAlignment.Center);
            Assert.AreEqual(VerticalAlignment.Center, cellStyle.VerticalAlignment);
            Assert.AreEqual(ST_VerticalAlignment.center, cellStyle.GetCellAlignment().GetCTCellAlignment().vertical);

            cellStyle.VerticalAlignment = VerticalAlignment.Justify;
            Assert.AreEqual(VerticalAlignment.Justify, cellStyle.VerticalAlignment);
            Assert.AreEqual(ST_VerticalAlignment.justify, cellStyle.GetCellAlignment().GetCTCellAlignment().vertical);

            cellStyle.VerticalAlignment = (VerticalAlignment.Bottom);
            Assert.AreEqual(VerticalAlignment.Bottom, cellStyle.VerticalAlignment);
            Assert.AreEqual(ST_VerticalAlignment.bottom, cellStyle.GetCellAlignment().GetCTCellAlignment().vertical);
        }
        [Test]
        public void TestGetSetWrapText()
        {
            Assert.IsFalse(cellStyle.WrapText);
            cellStyle.WrapText = (true);
            Assert.IsTrue(cellStyle.WrapText);
            cellStyle.WrapText = (false);
            Assert.IsFalse(cellStyle.WrapText);
        }

        /**
         * Cloning one XSSFCellStyle onto Another, same XSSFWorkbook
         */
        [Test]
        public void TestCloneStyleSameWB()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            Assert.AreEqual(1, wb.NumberOfFonts);

            XSSFFont fnt = (XSSFFont)wb.CreateFont();
            fnt.FontName = ("TestingFont");
            Assert.AreEqual(2, wb.NumberOfFonts);

            XSSFCellStyle orig = (XSSFCellStyle)wb.CreateCellStyle();
            orig.Alignment = (HorizontalAlignment.Right);
            orig.SetFont(fnt);
            orig.DataFormat = (short)18;

            Assert.AreEqual(HorizontalAlignment.Right, orig.Alignment);
            Assert.AreEqual(fnt, orig.GetFont());
            Assert.AreEqual(18, orig.DataFormat);

            XSSFCellStyle clone = (XSSFCellStyle)wb.CreateCellStyle();
            Assert.AreNotEqual(HorizontalAlignment.Right, clone.Alignment);
            Assert.AreNotEqual(fnt, clone.GetFont());
            Assert.AreNotEqual(18, clone.DataFormat);

            clone.CloneStyleFrom(orig);
            Assert.AreEqual(HorizontalAlignment.Right, clone.Alignment);
            Assert.AreEqual(fnt, clone.GetFont());
            Assert.AreEqual(18, clone.DataFormat);
            Assert.AreEqual(2, wb.NumberOfFonts);
            Assert.AreEqual(clone.GetCoreXf(), wb.GetStylesSource().GetStyleAt(clone.Index).GetCoreXf(), "Should be same CoreXF after cloning");

            clone.Alignment = HorizontalAlignment.Left;
            clone.DataFormat = 17;
            Assert.AreEqual(HorizontalAlignment.Right, orig.Alignment);
            Assert.AreEqual(18, orig.DataFormat);

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        /**
         * Cloning one XSSFCellStyle onto Another, different XSSFWorkbooks
         */
        [Test]
        public void TestCloneStyleDiffWB()
        {
            XSSFWorkbook wbOrig = new XSSFWorkbook();
            Assert.AreEqual(1, wbOrig.NumberOfFonts);
            Assert.AreEqual(0, wbOrig.GetStylesSource().GetNumberFormats().Count);

            XSSFFont fnt = (XSSFFont)wbOrig.CreateFont();
            fnt.FontName = ("TestingFont");
            Assert.AreEqual(2, wbOrig.NumberOfFonts);
            Assert.AreEqual(0, wbOrig.GetStylesSource().GetNumberFormats().Count);

            XSSFDataFormat fmt = (XSSFDataFormat)wbOrig.CreateDataFormat();
            fmt.GetFormat("MadeUpOne");
            fmt.GetFormat("MadeUpTwo");

            XSSFCellStyle orig = (XSSFCellStyle)wbOrig.CreateCellStyle();
            orig.Alignment = (HorizontalAlignment.Right);
            orig.SetFont(fnt);
            orig.DataFormat = (fmt.GetFormat("Test##"));

            Assert.IsTrue(HorizontalAlignment.Right == orig.Alignment);
            Assert.IsTrue(fnt == orig.GetFont());
            Assert.IsTrue(fmt.GetFormat("Test##") == orig.DataFormat);

            Assert.AreEqual(2, wbOrig.NumberOfFonts);
            Assert.AreEqual(3, wbOrig.GetStylesSource().GetNumberFormats().Count);


            // Now a style on another workbook
            XSSFWorkbook wbClone = new XSSFWorkbook();
            Assert.AreEqual(1, wbClone.NumberOfFonts);
            Assert.AreEqual(0, wbClone.GetStylesSource().GetNumberFormats().Count);
            Assert.AreEqual(1, wbClone.NumCellStyles);

            XSSFDataFormat fmtClone = (XSSFDataFormat)wbClone.CreateDataFormat();
            XSSFCellStyle clone = (XSSFCellStyle)wbClone.CreateCellStyle();

            Assert.AreEqual(1, wbClone.NumberOfFonts);
            Assert.AreEqual(0, wbClone.GetStylesSource().GetNumberFormats().Count);

            Assert.IsFalse(HorizontalAlignment.Right == clone.Alignment);
            Assert.IsFalse("TestingFont" == clone.GetFont().FontName);

            clone.CloneStyleFrom(orig);

            Assert.AreEqual(2, wbClone.NumberOfFonts);
            Assert.AreEqual(2, wbClone.NumCellStyles);
            Assert.AreEqual(1, wbClone.GetStylesSource().GetNumberFormats().Count);

            Assert.AreEqual(HorizontalAlignment.Right, clone.Alignment);
            Assert.AreEqual("TestingFont", clone.GetFont().FontName);
            Assert.AreEqual(fmtClone.GetFormat("Test##"), clone.DataFormat);
            Assert.IsFalse(fmtClone.GetFormat("Test##") == fmt.GetFormat("Test##"));

            // Save it and re-check
            XSSFWorkbook wbReload = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wbClone);
            Assert.AreEqual(2, wbReload.NumberOfFonts);
            Assert.AreEqual(2, wbReload.NumCellStyles);
            Assert.AreEqual(1, wbReload.GetStylesSource().GetNumberFormats().Count);

            XSSFCellStyle reload = (XSSFCellStyle)wbReload.GetCellStyleAt((short)1);
            Assert.AreEqual(HorizontalAlignment.Right, reload.Alignment);
            Assert.AreEqual("TestingFont", reload.GetFont().FontName);
            Assert.AreEqual(fmtClone.GetFormat("Test##"), reload.DataFormat);
            Assert.IsFalse(fmtClone.GetFormat("Test##") == fmt.GetFormat("Test##"));

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wbOrig));
            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wbClone));
        }

        /**
          * Avoid ArrayIndexOutOfBoundsException  when creating cell style
          * in a workbook that has an empty xf table.
          */
        [Test]
        public void TestBug52348()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("52348.xlsx");
            StylesTable st = workbook.GetStylesSource();
            Assert.AreEqual(0, st.StyleXfsSize);


            XSSFCellStyle style = workbook.CreateCellStyle() as XSSFCellStyle; // no exception at this point
            Assert.IsNull(style.GetStyleXf());

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }
        /**
         * Avoid ArrayIndexOutOfBoundsException  when getting cell style
         * in a workbook that has an empty xf table.
         */
        [Test]
        public void TestBug55650()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("52348.xlsx");
            StylesTable st = workbook.GetStylesSource();
            Assert.AreEqual(0, st.StyleXfsSize);

            // no exception at this point
            XSSFCellStyle style = workbook.GetSheetAt(0).GetRow(0).GetCell(0).CellStyle as XSSFCellStyle;
            Assert.IsNull(style.GetStyleXf());

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }

        [Test]
        public void TestShrinkToFit()
        {
            // Existing file
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ShrinkToFit.xlsx");
            ISheet s = wb.GetSheetAt(0);
            IRow r = s.GetRow(0);
            ICellStyle cs = r.GetCell(0).CellStyle;

            Assert.AreEqual(true, cs.ShrinkToFit);

            // New file
            XSSFWorkbook wbOrig = new XSSFWorkbook();
            s = wbOrig.CreateSheet();
            r = s.CreateRow(0);

            cs = wbOrig.CreateCellStyle();
            cs.ShrinkToFit = (/*setter*/false);
            r.CreateCell(0).CellStyle = (/*setter*/cs);

            cs = wbOrig.CreateCellStyle();
            cs.ShrinkToFit = (/*setter*/true);
            r.CreateCell(1).CellStyle = (/*setter*/cs);

            // Write out1, Read, and check
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wbOrig) as XSSFWorkbook;
            s = wb.GetSheetAt(0);
            r = s.GetRow(0);
            Assert.AreEqual(false, r.GetCell(0).CellStyle.ShrinkToFit);
            Assert.AreEqual(true, r.GetCell(1).CellStyle.ShrinkToFit);

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wbOrig));
        }

        [Test]
        public void TestSetColor()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);

            //CreationHelper ch = wb.GetCreationHelper();
            IDataFormat format = wb.CreateDataFormat();
            ICell cell = row.CreateCell(1);
            cell.SetCellValue("somEvalue");
            ICellStyle cellStyle = wb.CreateCellStyle();


            cellStyle.DataFormat = (/*setter*/format.GetFormat("###0"));

            cellStyle.FillBackgroundColor = (/*setter*/IndexedColors.DarkBlue.Index);
            cellStyle.FillForegroundColor = (/*setter*/IndexedColors.DarkBlue.Index);
            cellStyle.FillPattern = FillPattern.SolidForeground;

            cellStyle.Alignment = HorizontalAlignment.Right;
            cellStyle.VerticalAlignment = VerticalAlignment.Top;

            cell.CellStyle = (/*setter*/cellStyle);

            /*OutputStream stream = new FileOutputStream("C:\\temp\\CellColor.xlsx");
            try {
                wb.Write(stream);
            } finally {
                stream.Close();
            }*/

            IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            ICell cellBack = wbBack.GetSheetAt(0).GetRow(0).GetCell(1);
            Assert.IsNotNull(cellBack);
            ICellStyle styleBack = cellBack.CellStyle;
            Assert.AreEqual(IndexedColors.DarkBlue.Index, styleBack.FillBackgroundColor);
            Assert.AreEqual(IndexedColors.DarkBlue.Index, styleBack.FillForegroundColor);
            Assert.AreEqual(HorizontalAlignment.Right, styleBack.Alignment);
            Assert.AreEqual(VerticalAlignment.Top, styleBack.VerticalAlignment);
            Assert.AreEqual(FillPattern.SolidForeground, styleBack.FillPattern);

            wbBack.Close();

            wb.Close();
        }

        public static void copyStyles(IWorkbook reference, IWorkbook target)
        {
            int numberOfStyles = reference.NumCellStyles;
            // don't copy default style (style index 0)
            for (int i = 1; i < numberOfStyles; i++)
            {
                ICellStyle referenceStyle = reference.GetCellStyleAt(i);
                ICellStyle targetStyle = target.CreateCellStyle();
                targetStyle.CloneStyleFrom(referenceStyle);
            }
            /*System.out.println("Reference : "+reference.NumCellStyles);
            System.out.println("Target    : "+target.NumCellStyles);*/
        }
        [Test]
        public void Test58084()
        {
            IWorkbook reference = XSSFTestDataSamples.OpenSampleWorkbook("template.xlsx");
            IWorkbook target = new XSSFWorkbook();
            copyStyles(reference, target);

            Assert.AreEqual(reference.NumCellStyles, target.NumCellStyles);
            ISheet sheet = target.CreateSheet();
            IRow row = sheet.CreateRow(0);
            int col = 0;
            for (short i = 1; i < target.NumCellStyles; i++)
            {
                ICell cell = row.CreateCell(col++);
                cell.SetCellValue("Coucou" + i);
                cell.CellStyle = (target.GetCellStyleAt(i));
            }
            /*OutputStream out = new FileOutputStream("C:\\temp\\58084.xlsx");
            target.write(out);
            out.close();*/
            IWorkbook copy = XSSFTestDataSamples.WriteOutAndReadBack(target);
            // previously this Assert.Failed because the border-element was not copied over 
            var xxx = copy.GetCellStyleAt((short)1).BorderBottom;

            copy.Close();

            target.Close();
            reference.Close();
        }


        [Test]
        public void Test58043()
        {
            Assert.AreEqual(0, cellStyle.Rotation);
            cellStyle.Rotation = ((short)89);
            Assert.AreEqual(89, cellStyle.Rotation);

            cellStyle.Rotation = ((short)90);
            Assert.AreEqual(90, cellStyle.Rotation);

            cellStyle.Rotation = ((short)179);
            Assert.AreEqual(179, cellStyle.Rotation);

            cellStyle.Rotation = ((short)180);
            Assert.AreEqual(180, cellStyle.Rotation);

            // negative values are mapped to the correct values for compatibility between HSSF and XSSF
            cellStyle.Rotation = ((short)-1);
            Assert.AreEqual(91, cellStyle.Rotation);

            cellStyle.Rotation = ((short)-89);
            Assert.AreEqual(179, cellStyle.Rotation);

            cellStyle.Rotation = ((short)-90);
            Assert.AreEqual(180, cellStyle.Rotation);
        }

        [Test]
        public void Bug58996_UsedToWorkIn3_11_ButNotIn3_13()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFCellStyle cellStyle = workbook.CreateCellStyle() as XSSFCellStyle;
            cellStyle.FillForegroundColorColor = (null);
            Assert.IsNull(cellStyle.FillForegroundColorColor);
            cellStyle.FillBackgroundColorColor = (null);
            Assert.IsNull(cellStyle.FillBackgroundColorColor);
            cellStyle.FillPattern = FillPattern.NoFill;;
            Assert.AreEqual(FillPattern.NoFill, cellStyle.FillPattern);
            cellStyle.SetBottomBorderColor(null);
            Assert.IsNull(cellStyle.BottomBorderXSSFColor);
            cellStyle.SetTopBorderColor(null);
            Assert.IsNull(cellStyle.TopBorderXSSFColor);
            cellStyle.SetLeftBorderColor(null);
            Assert.IsNull(cellStyle.LeftBorderXSSFColor);
            cellStyle.SetRightBorderColor(null);
            Assert.IsNull(cellStyle.RightBorderXSSFColor);
        }

    }
}