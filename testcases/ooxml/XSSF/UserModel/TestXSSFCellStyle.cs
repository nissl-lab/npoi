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
using NUnit.Framework;using NUnit.Framework.Legacy;
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
            ClassicAssert.AreEqual(1, borderId);

            XSSFCellBorder borderB = new XSSFCellBorder();
            ClassicAssert.AreEqual(1, stylesTable.PutBorder(borderB));

            ctFill = new CT_Fill();
            XSSFCellFill fill = new XSSFCellFill(ctFill);
            long fillId = stylesTable.PutFill(fill);
            ClassicAssert.AreEqual(2, fillId);

            ctFont = new CT_Font();
            XSSFFont font = new XSSFFont(ctFont);
            long fontId = stylesTable.PutFont(font);
            ClassicAssert.AreEqual(1, fontId);

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

            ClassicAssert.IsNotNull(stylesTable.GetFillAt(1).GetCTFill().patternFill);
            ClassicAssert.AreEqual(ST_PatternType.darkGray, stylesTable.GetFillAt(1).GetCTFill().patternFill.patternType);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetBorderBottom()
        {
            //default values
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderBottom);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderBottom = (BorderStyle.Medium);
            ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderBottom);
            //a new border has been Added
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the Created border
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            ClassicAssert.AreEqual(ST_BorderStyle.medium, ctBorder.bottom.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderBottom = (BorderStyle.Medium);
                ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderBottom);
            }
            ClassicAssert.AreEqual((uint)borderId, cellStyle.GetCoreXf().borderId);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            ClassicAssert.AreSame(ctBorder, stylesTable.GetBorderAt(borderId).GetCTBorder());

            //setting border to none Removes the <bottom> element
            cellStyle.BorderBottom = (BorderStyle.None);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = (int)cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            ClassicAssert.IsFalse(ctBorder.IsSetBottom());
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestSetServeralBordersOnSameCell()
        {
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderRight);
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderLeft);
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderTop);
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderBottom);
            ClassicAssert.AreEqual(2, stylesTable.GetBorders().Count);

            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.BorderLeft = BorderStyle.DashDotDot;
            cellStyle.LeftBorderColor = HSSFColor.Green.Index;
            cellStyle.BorderRight = BorderStyle.Hair;
            cellStyle.RightBorderColor = HSSFColor.Blue.Index;
            cellStyle.BorderTop = BorderStyle.MediumDashed;
            cellStyle.TopBorderColor = HSSFColor.Orange.Index;
            //only one border style should be generated
            ClassicAssert.AreEqual(3, stylesTable.GetBorders().Count);

        }
        [Ignore("not found in poi")]
        [Test]
        public void TestGetSetBorderDiagonal()
        {
            ClassicAssert.AreEqual(BorderDiagonal.None, cellStyle.BorderDiagonal);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderDiagonalLineStyle = BorderStyle.Medium;
            cellStyle.BorderDiagonalColor = HSSFColor.Red.Index;
            cellStyle.BorderDiagonal = BorderDiagonal.Backward;

            ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderDiagonalLineStyle);
            //a new border has been added
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);

            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.AreEqual(ST_BorderStyle.medium, ctBorder.diagonal.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderDiagonal = BorderDiagonal.Backward;
                ClassicAssert.AreEqual(BorderDiagonal.Backward, cellStyle.BorderDiagonal);
            }
            ClassicAssert.AreEqual(borderId, cellStyle.GetCoreXf().borderId);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            ClassicAssert.AreSame(ctBorder, stylesTable.GetBorderAt((int)borderId).GetCTBorder());

            cellStyle.BorderDiagonal = (BorderDiagonal.None);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.IsFalse(ctBorder.IsSetDiagonal());
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetBorderRight()
        {
            //default values
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderRight);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderRight = (BorderStyle.Medium);
            ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderRight);
            //a new border has been Added
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the Created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.AreEqual(ST_BorderStyle.medium, ctBorder.right.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderRight = (BorderStyle.Medium);
                ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderRight);
            }
            ClassicAssert.AreEqual(borderId, cellStyle.GetCoreXf().borderId);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            ClassicAssert.AreSame(ctBorder, stylesTable.GetBorderAt((int)borderId).GetCTBorder());

            //setting border to none Removes the <right> element
            cellStyle.BorderRight = (BorderStyle.None);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.IsFalse(ctBorder.IsSetRight());
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetBorderLeft()
        {
            //default values
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderLeft);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderLeft = (BorderStyle.Medium);
            ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderLeft);
            //a new border has been Added
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the Created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.AreEqual(ST_BorderStyle.medium, ctBorder.left.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderLeft = (BorderStyle.Medium);
                ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderLeft);
            }
            ClassicAssert.AreEqual(borderId, cellStyle.GetCoreXf().borderId);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            ClassicAssert.AreSame(ctBorder, stylesTable.GetBorderAt((int)borderId).GetCTBorder());

            //setting border to none Removes the <left> element
            cellStyle.BorderLeft = (BorderStyle.None);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.IsFalse(ctBorder.IsSetLeft());
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetBorderTop()
        {
            //default values
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderTop);

            int num = stylesTable.GetBorders().Count;
            cellStyle.BorderTop = BorderStyle.Medium;
            ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderTop);
            //a new border has been Added
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);
            //id of the Created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.AreEqual(ST_BorderStyle.medium, ctBorder.top.style);

            num = stylesTable.GetBorders().Count;
            //setting the same border multiple times should not change borderId
            for (int i = 0; i < 3; i++)
            {
                cellStyle.BorderTop = BorderStyle.Medium;
                ClassicAssert.AreEqual(BorderStyle.Medium, cellStyle.BorderTop);
            }
            ClassicAssert.AreEqual((uint)borderId, cellStyle.GetCoreXf().borderId);
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            ClassicAssert.AreSame(ctBorder, stylesTable.GetBorderAt((int)borderId).GetCTBorder());

            //setting border to none Removes the <top> element
            cellStyle.BorderTop = BorderStyle.None;
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);
            borderId = cellStyle.GetCoreXf().borderId;
            ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.IsFalse(ctBorder.IsSetTop());
        }

        private void TestGetSetBorderXMLBean(BorderStyle border, ST_BorderStyle expected)
        {
            cellStyle.BorderTop = border;
            ClassicAssert.AreEqual(border, cellStyle.BorderTop);
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            ClassicAssert.AreEqual(expected, ctBorder.top.style);
        }
        // Border Styles, in BorderStyle/ST_BorderStyle enum order
        [Test]
        public void TestGetSetBorderNone() {
            cellStyle.BorderTop = BorderStyle.None;
            ClassicAssert.AreEqual(BorderStyle.None, cellStyle.BorderTop);
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            ClassicAssert.IsNull(ctBorder.top);
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
            ClassicAssert.AreEqual(IndexedColors.Black.Index, cellStyle.BottomBorderColor);
            ClassicAssert.IsNull(cellStyle.BottomBorderXSSFColor);

            int num = stylesTable.GetBorders().Count;

            XSSFColor clr;

            //setting indexed color
            cellStyle.BottomBorderColor = (IndexedColors.BlueGrey.Index);
            ClassicAssert.AreEqual(IndexedColors.BlueGrey.Index, cellStyle.BottomBorderColor);
            clr = cellStyle.BottomBorderXSSFColor;
            ClassicAssert.IsTrue(clr.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(IndexedColors.BlueGrey.Index, clr.Indexed);
            //a new border was Added to the styles table
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);

            //id of the Created border
            uint borderId = cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt((int)borderId).GetCTBorder();
            ClassicAssert.AreEqual((uint)IndexedColors.BlueGrey.Index, ctBorder.bottom.color.indexed);

            //setting XSSFColor
            num = stylesTable.GetBorders().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetBottomBorderColor(clr);
            ClassicAssert.AreEqual(clr.GetCTColor().ToString(), cellStyle.BottomBorderXSSFColor.GetCTColor().ToString());
            byte[] rgb = cellStyle.BottomBorderXSSFColor.RGB;
            ClassicAssert.AreEqual(Color.Cyan, Color.FromRgb((byte)(rgb[0] & 0xFF), (byte)(rgb[1] & 0xFF), (byte)(rgb[2] & 0xFF)));
            //another border was Added to the styles table
            ClassicAssert.AreEqual(num+1, stylesTable.GetBorders().Count);

            //passing null unsets the color
            cellStyle.SetBottomBorderColor(null);
            ClassicAssert.IsNull(cellStyle.BottomBorderXSSFColor);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetTopBorderColor()
        {
            //defaults
            ClassicAssert.AreEqual(IndexedColors.Black.Index, cellStyle.TopBorderColor);
            ClassicAssert.IsNull(cellStyle.TopBorderXSSFColor);

            int num = stylesTable.GetBorders().Count;

            XSSFColor clr;

            //setting indexed color
            cellStyle.TopBorderColor = (IndexedColors.BlueGrey.Index);
            ClassicAssert.AreEqual(IndexedColors.BlueGrey.Index, cellStyle.TopBorderColor);
            clr = cellStyle.TopBorderXSSFColor;
            ClassicAssert.IsTrue(clr.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(IndexedColors.BlueGrey.Index, clr.Indexed);
            //a new border was added to the styles table
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);

            //id of the Created border
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            ClassicAssert.AreEqual((uint)IndexedColors.BlueGrey.Index, ctBorder.top.color.indexed);

            //setting XSSFColor
            num = stylesTable.GetBorders().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetTopBorderColor(clr);
            ClassicAssert.AreEqual(clr.GetCTColor().ToString(), cellStyle.TopBorderXSSFColor.GetCTColor().ToString());
            byte[] rgb = cellStyle.TopBorderXSSFColor.RGB;
            ClassicAssert.AreEqual(Color.Cyan, Color.FromRgb(rgb[0], rgb[1], rgb[2]));
            //another border was added to the styles table
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);

            //passing null unsets the color
            cellStyle.SetTopBorderColor(null);
            ClassicAssert.IsNull(cellStyle.TopBorderXSSFColor);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetLeftBorderColor()
        {
            //defaults
            ClassicAssert.AreEqual(IndexedColors.Black.Index, cellStyle.LeftBorderColor);
            ClassicAssert.IsNull(cellStyle.LeftBorderXSSFColor);

            int num = stylesTable.GetBorders().Count;

            XSSFColor clr;

            //setting indexed color
            cellStyle.LeftBorderColor = (IndexedColors.BlueGrey.Index);
            ClassicAssert.AreEqual(IndexedColors.BlueGrey.Index, cellStyle.LeftBorderColor);
            clr = cellStyle.LeftBorderXSSFColor;
            ClassicAssert.IsTrue(clr.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(IndexedColors.BlueGrey.Index, clr.Indexed);
            //a new border was Added to the styles table
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);

            //id of the Created border
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            ClassicAssert.AreEqual((uint)IndexedColors.BlueGrey.Index, ctBorder.left.color.indexed);

            //setting XSSFColor
            num = stylesTable.GetBorders().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetLeftBorderColor(clr);
            ClassicAssert.AreEqual(clr.GetCTColor().ToString(), cellStyle.LeftBorderXSSFColor.GetCTColor().ToString());
            byte[] rgb = cellStyle.LeftBorderXSSFColor.RGB;
            ClassicAssert.AreEqual(Color.Cyan, Color.FromRgb((byte)(rgb[0] & 0xFF), (byte)(rgb[1] & 0xFF), (byte)(rgb[2] & 0xFF)));
            //another border was Added to the styles table
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);

            //passing null unsets the color
            cellStyle.SetLeftBorderColor(null);
            ClassicAssert.IsNull(cellStyle.LeftBorderXSSFColor);
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestGetSetRightBorderColor()
        {
            //defaults
            ClassicAssert.AreEqual(IndexedColors.Black.Index, cellStyle.RightBorderColor);
            ClassicAssert.IsNull(cellStyle.RightBorderXSSFColor);

            int num = stylesTable.GetBorders().Count;

            XSSFColor clr;

            //setting indexed color
            cellStyle.RightBorderColor = (IndexedColors.BlueGrey.Index);
            ClassicAssert.AreEqual(IndexedColors.BlueGrey.Index, cellStyle.RightBorderColor);
            clr = cellStyle.RightBorderXSSFColor;
            ClassicAssert.IsTrue(clr.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(IndexedColors.BlueGrey.Index, clr.Indexed);
            //a new border was Added to the styles table
            ClassicAssert.AreEqual(num + 1, stylesTable.GetBorders().Count);

            //id of the Created border
            int borderId = (int)cellStyle.GetCoreXf().borderId;
            ClassicAssert.IsTrue(borderId > 0);
            //check Changes in the underlying xml bean
            CT_Border ctBorder = stylesTable.GetBorderAt(borderId).GetCTBorder();
            ClassicAssert.AreEqual((uint)IndexedColors.BlueGrey.Index, ctBorder.right.color.indexed);

            //setting XSSFColor
            num = stylesTable.GetBorders().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetRightBorderColor(clr);
            ClassicAssert.AreEqual(clr.GetCTColor().ToString(), cellStyle.RightBorderXSSFColor.GetCTColor().ToString());
            byte[] rgb = cellStyle.RightBorderXSSFColor.RGB;
            ClassicAssert.AreEqual(Color.Cyan, Color.FromRgb((byte)(rgb[0] & 0xFF), (byte)(rgb[1] & 0xFF), (byte)(rgb[2] & 0xFF)));
            //another border was Added to the styles table
            ClassicAssert.AreEqual(num, stylesTable.GetBorders().Count);

            //passing null unsets the color
            cellStyle.SetRightBorderColor(null);
            ClassicAssert.IsNull(cellStyle.RightBorderXSSFColor);
        }


        [Test]
        public void TestGetSetFillBackgroundColor()
        {
            ClassicAssert.AreEqual(IndexedColors.Automatic.Index, cellStyle.FillBackgroundColor);
            ClassicAssert.IsNull(cellStyle.FillBackgroundColorColor);

            XSSFColor clr;

            int num = stylesTable.GetFills().Count;

            //setting indexed color
            cellStyle.FillBackgroundColor = (IndexedColors.Red.Index);
            ClassicAssert.AreEqual(IndexedColors.Red.Index, cellStyle.FillBackgroundColor);
            clr = (XSSFColor)cellStyle.FillBackgroundColorColor;
            ClassicAssert.IsTrue(clr.GetCTColor().IsSetIndexed());
            ClassicAssert.AreEqual(IndexedColors.Red.Index, clr.Indexed);
            //a new fill was Added to the styles table
            ClassicAssert.AreEqual(num + 1, stylesTable.GetFills().Count);

            //id of the Created border
            int FillId = (int)cellStyle.GetCoreXf().fillId;
            ClassicAssert.IsTrue(FillId > 0);
            //check changes in the underlying xml bean
            CT_Fill ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
            ClassicAssert.AreEqual((uint)IndexedColors.Red.Index, ctFill.GetPatternFill().bgColor.indexed);

            //setting XSSFColor
            num = stylesTable.GetFills().Count;
            clr = new XSSFColor(Color.Cyan);
            cellStyle.SetFillBackgroundColor(clr); // TODO this testcase assumes that cellStyle creates a new CT_Fill, but the implementation changes the existing style. - do not know whats right 8-(
            ClassicAssert.AreEqual(clr.GetCTColor().ToString(), ((XSSFColor)cellStyle.FillBackgroundColorColor).GetCTColor().ToString());
            byte[] rgb = ((XSSFColor)cellStyle.FillBackgroundColorColor).RGB;
            ClassicAssert.AreEqual(Color.Cyan, Color.FromRgb((byte)(rgb[0] & 0xFF), (byte)(rgb[1] & 0xFF), (byte)(rgb[2] & 0xFF)));
            //another border was added to the styles table
            ClassicAssert.AreEqual(num + 1, stylesTable.GetFills().Count);

            //passing null unsets the color
            cellStyle.SetFillBackgroundColor(null);
            ClassicAssert.IsNull(cellStyle.FillBackgroundColorColor);
            ClassicAssert.AreEqual(IndexedColors.Automatic.Index, cellStyle.FillBackgroundColor);
        }
        [Test]
        public void TestDefaultStyles()
        {

            XSSFWorkbook wb1 = new XSSFWorkbook();

            XSSFCellStyle style1 = (XSSFCellStyle)wb1.CreateCellStyle();
            ClassicAssert.AreEqual(IndexedColors.Automatic.Index, style1.FillBackgroundColor);
            ClassicAssert.IsNull(style1.FillBackgroundColorColor);

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb1));

            //compatibility with HSSF
            HSSFWorkbook wb2 = new HSSFWorkbook();
            HSSFCellStyle style2 = (HSSFCellStyle)wb2.CreateCellStyle();
            ClassicAssert.AreEqual(style2.FillBackgroundColor, style1.FillBackgroundColor);
            ClassicAssert.AreEqual(style2.FillForegroundColor, style1.FillForegroundColor);
            ClassicAssert.AreEqual(style2.FillPattern, style1.FillPattern);

            ClassicAssert.AreEqual(style2.LeftBorderColor, style1.LeftBorderColor);
            ClassicAssert.AreEqual(style2.TopBorderColor, style1.TopBorderColor);
            ClassicAssert.AreEqual(style2.RightBorderColor, style1.RightBorderColor);
            ClassicAssert.AreEqual(style2.BottomBorderColor, style1.BottomBorderColor);

            ClassicAssert.AreEqual(style2.BorderBottom, style1.BorderBottom);
            ClassicAssert.AreEqual(style2.BorderLeft, style1.BorderLeft);
            ClassicAssert.AreEqual(style2.BorderRight, style1.BorderRight);
            ClassicAssert.AreEqual(style2.BorderTop, style1.BorderTop);

            wb2.Close();
        }

        [Ignore("test")]
        public void TestGetFillForegroundColor()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            StylesTable styles = wb.GetStylesSource();
            ClassicAssert.AreEqual(1, wb.NumCellStyles);
            ClassicAssert.AreEqual(2, styles.GetFills().Count);

            XSSFCellStyle defaultStyle = (XSSFCellStyle)wb.GetCellStyleAt((short)0);
            ClassicAssert.AreEqual(IndexedColors.Automatic.Index, defaultStyle.FillForegroundColor);
            ClassicAssert.AreEqual(null, defaultStyle.FillForegroundColorColor);
            ClassicAssert.AreEqual(FillPattern.NoFill, defaultStyle.FillPattern);

            XSSFCellStyle customStyle = (XSSFCellStyle)wb.CreateCellStyle();

            customStyle.FillPattern = (FillPattern.SolidForeground);
            ClassicAssert.AreEqual(FillPattern.SolidForeground, customStyle.FillPattern);
            ClassicAssert.AreEqual(3, styles.GetFills().Count);

            customStyle.FillForegroundColor = (IndexedColors.BrightGreen.Index);
            ClassicAssert.AreEqual(IndexedColors.BrightGreen.Index, customStyle.FillForegroundColor);
            ClassicAssert.AreEqual(4, styles.GetFills().Count);

            for (int i = 0; i < 3; i++)
            {
                XSSFCellStyle style = (XSSFCellStyle)wb.CreateCellStyle();

                style.FillPattern = (FillPattern.SolidForeground);
                ClassicAssert.AreEqual(FillPattern.SolidForeground, style.FillPattern);
                ClassicAssert.AreEqual(4, styles.GetFills().Count);

                style.FillForegroundColor = (IndexedColors.BrightGreen.Index);
                ClassicAssert.AreEqual(IndexedColors.BrightGreen.Index, style.FillForegroundColor);
                ClassicAssert.AreEqual(4, styles.GetFills().Count);
            }

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        [Test]
        public void TestGetFillPattern()
        {
            //???
            //assertEquals(STPatternType.INT_DARK_GRAY-1, cellStyle.getFillPattern());

            ClassicAssert.AreEqual((int)ST_PatternType.darkGray, (int)cellStyle.FillPattern);

            int num = stylesTable.GetFills().Count;
            cellStyle.FillPattern = (FillPattern.SolidForeground);
            ClassicAssert.AreEqual(FillPattern.SolidForeground, cellStyle.FillPattern);
            ClassicAssert.AreEqual(num + 1, stylesTable.GetFills().Count);
            int FillId = (int)cellStyle.GetCoreXf().fillId;
            ClassicAssert.IsTrue(FillId > 0);
            //check Changes in the underlying xml bean
            CT_Fill ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
            ClassicAssert.AreEqual(ST_PatternType.solid, ctFill.GetPatternFill().patternType);

            //setting the same fill multiple time does not update the styles table
            for (int i = 0; i < 3; i++)
            {
                cellStyle.FillPattern = (FillPattern.SolidForeground);
            }
            ClassicAssert.AreEqual(num + 1, stylesTable.GetFills().Count);

            cellStyle.FillPattern = (FillPattern.NoFill);
            ClassicAssert.AreEqual(FillPattern.NoFill, cellStyle.FillPattern);
            FillId = (int)cellStyle.GetCoreXf().fillId;
            ctFill = stylesTable.GetFillAt(FillId).GetCTFill();
            ClassicAssert.IsFalse(ctFill.GetPatternFill().IsSetPatternType());

        }
        [Test]
        public void TestGetFont()
        {
            ClassicAssert.IsNotNull(cellStyle.GetFont());
        }
        [Test]
        public void TestGetSetHidden()
        {
            ClassicAssert.IsFalse(cellStyle.IsHidden);
            cellStyle.IsHidden = (true);
            ClassicAssert.IsTrue(cellStyle.IsHidden);
            cellStyle.IsHidden = (false);
            ClassicAssert.IsFalse(cellStyle.IsHidden);
        }
        [Test]
        public void TestGetSetLocked()
        {
            ClassicAssert.IsTrue(cellStyle.IsLocked);
            cellStyle.IsLocked = (true);
            ClassicAssert.IsTrue(cellStyle.IsLocked);
            cellStyle.IsLocked = (false);
            ClassicAssert.IsFalse(cellStyle.IsLocked);
        }
        [Test]
        public void TestBug738()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("738.xlsx");

            ISheet sheet = wb.GetSheet("Sheet1");
            IRow row = sheet.GetRow(0);
            ICell cell = row.GetCell(0);
            ClassicAssert.IsTrue(cell.CellStyle.IsLocked);
            cell.CellStyle.IsLocked = true;
            ClassicAssert.IsTrue(cell.CellStyle.IsLocked);
            ClassicAssert.IsTrue(cell.CellStyle.IsHidden);
        }

        [Test]
        public void TestGetSetIndent()
        {
            
            ClassicAssert.AreEqual((short)0, cellStyle.Indention);
            cellStyle.Indention = ((short)3);
            ClassicAssert.AreEqual((short)3, cellStyle.Indention);
            cellStyle.Indention = ((short)13);
            ClassicAssert.AreEqual((short)13, cellStyle.Indention);
        }
        [Test]
        public void TestGetSetAlignement()
        {
            ClassicAssert.IsTrue(!cellStyle.GetCellAlignment().GetCTCellAlignment().horizontalSpecified);
            ClassicAssert.AreEqual(HorizontalAlignment.General, cellStyle.Alignment);

            cellStyle.Alignment = HorizontalAlignment.Left;
            ClassicAssert.AreEqual(HorizontalAlignment.Left, cellStyle.Alignment);
            ClassicAssert.AreEqual(ST_HorizontalAlignment.left, cellStyle.GetCellAlignment().GetCTCellAlignment().horizontal);

            cellStyle.Alignment = (HorizontalAlignment.Justify);
            ClassicAssert.AreEqual(HorizontalAlignment.Justify, cellStyle.Alignment);
            ClassicAssert.AreEqual(ST_HorizontalAlignment.justify, cellStyle.GetCellAlignment().GetCTCellAlignment().horizontal);

            cellStyle.Alignment = (HorizontalAlignment.Center);
            ClassicAssert.AreEqual(HorizontalAlignment.Center, cellStyle.Alignment);
            ClassicAssert.AreEqual(ST_HorizontalAlignment.center, cellStyle.GetCellAlignment().GetCTCellAlignment().horizontal);
        }
        [Test]
        public void TestGetSetVerticalAlignment()
        {
            ClassicAssert.AreEqual(VerticalAlignment.Bottom, cellStyle.VerticalAlignment);
            ClassicAssert.IsTrue(!cellStyle.GetCellAlignment().GetCTCellAlignment().verticalSpecified);

            cellStyle.VerticalAlignment = (VerticalAlignment.Top);
            ClassicAssert.AreEqual(VerticalAlignment.Top, cellStyle.VerticalAlignment);
            ClassicAssert.AreEqual(ST_VerticalAlignment.top, cellStyle.GetCellAlignment().GetCTCellAlignment().vertical);

            cellStyle.VerticalAlignment = (VerticalAlignment.Center);
            ClassicAssert.AreEqual(VerticalAlignment.Center, cellStyle.VerticalAlignment);
            ClassicAssert.AreEqual(ST_VerticalAlignment.center, cellStyle.GetCellAlignment().GetCTCellAlignment().vertical);

            cellStyle.VerticalAlignment = VerticalAlignment.Justify;
            ClassicAssert.AreEqual(VerticalAlignment.Justify, cellStyle.VerticalAlignment);
            ClassicAssert.AreEqual(ST_VerticalAlignment.justify, cellStyle.GetCellAlignment().GetCTCellAlignment().vertical);

            cellStyle.VerticalAlignment = (VerticalAlignment.Bottom);
            ClassicAssert.AreEqual(VerticalAlignment.Bottom, cellStyle.VerticalAlignment);
            ClassicAssert.AreEqual(ST_VerticalAlignment.bottom, cellStyle.GetCellAlignment().GetCTCellAlignment().vertical);
        }
        [Test]
        public void TestGetSetWrapText()
        {
            ClassicAssert.IsFalse(cellStyle.WrapText);
            cellStyle.WrapText = (true);
            ClassicAssert.IsTrue(cellStyle.WrapText);
            cellStyle.WrapText = (false);
            ClassicAssert.IsFalse(cellStyle.WrapText);
        }

        /**
         * Cloning one XSSFCellStyle onto Another, same XSSFWorkbook
         */
        [Test]
        public void TestCloneStyleSameWB()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ClassicAssert.AreEqual(1, wb.NumberOfFonts);

            XSSFFont fnt = (XSSFFont)wb.CreateFont();
            fnt.FontName = ("TestingFont");
            ClassicAssert.AreEqual(2, wb.NumberOfFonts);

            XSSFCellStyle orig = (XSSFCellStyle)wb.CreateCellStyle();
            orig.Alignment = (HorizontalAlignment.Right);
            orig.SetFont(fnt);
            orig.DataFormat = (short)18;

            ClassicAssert.AreEqual(HorizontalAlignment.Right, orig.Alignment);
            ClassicAssert.AreEqual(fnt, orig.GetFont());
            ClassicAssert.AreEqual(18, orig.DataFormat);

            XSSFCellStyle clone = (XSSFCellStyle)wb.CreateCellStyle();
            ClassicAssert.AreNotEqual(HorizontalAlignment.Right, clone.Alignment);
            ClassicAssert.AreNotEqual(fnt, clone.GetFont());
            ClassicAssert.AreNotEqual(18, clone.DataFormat);

            clone.CloneStyleFrom(orig);
            ClassicAssert.AreEqual(HorizontalAlignment.Right, clone.Alignment);
            ClassicAssert.AreEqual(fnt, clone.GetFont());
            ClassicAssert.AreEqual(18, clone.DataFormat);
            ClassicAssert.AreEqual(2, wb.NumberOfFonts);
            ClassicAssert.AreEqual(clone.GetCoreXf(), wb.GetStylesSource().GetStyleAt(clone.Index).GetCoreXf(), "Should be same CoreXF after cloning");

            clone.Alignment = HorizontalAlignment.Left;
            clone.DataFormat = 17;
            ClassicAssert.AreEqual(HorizontalAlignment.Right, orig.Alignment);
            ClassicAssert.AreEqual(18, orig.DataFormat);

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        /**
         * Cloning one XSSFCellStyle onto Another, different XSSFWorkbooks
         */
        [Test]
        public void TestCloneStyleDiffWB()
        {
            XSSFWorkbook wbOrig = new XSSFWorkbook();
            ClassicAssert.AreEqual(1, wbOrig.NumberOfFonts);
            ClassicAssert.AreEqual(0, wbOrig.GetStylesSource().GetNumberFormats().Count);

            XSSFFont fnt = (XSSFFont)wbOrig.CreateFont();
            fnt.FontName = ("TestingFont");
            ClassicAssert.AreEqual(2, wbOrig.NumberOfFonts);
            ClassicAssert.AreEqual(0, wbOrig.GetStylesSource().GetNumberFormats().Count);

            XSSFDataFormat fmt = (XSSFDataFormat)wbOrig.CreateDataFormat();
            fmt.GetFormat("MadeUpOne");
            fmt.GetFormat("MadeUpTwo");

            XSSFCellStyle orig = (XSSFCellStyle)wbOrig.CreateCellStyle();
            orig.Alignment = (HorizontalAlignment.Right);
            orig.SetFont(fnt);
            orig.DataFormat = (fmt.GetFormat("Test##"));

            ClassicAssert.IsTrue(HorizontalAlignment.Right == orig.Alignment);
            ClassicAssert.IsTrue(fnt == orig.GetFont());
            ClassicAssert.IsTrue(fmt.GetFormat("Test##") == orig.DataFormat);

            ClassicAssert.AreEqual(2, wbOrig.NumberOfFonts);
            ClassicAssert.AreEqual(3, wbOrig.GetStylesSource().GetNumberFormats().Count);


            // Now a style on another workbook
            XSSFWorkbook wbClone = new XSSFWorkbook();
            ClassicAssert.AreEqual(1, wbClone.NumberOfFonts);
            ClassicAssert.AreEqual(0, wbClone.GetStylesSource().GetNumberFormats().Count);
            ClassicAssert.AreEqual(1, wbClone.NumCellStyles);

            XSSFDataFormat fmtClone = (XSSFDataFormat)wbClone.CreateDataFormat();
            XSSFCellStyle clone = (XSSFCellStyle)wbClone.CreateCellStyle();

            ClassicAssert.AreEqual(1, wbClone.NumberOfFonts);
            ClassicAssert.AreEqual(0, wbClone.GetStylesSource().GetNumberFormats().Count);

            ClassicAssert.IsFalse(HorizontalAlignment.Right == clone.Alignment);
            ClassicAssert.IsFalse("TestingFont" == clone.GetFont().FontName);

            clone.CloneStyleFrom(orig);

            ClassicAssert.AreEqual(2, wbClone.NumberOfFonts);
            ClassicAssert.AreEqual(2, wbClone.NumCellStyles);
            ClassicAssert.AreEqual(1, wbClone.GetStylesSource().GetNumberFormats().Count);

            ClassicAssert.AreEqual(HorizontalAlignment.Right, clone.Alignment);
            ClassicAssert.AreEqual("TestingFont", clone.GetFont().FontName);
            ClassicAssert.AreEqual(fmtClone.GetFormat("Test##"), clone.DataFormat);
            ClassicAssert.IsFalse(fmtClone.GetFormat("Test##") == fmt.GetFormat("Test##"));

            // Save it and re-check
            XSSFWorkbook wbReload = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wbClone);
            ClassicAssert.AreEqual(2, wbReload.NumberOfFonts);
            ClassicAssert.AreEqual(2, wbReload.NumCellStyles);
            ClassicAssert.AreEqual(1, wbReload.GetStylesSource().GetNumberFormats().Count);

            XSSFCellStyle reload = (XSSFCellStyle)wbReload.GetCellStyleAt((short)1);
            ClassicAssert.AreEqual(HorizontalAlignment.Right, reload.Alignment);
            ClassicAssert.AreEqual("TestingFont", reload.GetFont().FontName);
            ClassicAssert.AreEqual(fmtClone.GetFormat("Test##"), reload.DataFormat);
            ClassicAssert.IsFalse(fmtClone.GetFormat("Test##") == fmt.GetFormat("Test##"));

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wbOrig));
            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wbClone));
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
            ClassicAssert.AreEqual(0, st.StyleXfsSize);


            XSSFCellStyle style = workbook.CreateCellStyle() as XSSFCellStyle; // no exception at this point
            ClassicAssert.IsNull(style.GetStyleXf());

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
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
            ClassicAssert.AreEqual(0, st.StyleXfsSize);

            // no exception at this point
            XSSFCellStyle style = workbook.GetSheetAt(0).GetRow(0).GetCell(0).CellStyle as XSSFCellStyle;
            ClassicAssert.IsNull(style.GetStyleXf());

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }

        [Test]
        public void TestShrinkToFit()
        {
            // Existing file
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ShrinkToFit.xlsx");
            ISheet s = wb.GetSheetAt(0);
            IRow r = s.GetRow(0);
            ICellStyle cs = r.GetCell(0).CellStyle;

            ClassicAssert.AreEqual(true, cs.ShrinkToFit);

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
            ClassicAssert.AreEqual(false, r.GetCell(0).CellStyle.ShrinkToFit);
            ClassicAssert.AreEqual(true, r.GetCell(1).CellStyle.ShrinkToFit);

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wbOrig));
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
            ClassicAssert.IsNotNull(cellBack);
            ICellStyle styleBack = cellBack.CellStyle;
            ClassicAssert.AreEqual(IndexedColors.DarkBlue.Index, styleBack.FillBackgroundColor);
            ClassicAssert.AreEqual(IndexedColors.DarkBlue.Index, styleBack.FillForegroundColor);
            ClassicAssert.AreEqual(HorizontalAlignment.Right, styleBack.Alignment);
            ClassicAssert.AreEqual(VerticalAlignment.Top, styleBack.VerticalAlignment);
            ClassicAssert.AreEqual(FillPattern.SolidForeground, styleBack.FillPattern);

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

            ClassicAssert.AreEqual(reference.NumCellStyles, target.NumCellStyles);
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
            ClassicAssert.AreEqual(0, cellStyle.Rotation);
            cellStyle.Rotation = ((short)89);
            ClassicAssert.AreEqual(89, cellStyle.Rotation);

            cellStyle.Rotation = ((short)90);
            ClassicAssert.AreEqual(90, cellStyle.Rotation);

            cellStyle.Rotation = ((short)179);
            ClassicAssert.AreEqual(179, cellStyle.Rotation);

            cellStyle.Rotation = ((short)180);
            ClassicAssert.AreEqual(180, cellStyle.Rotation);

            // negative values are mapped to the correct values for compatibility between HSSF and XSSF
            cellStyle.Rotation = ((short)-1);
            ClassicAssert.AreEqual(91, cellStyle.Rotation);

            cellStyle.Rotation = ((short)-89);
            ClassicAssert.AreEqual(179, cellStyle.Rotation);

            cellStyle.Rotation = ((short)-90);
            ClassicAssert.AreEqual(180, cellStyle.Rotation);
        }

        [Test]
        public void Bug58996_UsedToWorkIn3_11_ButNotIn3_13()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFCellStyle cellStyle = workbook.CreateCellStyle() as XSSFCellStyle;
            cellStyle.FillForegroundColorColor = (null);
            ClassicAssert.IsNull(cellStyle.FillForegroundColorColor);
            cellStyle.FillBackgroundColorColor = (null);
            ClassicAssert.IsNull(cellStyle.FillBackgroundColorColor);
            cellStyle.FillPattern = FillPattern.NoFill;;
            ClassicAssert.AreEqual(FillPattern.NoFill, cellStyle.FillPattern);
            cellStyle.SetBottomBorderColor(null);
            ClassicAssert.IsNull(cellStyle.BottomBorderXSSFColor);
            cellStyle.SetTopBorderColor(null);
            ClassicAssert.IsNull(cellStyle.TopBorderXSSFColor);
            cellStyle.SetLeftBorderColor(null);
            ClassicAssert.IsNull(cellStyle.LeftBorderXSSFColor);
            cellStyle.SetRightBorderColor(null);
            ClassicAssert.IsNull(cellStyle.RightBorderXSSFColor);
        }

    }
}