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

using TestCases.SS.UserModel;
using NUnit.Framework;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.Util;
using System.Text;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System.Collections;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;

namespace TestCases.XSSF.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFPicture : BaseTestPicture
    {

        public TestXSSFPicture()
            : base(XSSFITestDataProvider.instance)
        {

        }
        [Test]
        public void Resize()
        {
            XSSFWorkbook wb = XSSFITestDataProvider.instance.OpenSampleWorkbook("resize_compare.xlsx") as XSSFWorkbook;
            XSSFDrawing dp = wb.GetSheetAt(0).CreateDrawingPatriarch() as XSSFDrawing;
            List<XSSFShape> pics = dp.GetShapes();
            XSSFPicture inpPic = (XSSFPicture)pics[(0)];
            XSSFPicture cmpPic = (XSSFPicture)pics[(0)];

            BaseTestResize(inpPic, cmpPic, 2.0, 2.0);
            wb.Close();
        }


        [Test]
        public void Create()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            byte[] jpegData = Encoding.UTF8.GetBytes("test jpeg data");

            IList pictures = wb.GetAllPictures();
            Assert.AreEqual(0, pictures.Count);

            int jpegIdx = wb.AddPicture(jpegData, PictureType.JPEG);
            Assert.AreEqual(1, pictures.Count);
            Assert.AreEqual("jpeg", ((XSSFPictureData)pictures[jpegIdx]).SuggestFileExtension());
            Assert.IsTrue(Arrays.Equals(jpegData, ((XSSFPictureData)pictures[jpegIdx]).Data));

            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            Assert.AreEqual(AnchorType.MoveAndResize, (AnchorType)anchor.AnchorType);
            anchor.AnchorType = AnchorType.DontMoveAndResize;
            Assert.AreEqual(AnchorType.DontMoveAndResize, (AnchorType)anchor.AnchorType);

            XSSFPicture shape = (XSSFPicture)drawing.CreatePicture(anchor, jpegIdx);
            Assert.IsTrue(anchor.Equals(shape.GetAnchor()));
            Assert.IsNotNull(shape.PictureData);
            Assert.IsTrue(Arrays.Equals(jpegData, shape.PictureData.Data));

            CT_TwoCellAnchor ctShapeHolder = (CT_TwoCellAnchor)drawing.GetCTDrawing().CellAnchors[0];
            // STEditAs.ABSOLUTE corresponds to ClientAnchor.DONT_MOVE_AND_RESIZE
            Assert.AreEqual(ST_EditAs.absolute, ctShapeHolder.editAs);
        }

        /**
         * Test that ShapeId in CTNonVisualDrawingProps is incremented
         *
         * See Bugzilla 50458
         */
        [Test]
        public void IncrementShapeId()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            byte[] jpegData = Encoding.UTF8.GetBytes("picture1");
            int jpegIdx = wb.AddPicture(jpegData, PictureType.JPEG);

            XSSFPicture shape1 = (XSSFPicture)drawing.CreatePicture(anchor, jpegIdx);
            Assert.AreEqual((uint)1, shape1.GetCTPicture().nvPicPr.cNvPr.id);

            jpegData = Encoding.UTF8.GetBytes("picture2");
            jpegIdx = wb.AddPicture(jpegData, PictureType.JPEG);
            XSSFPicture shape2 = (XSSFPicture)drawing.CreatePicture(anchor, jpegIdx);
            Assert.AreEqual((uint)2, shape2.GetCTPicture().nvPicPr.cNvPr.id);
        }

        /**
     * same image refrerred by mulitple sheets
     */
        [Test]
        public void multiRelationShips()
        {
            XSSFWorkbook wb = new XSSFWorkbook();

            byte[] pic1Data = Encoding.UTF8.GetBytes("test jpeg data");
            byte[] pic2Data = Encoding.UTF8.GetBytes("test png data");

            List<XSSFPictureData> pictures = wb.GetAllPictures() as List<XSSFPictureData>;
            Assert.AreEqual(0, pictures.Count);

            int pic1 = wb.AddPicture(pic1Data, XSSFWorkbook.PICTURE_TYPE_JPEG);
            int pic2 = wb.AddPicture(pic2Data, XSSFWorkbook.PICTURE_TYPE_PNG);

            XSSFSheet sheet1 = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing drawing1 = sheet1.CreateDrawingPatriarch() as XSSFDrawing;
            XSSFPicture shape1 = drawing1.CreatePicture(new XSSFClientAnchor(), pic1) as XSSFPicture;
            XSSFPicture shape2 = drawing1.CreatePicture(new XSSFClientAnchor(), pic2) as XSSFPicture;

            XSSFSheet sheet2 = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing drawing2 = sheet2.CreateDrawingPatriarch() as XSSFDrawing;
            XSSFPicture shape3 = drawing2.CreatePicture(new XSSFClientAnchor(), pic2) as XSSFPicture;
            XSSFPicture shape4 = drawing2.CreatePicture(new XSSFClientAnchor(), pic1) as XSSFPicture;

            Assert.AreEqual(2, pictures.Count);

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            pictures = wb.GetAllPictures() as List<XSSFPictureData>;
            Assert.AreEqual(2, pictures.Count);

            sheet1 = wb.GetSheetAt(0) as XSSFSheet;
            drawing1 = sheet1.CreateDrawingPatriarch() as XSSFDrawing;
            XSSFPicture shape11 = (XSSFPicture)drawing1.GetShapes()[0];
            Assert.IsTrue(Arrays.Equals(shape1.PictureData.Data, shape11.PictureData.Data));
            XSSFPicture shape22 = (XSSFPicture)drawing1.GetShapes()[1];
            Assert.IsTrue(Arrays.Equals(shape2.PictureData.Data, shape22.PictureData.Data));

            sheet2 = wb.GetSheetAt(1) as XSSFSheet;
            drawing2 = sheet2.CreateDrawingPatriarch() as XSSFDrawing;
            XSSFPicture shape33 = (XSSFPicture)drawing2.GetShapes()[0];
            Assert.IsTrue(Arrays.Equals(shape3.PictureData.Data, shape33.PictureData.Data));
            XSSFPicture shape44 = (XSSFPicture)drawing2.GetShapes()[1];
            Assert.IsTrue(Arrays.Equals(shape4.PictureData.Data, shape44.PictureData.Data));

        }
    }
}

