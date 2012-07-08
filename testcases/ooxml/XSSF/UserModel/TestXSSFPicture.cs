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
namespace NPOI.XSSF.UserModel
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
        public void TestResize()
        {
            BaseTestResize(new XSSFClientAnchor(0, 0, 504825, 85725, (short)0, 0, (short)1, 8));
        }

        [Test]
        public void TestCreate()
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
            Assert.AreEqual(AnchorType.MOVE_AND_RESIZE, (AnchorType)anchor.AnchorType);
            anchor.AnchorType = (int)AnchorType.DONT_MOVE_AND_RESIZE;
            Assert.AreEqual(AnchorType.DONT_MOVE_AND_RESIZE, (AnchorType)anchor.AnchorType);

            XSSFPicture shape = (XSSFPicture)drawing.CreatePicture(anchor, jpegIdx);
            Assert.IsTrue(anchor.Equals(shape.GetAnchor()));
            Assert.IsNotNull(shape.PictureData);
            Assert.IsTrue(Arrays.Equals(jpegData, shape.PictureData.Data));

            CT_TwoCellAnchor ctShapeHolder = drawing.GetCTDrawing().TwoCellAnchors[0];
            // STEditAs.ABSOLUTE corresponds to ClientAnchor.DONT_MOVE_AND_RESIZE
            Assert.AreEqual(ST_EditAs.absolute, ctShapeHolder.editAs);
        }

        /**
         * Test that ShapeId in CTNonVisualDrawingProps is incremented
         *
         * See Bugzilla 50458
         */
        [Test]
        public void TestShapeId()
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
    }
}

