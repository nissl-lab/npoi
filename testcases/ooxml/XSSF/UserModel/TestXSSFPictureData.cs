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

using NUnit.Framework;
using System.Collections.Generic;
using System;
using NPOI.Util;
using NPOI.SS.UserModel;
using System.Collections;
using System.Text;
namespace NPOI.XSSF.UserModel
{
    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFPictureData
    {
        [Test]
        public void TestRead()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithDrawing.xlsx");
            IList pictures = wb.GetAllPictures();
            //wb.GetAllPictures() should return the same instance across multiple calls
            Assert.AreSame(pictures, wb.GetAllPictures());

            Assert.AreEqual(5, pictures.Count);
            String[] ext = { "jpeg", "emf", "png", "emf", "wmf" };
            String[] mimetype = { "image/jpeg", "image/x-emf", "image/png", "image/x-emf", "image/x-wmf" };
            for (int i = 0; i < pictures.Count; i++)
            {
                Assert.AreEqual(ext[i], ((XSSFPictureData)pictures[i]).SuggestFileExtension());
                Assert.AreEqual(mimetype[i], ((XSSFPictureData)pictures[i]).MimeType);
            }

            int num = pictures.Count;

            byte[] pictureData = { 0xA, 0xB, 0XC, 0xD, 0xE, 0xF };

            int idx = wb.AddPicture(pictureData, PictureType.JPEG);
            Assert.AreEqual(num + 1, pictures.Count);
            //idx is 0-based index in the #pictures array
            Assert.AreEqual(pictures.Count - 1, idx);
            XSSFPictureData pict = (XSSFPictureData)pictures[idx];
            Assert.AreEqual("jpeg", pict.SuggestFileExtension());
            Assert.IsTrue(Arrays.Equals(pictureData, pict.Data));
        }
        [Test]
        public void TestNew()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing Drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();

            byte[] jpegData = Encoding.UTF8.GetBytes("test jpeg data");
            byte[] wmfData = Encoding.UTF8.GetBytes("test wmf data");
            byte[] pngData = Encoding.UTF8.GetBytes("test png data");

            IList pictures = wb.GetAllPictures();
            Assert.AreEqual(0, pictures.Count);

            int jpegIdx = wb.AddPicture(jpegData, PictureType.JPEG);
            Assert.AreEqual(1, pictures.Count);
            Assert.AreEqual("jpeg", ((XSSFPictureData)pictures[jpegIdx]).SuggestFileExtension());
            Assert.IsTrue(Arrays.Equals(jpegData, ((XSSFPictureData)pictures[jpegIdx]).Data));

            int wmfIdx = wb.AddPicture(wmfData, PictureType.WMF);
            Assert.AreEqual(2, pictures.Count);
            Assert.AreEqual("wmf", ((XSSFPictureData)pictures[wmfIdx]).SuggestFileExtension());
            Assert.IsTrue(Arrays.Equals(wmfData, ((XSSFPictureData)pictures[wmfIdx]).Data));

            int pngIdx = wb.AddPicture(pngData, PictureType.PNG);
            Assert.AreEqual(3, pictures.Count);
            Assert.AreEqual("png", ((XSSFPictureData)pictures[pngIdx]).SuggestFileExtension());
            Assert.IsTrue(Arrays.Equals(pngData, ((XSSFPictureData)pictures[pngIdx]).Data));

            //TODO finish usermodel API for XSSFPicture
            XSSFPicture p1 = (XSSFPicture)Drawing.CreatePicture(new XSSFClientAnchor(), jpegIdx);
            Assert.IsNotNull(p1);
            XSSFPicture p2 = (XSSFPicture)Drawing.CreatePicture(new XSSFClientAnchor(), wmfIdx);
            Assert.IsNotNull(p1);
            XSSFPicture p3 = (XSSFPicture)Drawing.CreatePicture(new XSSFClientAnchor(), pngIdx);
            Assert.IsNotNull(p1);

            //check that the Added pictures are accessible After write
            wb = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            IList pictures2 = wb.GetAllPictures();
            Assert.AreEqual(3, pictures2.Count);
            
            Assert.AreEqual("jpeg", ((XSSFPictureData)pictures2[jpegIdx]).SuggestFileExtension());
            Assert.IsTrue(Arrays.Equals(jpegData, ((XSSFPictureData)pictures2[jpegIdx]).Data));

            Assert.AreEqual("wmf", ((XSSFPictureData)pictures2[wmfIdx]).SuggestFileExtension());
            Assert.IsTrue(Arrays.Equals(wmfData, ((XSSFPictureData)pictures2[wmfIdx]).Data));

            Assert.AreEqual("png", ((XSSFPictureData)pictures2[pngIdx]).SuggestFileExtension());
            Assert.IsTrue(Arrays.Equals(pngData, ((XSSFPictureData)pictures2[pngIdx]).Data));

        }

        /**
         * Bug 53568:  XSSFPicture.PictureData can return null.
         */
        [Test]
        public void Test53568()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53568.xlsx");
            List<XSSFPictureData> pictures = wb.GetAllPictures() as List<XSSFPictureData>;
            Assert.IsNotNull(pictures);
            Assert.AreEqual(4, pictures.Count);

            XSSFSheet sheet1 = wb.GetSheetAt(0) as XSSFSheet;
            List<XSSFShape> shapes1 = (sheet1.CreateDrawingPatriarch() as XSSFDrawing).GetShapes();
            Assert.IsNotNull(shapes1);
            Assert.AreEqual(5, shapes1.Count);

            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                XSSFSheet sheet = wb.GetSheetAt(i) as XSSFSheet;
                XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
                foreach (XSSFShape shape in Drawing.GetShapes())
                {
                    if (shape is XSSFPicture)
                    {
                        XSSFPicture pic = (XSSFPicture)shape;
                        XSSFPictureData picData = pic.PictureData as XSSFPictureData;
                        Assert.IsNotNull(picData);
                    }
                }
            }

        }
    }


}