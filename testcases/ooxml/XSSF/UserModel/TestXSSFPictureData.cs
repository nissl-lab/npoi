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

using NUnit.Framework;using NUnit.Framework.Legacy;
using System.Collections.Generic;
using System;
using NPOI.Util;
using NPOI.SS.UserModel;
using System.Collections;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;

namespace TestCases.XSSF.UserModel
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
            ClassicAssert.AreSame(pictures, wb.GetAllPictures());

            ClassicAssert.AreEqual(5, pictures.Count);
            String[] ext = { "jpeg", "emf", "png", "emf", "wmf" };
            String[] mimetype = { "image/jpeg", "image/x-emf", "image/png", "image/x-emf", "image/x-wmf" };
            for (int i = 0; i < pictures.Count; i++)
            {
                ClassicAssert.AreEqual(ext[i], ((XSSFPictureData)pictures[i]).SuggestFileExtension());
                ClassicAssert.AreEqual(mimetype[i], ((XSSFPictureData)pictures[i]).MimeType);
            }

            int num = pictures.Count;

            byte[] pictureData = { 0xA, 0xB, 0XC, 0xD, 0xE, 0xF };

            int idx = wb.AddPicture(pictureData, PictureType.JPEG);
            ClassicAssert.AreEqual(num + 1, pictures.Count);
            //idx is 0-based index in the #pictures array
            ClassicAssert.AreEqual(pictures.Count - 1, idx);
            XSSFPictureData pict = (XSSFPictureData)pictures[idx];
            ClassicAssert.AreEqual("jpeg", pict.SuggestFileExtension());
            ClassicAssert.IsTrue(Arrays.Equals(pictureData, pict.Data));
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
            ClassicAssert.AreEqual(0, pictures.Count);

            int jpegIdx = wb.AddPicture(jpegData, PictureType.JPEG);
            ClassicAssert.AreEqual(1, pictures.Count);
            ClassicAssert.AreEqual("jpeg", ((XSSFPictureData)pictures[jpegIdx]).SuggestFileExtension());
            ClassicAssert.IsTrue(Arrays.Equals(jpegData, ((XSSFPictureData)pictures[jpegIdx]).Data));

            int wmfIdx = wb.AddPicture(wmfData, PictureType.WMF);
            ClassicAssert.AreEqual(2, pictures.Count);
            ClassicAssert.AreEqual("wmf", ((XSSFPictureData)pictures[wmfIdx]).SuggestFileExtension());
            ClassicAssert.IsTrue(Arrays.Equals(wmfData, ((XSSFPictureData)pictures[wmfIdx]).Data));

            int pngIdx = wb.AddPicture(pngData, PictureType.PNG);
            ClassicAssert.AreEqual(3, pictures.Count);
            ClassicAssert.AreEqual("png", ((XSSFPictureData)pictures[pngIdx]).SuggestFileExtension());
            ClassicAssert.IsTrue(Arrays.Equals(pngData, ((XSSFPictureData)pictures[pngIdx]).Data));

            //TODO finish usermodel API for XSSFPicture
            XSSFPicture p1 = (XSSFPicture)Drawing.CreatePicture(new XSSFClientAnchor(), jpegIdx);
            ClassicAssert.IsNotNull(p1);
            XSSFPicture p2 = (XSSFPicture)Drawing.CreatePicture(new XSSFClientAnchor(), wmfIdx);
            ClassicAssert.IsNotNull(p1);
            XSSFPicture p3 = (XSSFPicture)Drawing.CreatePicture(new XSSFClientAnchor(), pngIdx);
            ClassicAssert.IsNotNull(p1);

            //check that the Added pictures are accessible After write
            wb = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            IList pictures2 = wb.GetAllPictures();
            ClassicAssert.AreEqual(3, pictures2.Count);
            
            ClassicAssert.AreEqual("jpeg", ((XSSFPictureData)pictures2[jpegIdx]).SuggestFileExtension());
            ClassicAssert.IsTrue(Arrays.Equals(jpegData, ((XSSFPictureData)pictures2[jpegIdx]).Data));

            ClassicAssert.AreEqual("wmf", ((XSSFPictureData)pictures2[wmfIdx]).SuggestFileExtension());
            ClassicAssert.IsTrue(Arrays.Equals(wmfData, ((XSSFPictureData)pictures2[wmfIdx]).Data));

            ClassicAssert.AreEqual("png", ((XSSFPictureData)pictures2[pngIdx]).SuggestFileExtension());
            ClassicAssert.IsTrue(Arrays.Equals(pngData, ((XSSFPictureData)pictures2[pngIdx]).Data));

        }

        /**
         * Bug 53568:  XSSFPicture.PictureData can return null.
         */
        [Test]
        public void Test53568()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53568.xlsx");
            List<XSSFPictureData> pictures = wb.GetAllPictures() as List<XSSFPictureData>;
            ClassicAssert.IsNotNull(pictures);
            ClassicAssert.AreEqual(4, pictures.Count);

            XSSFSheet sheet1 = wb.GetSheetAt(0) as XSSFSheet;
            List<XSSFShape> shapes1 = (sheet1.CreateDrawingPatriarch() as XSSFDrawing).GetShapes();
            ClassicAssert.IsNotNull(shapes1);
            ClassicAssert.AreEqual(5, shapes1.Count);

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
                        ClassicAssert.IsNotNull(picData);
                    }
                }
            }

        }
    }


}