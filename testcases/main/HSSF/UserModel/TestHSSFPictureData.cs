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

namespace TestCases.HSSF.UserModel
{

    using System;
    using System.Collections;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    using System.IO;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;
    using SixLabors.ImageSharp;


    /**
     * Test <c>HSSFPictureData</c>.
     * The code to retrieve images from a workbook provided by Trejkaz (trejkaz at trypticon dot org) in Bug 41223.
     *
     * @author Yegor Kozlov (yegor at apache dot org)
     * @author Trejkaz (trejkaz at trypticon dot org)
     */
    [TestFixture]
    public class TestHSSFPictureData
    {

        [Test]
        public void TestPictures()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithImages.xls");

            IList lst = wb.GetAllPictures();
            //Assert.AreEqual(2, lst.Count);

            for (IEnumerator it = lst.GetEnumerator(); it.MoveNext(); )
            {
                HSSFPictureData pict = (HSSFPictureData)it.Current;
                String ext = pict.SuggestFileExtension();
                byte[] data = pict.Data;
                if (ext.Equals("jpeg"))
                {
                    //try to read image data using javax.imageio.* (JDK 1.4+)
                    Image jpg = Image.Load(new MemoryStream(data));
                    Assert.IsNotNull(jpg);
                    Assert.AreEqual(192, jpg.Width);
                    Assert.AreEqual(176, jpg.Height);
                }
                else if (ext.Equals("png"))
                {
                    //try to read image data using javax.imageio.* (JDK 1.4+)
                    Image png = Image.Load(new MemoryStream(data));
                    Assert.IsNotNull(png);
                    Assert.AreEqual(300, png.Width);
                    Assert.AreEqual(300, png.Height);
                }
                else
                {
                    //TODO: Test code for PICT, WMF and EMF
                }
            }
        }
        [Test]
        public void TestMacPicture()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("53446.xls");

            List<HSSFPictureData> lst = (List<HSSFPictureData>)wb.GetAllPictures();
            Assert.AreEqual(1, lst.Count);

            HSSFPictureData pict = lst[(0)];
            String ext = pict.SuggestFileExtension();
            if (!ext.Equals("png"))
            {
                Assert.Fail("Expected a PNG.");
            }

            //try to read image data using javax.imageio.* (JDK 1.4+)
            byte[] data = pict.Data;
            using (MemoryStream ms = new MemoryStream(data))
            {
                Image png = Image.Load(ms);
                Assert.IsNotNull(png);
                Assert.AreEqual(78, png.Width);
                Assert.AreEqual(76, png.Height);
                Assert.AreEqual((int)PictureType.PNG, pict.Format);
                Assert.AreEqual("image/png", pict.MimeType);
            }
        }

        [Test]
        public void TestNotNullPictures()
        {

            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SheetWithDrawing.xls");

            // TODO - add GetFormat() to interface PictureData and genericise wb.AllPictures
            List<HSSFPictureData> lst = (List<HSSFPictureData>)wb.GetAllPictures();
            foreach (HSSFPictureData pict in lst)
            {
                Assert.IsNotNull(pict);
            }
        }
    }
}