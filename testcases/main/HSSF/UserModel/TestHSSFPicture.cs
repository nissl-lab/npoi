/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using NPOI.DDF;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NUnit.Framework;
    using TestCases.SS.UserModel;

    /**
     * Test <c>HSSFPicture</c>.
     *
     * @author Yegor Kozlov (yegor at apache.org)
     */
    [TestFixture]
    public class TestHSSFPicture : BaseTestPicture
    {
        public TestHSSFPicture()
            : base(HSSFITestDataProvider.Instance)
        {

        }

        [Test]
        public void Resize()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("resize_compare.xls");
            HSSFPatriarch dp = wb.GetSheetAt(0).CreateDrawingPatriarch() as HSSFPatriarch;
            IList<HSSFShape> pics = dp.Children;
            HSSFPicture inpPic = (HSSFPicture)pics[(0)];
            HSSFPicture cmpPic = (HSSFPicture)pics[(1)];

            BaseTestResize(inpPic, cmpPic, 2.0, 2.0);
            //wb.Close();
        }


        /**
         * Bug # 45829 reported ArithmeticException (/ by zero) when resizing png with zero DPI.
         */
        [Test]
        public void Bug45829()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sh1 = wb.CreateSheet();
            IDrawing p1 = sh1.CreateDrawingPatriarch();

            byte[] pictureData = HSSFTestDataSamples.GetTestDataFileContent("45829.png");
            int idx1 = wb.AddPicture(pictureData, PictureType.PNG);
            IPicture pic = p1.CreatePicture(new HSSFClientAnchor(), idx1);
            pic.Resize();
        }
        [Test]
        public void AddPictures()
        {
            IWorkbook wb = new HSSFWorkbook();

            ISheet sh = wb.CreateSheet("Pictures");
            IDrawing dr = sh.CreateDrawingPatriarch();
            Assert.AreEqual(0, ((HSSFPatriarch)dr).Children.Count);
            IClientAnchor anchor = wb.GetCreationHelper().CreateClientAnchor();

            //register a picture
            byte[] data1 = new byte[] { 1, 2, 3 };
            int idx1 = wb.AddPicture(data1, PictureType.JPEG);
            Assert.AreEqual(1, idx1);
            IPicture p1 = dr.CreatePicture(anchor, idx1);
            Assert.IsTrue(Arrays.Equals(data1, ((HSSFPicture)p1).PictureData.Data));

            // register another one
            byte[] data2 = new byte[] { 4, 5, 6 };
            int idx2 = wb.AddPicture(data2, PictureType.JPEG);
            Assert.AreEqual(2, idx2);
            IPicture p2 = dr.CreatePicture(anchor, idx2);
            Assert.AreEqual(2, ((HSSFPatriarch)dr).Children.Count);
            Assert.IsTrue(Arrays.Equals(data2, ((HSSFPicture)p2).PictureData.Data));

            // confirm that HSSFPatriarch.Children returns two picture shapes 
            Assert.IsTrue(Arrays.Equals(data1, ((HSSFPicture)((HSSFPatriarch)dr).Children[(0)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data2, ((HSSFPicture)((HSSFPatriarch)dr).Children[(1)]).PictureData.Data));

            // Write, read back and verify that our pictures are there
            wb = HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)wb);
            IList lst2 = wb.GetAllPictures();
            Assert.AreEqual(2, lst2.Count);
            Assert.IsTrue(Arrays.Equals(data1, (lst2[(0)] as HSSFPictureData).Data));
            Assert.IsTrue(Arrays.Equals(data2, (lst2[(1)] as HSSFPictureData).Data));

            // confirm that the pictures are in the Sheet's Drawing
            sh = wb.GetSheet("Pictures");
            dr = sh.CreateDrawingPatriarch();
            Assert.AreEqual(2, ((HSSFPatriarch)dr).Children.Count);
            Assert.IsTrue(Arrays.Equals(data1, ((HSSFPicture)((HSSFPatriarch)dr).Children[(0)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data2, ((HSSFPicture)((HSSFPatriarch)dr).Children[(1)]).PictureData.Data));

            // add a third picture
            byte[] data3 = new byte[] { 7, 8, 9 };
            // picture index must increment across Write-read
            int idx3 = wb.AddPicture(data3, PictureType.JPEG);
            Assert.AreEqual(3, idx3);
            IPicture p3 = dr.CreatePicture(anchor, idx3);
            Assert.IsTrue(Arrays.Equals(data3, ((HSSFPicture)p3).PictureData.Data));
            Assert.AreEqual(3, ((HSSFPatriarch)dr).Children.Count);
            Assert.IsTrue(Arrays.Equals(data1, ((HSSFPicture)((HSSFPatriarch)dr).Children[(0)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data2, ((HSSFPicture)((HSSFPatriarch)dr).Children[(1)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data3, ((HSSFPicture)((HSSFPatriarch)dr).Children[(2)]).PictureData.Data));

            // write and read again
            wb = HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)wb);
            IList lst3 = wb.GetAllPictures();
            // all three should be there
            Assert.AreEqual(3, lst3.Count);
            Assert.IsTrue(Arrays.Equals(data1, (lst3[(0)] as HSSFPictureData).Data));
            Assert.IsTrue(Arrays.Equals(data2, (lst3[(1)] as HSSFPictureData).Data));
            Assert.IsTrue(Arrays.Equals(data3, (lst3[(2)] as HSSFPictureData).Data));

            sh = wb.GetSheet("Pictures");
            dr = sh.CreateDrawingPatriarch();
            Assert.AreEqual(3, ((HSSFPatriarch)dr).Children.Count);

            // forth picture
            byte[] data4 = new byte[] { 10, 11, 12 };
            int idx4 = wb.AddPicture(data4, PictureType.JPEG);
            Assert.AreEqual(4, idx4);
            dr.CreatePicture(anchor, idx4);
            Assert.AreEqual(4, ((HSSFPatriarch)dr).Children.Count);
            Assert.IsTrue(Arrays.Equals(data1, ((HSSFPicture)((HSSFPatriarch)dr).Children[(0)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data2, ((HSSFPicture)((HSSFPatriarch)dr).Children[(1)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data3, ((HSSFPicture)((HSSFPatriarch)dr).Children[(2)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data4, ((HSSFPicture)((HSSFPatriarch)dr).Children[(3)]).PictureData.Data));

            wb = HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)wb);
            IList lst4 = wb.GetAllPictures();
            Assert.AreEqual(4, lst4.Count);
            Assert.IsTrue(Arrays.Equals(data1, (lst4[(0)] as HSSFPictureData).Data));
            Assert.IsTrue(Arrays.Equals(data2, (lst4[(1)] as HSSFPictureData).Data));
            Assert.IsTrue(Arrays.Equals(data3, (lst4[(2)] as HSSFPictureData).Data));
            Assert.IsTrue(Arrays.Equals(data4, (lst4[(3)] as HSSFPictureData).Data));
            sh = wb.GetSheet("Pictures");
            dr = sh.CreateDrawingPatriarch();
            Assert.AreEqual(4, ((HSSFPatriarch)dr).Children.Count);
            Assert.IsTrue(Arrays.Equals(data1, ((HSSFPicture)((HSSFPatriarch)dr).Children[(0)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data2, ((HSSFPicture)((HSSFPatriarch)dr).Children[(1)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data3, ((HSSFPicture)((HSSFPatriarch)dr).Children[(2)]).PictureData.Data));
            Assert.IsTrue(Arrays.Equals(data4, ((HSSFPicture)((HSSFPatriarch)dr).Children[(3)]).PictureData.Data));
        }
        [Test]
        public void BSEPictureRef()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            HSSFSheet sh = wb.CreateSheet("Pictures") as HSSFSheet;
            HSSFPatriarch dr = sh.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFClientAnchor anchor = new HSSFClientAnchor();

            InternalSheet ish = HSSFTestHelper.GetSheetForTest(sh);

            //register a picture
            byte[] data1 = new byte[] { 1, 2, 3 };
            int idx1 = wb.AddPicture(data1, PictureType.JPEG);
            Assert.AreEqual(1, idx1);
            HSSFPicture p1 = dr.CreatePicture(anchor, idx1) as HSSFPicture;

            EscherBSERecord bse = wb.Workbook.GetBSERecord(idx1);

            Assert.AreEqual(bse.Ref, 1);
            dr.CreatePicture(new HSSFClientAnchor(), idx1);
            Assert.AreEqual(bse.Ref, 2);

            HSSFShapeGroup gr = dr.CreateGroup(new HSSFClientAnchor());
            gr.CreatePicture(new HSSFChildAnchor(), idx1);
            Assert.AreEqual(bse.Ref, 3);
        }
        [Test]
        public void ReadExistingImage()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("picture") as HSSFSheet;
            HSSFPatriarch Drawing = sheet.DrawingPatriarch as HSSFPatriarch;
            Assert.AreEqual(1, Drawing.Children.Count);

            HSSFPicture picture = (HSSFPicture)Drawing.Children[0];
            Assert.AreEqual(picture.FileName, "test");
        }
        [Test]
        public void SetGetProperties()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            HSSFSheet sh = wb.CreateSheet("Pictures") as HSSFSheet;
            HSSFPatriarch dr = sh.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFClientAnchor anchor = new HSSFClientAnchor();

            //register a picture
            byte[] data1 = new byte[] { 1, 2, 3 };
            int idx1 = wb.AddPicture(data1, PictureType.JPEG);
            HSSFPicture p1 = dr.CreatePicture(anchor, idx1) as HSSFPicture;

            Assert.AreEqual(p1.FileName, "");
            p1.FileName = ("aaa");
            Assert.AreEqual(p1.FileName, "aaa");

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheet("Pictures") as HSSFSheet;
            dr = sh.DrawingPatriarch as HSSFPatriarch;

            p1 = (HSSFPicture)dr.Children[0];
            Assert.AreEqual(p1.FileName, "aaa");
        }

        [Test]
        public void Bug49658()
        {
            // test if inserted EscherMetafileBlip will be read again
            IWorkbook wb = new HSSFWorkbook();

            byte[] pictureDataEmf = POIDataSamples.GetDocumentInstance().ReadFile("vector_image.emf");
            int indexEmf = wb.AddPicture(pictureDataEmf, PictureType.EMF);
            byte[] pictureDataPng = POIDataSamples.GetSpreadSheetInstance().ReadFile("logoKarmokar4.png");
            int indexPng = wb.AddPicture(pictureDataPng, PictureType.PNG);
            byte[] pictureDataWmf = POIDataSamples.GetSlideShowInstance().ReadFile("santa.wmf");
            int indexWmf = wb.AddPicture(pictureDataWmf, PictureType.WMF);

            ISheet sheet = wb.CreateSheet();
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
            ICreationHelper ch = wb.GetCreationHelper();

            IClientAnchor anchor = ch.CreateClientAnchor();
            anchor.Col1 = (/*setter*/2);
            anchor.Col2 = (/*setter*/5);
            anchor.Row1 = (/*setter*/1);
            anchor.Row2 = (/*setter*/6);
            patriarch.CreatePicture(anchor, indexEmf);

            anchor = ch.CreateClientAnchor();
            anchor.Col1 = (/*setter*/2);
            anchor.Col2 = (/*setter*/5);
            anchor.Row1 = (/*setter*/10);
            anchor.Row2 = (/*setter*/16);
            patriarch.CreatePicture(anchor, indexPng);

            anchor = ch.CreateClientAnchor();
            anchor.Col1 = (/*setter*/6);
            anchor.Col2 = (/*setter*/9);
            anchor.Row1 = (/*setter*/1);
            anchor.Row2 = (/*setter*/6);
            patriarch.CreatePicture(anchor, indexWmf);


            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb as HSSFWorkbook);
            byte[] pictureDataOut = (wb.GetAllPictures()[0] as HSSFPictureData).Data;
            Assert.IsTrue(Arrays.Equals(pictureDataEmf, pictureDataOut));

            byte[] wmfNoHeader = new byte[pictureDataWmf.Length - 22];
            Array.Copy(pictureDataWmf, 22, wmfNoHeader, 0, pictureDataWmf.Length - 22);
            pictureDataOut = (wb.GetAllPictures()[2] as HSSFPictureData).Data;
            Assert.IsTrue(Arrays.Equals(wmfNoHeader, pictureDataOut));
        }

    }
}