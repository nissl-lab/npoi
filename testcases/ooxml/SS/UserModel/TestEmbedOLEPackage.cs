/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.UserModel
{
    using NPOI.HSSF;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Extractor;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;

    [TestFixture]
    public class TestEmbedOLEPackage
    {
        [Test]
        public void EmbedXSSF()
        {

            IWorkbook wb1 = new XSSFWorkbook();
            ISheet sh = wb1.CreateSheet();
            int picIdx = wb1.AddPicture(GetSamplePng(), PictureType.PNG);
            byte[] samplePPTX = GetSamplePPT(true);
            int oleIdx = wb1.AddOlePackage(samplePPTX, "dummy.pptx", "dummy.pptx", "dummy.pptx");

            IDrawing<IShape> pat = sh.CreateDrawingPatriarch();
            IClientAnchor anchor = pat.CreateAnchor(0, 0, 0, 0, 1, 1, 3, 6);
            pat.CreateObjectData(anchor, oleIdx, picIdx);

            IWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);

            pat = wb2.GetSheetAt(0).DrawingPatriarch;
            var enu = pat.GetEnumerator();
            enu.MoveNext();
            Assert.IsTrue(enu.Current is IObjectData);

            EmbeddedExtractor ee = new EmbeddedExtractor();
            EmbeddedData ed = ee.ExtractAll(wb2.GetSheetAt(0))[0];

            CollectionAssert.AreEqual(samplePPTX, ed.GetEmbeddedData());

            wb2.Close();
            wb1.Close();
        }

        [Test]
        public void EmbedHSSF()
        {

            //try
            //{
            //    Class.forName("NPOI.HSLF.UserModel.HSLFSlideShow");
            //}
            //catch(Exception e)
            //{
            //    assumeTrue(false);
            //}

            IWorkbook wb1 = new HSSFWorkbook();
            ISheet sh = wb1.CreateSheet();
            int picIdx = wb1.AddPicture(GetSamplePng(), PictureType.PNG);
            byte[] samplePPT = GetSamplePPT(false);
            int oleIdx = wb1.AddOlePackage(samplePPT, "dummy.ppt", "dummy.ppt", "dummy.ppt");

            IDrawing<IShape> pat = sh.CreateDrawingPatriarch();
            IClientAnchor anchor = pat.CreateAnchor(0, 0, 0, 0, 1, 1, 3, 6);
            pat.CreateObjectData(anchor, oleIdx, picIdx);

            IWorkbook wb2 = HSSF.HSSFTestDataSamples.WriteOutAndReadBack((HSSFWorkbook)wb1);

            pat = wb2.GetSheetAt(0).DrawingPatriarch;
            var enu = pat.GetEnumerator();
            enu.MoveNext();
            Assert.IsTrue(enu.Current is IObjectData);

            EmbeddedExtractor ee = new EmbeddedExtractor();
            EmbeddedData ed = ee.ExtractAll(wb2.GetSheetAt(0))[0];
            CollectionAssert.AreEqual(samplePPT, ed.GetEmbeddedData());

            wb2.Close();
            wb1.Close();
        }

        static byte[] GetSamplePng()
        {
            var provider = XSSFITestDataProvider.instance;
            return provider.GetTestDataFileContent("logoKarmokar4.png");
        }

        static byte[] GetSamplePPT(bool ooxml)
        {
            var provider = POIDataSamples.GetSlideShowInstance();
            string filename = ooxml ? "49386-null_dates.pptx":"41071.ppt";
            return provider.ReadFile(filename);
            //SlideShow<?,?> ppt = (ooxml) ? new XMLSlideShow() : new NPOI.HSLF.UserModel.HSLFSlideShow();
            //Slide<?,?> slide = ppt.CreateSlide();

            //AutoShape<?,?> sh1 = slide.CreateAutoShape();
            //sh1.SetShapeType(ShapeType.STAR_32);
            //sh1.SetAnchor(new java.awt.Rectangle(50, 50, 100, 200));
            //sh1.SetFillColor(java.awt.Color.red);

            //MemoryStream bos = new MemoryStream();
            //ppt.write(bos);
            //ppt.Close();

            //return bos.ToArray();
        }
    }
}

