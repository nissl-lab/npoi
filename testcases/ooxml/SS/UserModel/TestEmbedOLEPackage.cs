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
using System.Linq;
using System.Text;

namespace TestCases.SS.UserModel
{
    using NPOI.HSSF;
    using NPOI.HSSF.UserModel;
    using NPOI.OpenXmlFormats.Vml;
    using NPOI.OpenXmlFormats.Vml.Spreadsheet;
    using NPOI.SS.Extractor;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

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
            ClassicAssert.IsTrue(enu.Current is IObjectData);

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
            ClassicAssert.IsTrue(enu.Current is IObjectData);

            EmbeddedExtractor ee = new EmbeddedExtractor();
            EmbeddedData ed = ee.ExtractAll(wb2.GetSheetAt(0))[0];
            CollectionAssert.AreEqual(samplePPT, ed.GetEmbeddedData());

            wb2.Close();
            wb1.Close();
        }

        [Test]
        public void EmbedXSSF_VMLShapeCreated()
        {
            IWorkbook wb1 = new XSSFWorkbook();
            ISheet sh = wb1.CreateSheet();
            int picIdx = wb1.AddPicture(GetSamplePng(), PictureType.PNG);
            byte[] samplePPTX = GetSamplePPT(true);
            int oleIdx = wb1.AddOlePackage(samplePPTX, "dummy.pptx", "dummy.pptx", "dummy.pptx");

            IDrawing<IShape> pat = sh.CreateDrawingPatriarch();
            IClientAnchor anchor = pat.CreateAnchor(0, 0, 0, 0, 2, 3, 5, 8);
            pat.CreateObjectData(anchor, oleIdx, picIdx);

            IWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);

            // Verify VML drawing contains an OLE shape with proper metadata
            XSSFSheet xssfSheet = (XSSFSheet)wb2.GetSheetAt(0);
            XSSFVMLDrawing vml = xssfSheet.GetVMLDrawing(false);
            ClassicAssert.IsNotNull(vml, "VML drawing should exist after CreateObjectData");

            CT_ClientData oleClientData = null;
            foreach (object item in vml.GetItems())
            {
                if (item is CT_Shape shape && shape.sizeOfClientDataArray() > 0)
                {
                    CT_ClientData cd = shape.GetClientDataArray(0);
                    if (cd.ObjectType == ST_ObjectType.Pict)
                    {
                        oleClientData = cd;
                        break;
                    }
                }
            }
            ClassicAssert.IsNotNull(oleClientData, "VML must contain a Pict ClientData for the OLE object");
            ClassicAssert.AreEqual("=EMBED(\"Packager Shell Object\",\"\")", oleClientData.FmlaPict);
            ClassicAssert.IsNotNull(oleClientData.anchor, "VML ClientData should have an Anchor");

            // Embedded data must survive round-trip
            EmbeddedExtractor ee = new EmbeddedExtractor();
            EmbeddedData ed = ee.ExtractAll(wb2.GetSheetAt(0))[0];
            CollectionAssert.AreEqual(samplePPTX, ed.GetEmbeddedData());

            wb2.Close();
            wb1.Close();
        }

        [Test]
        public void EmbedXSSF_MultipleOleObjects()
        {
            IWorkbook wb1 = new XSSFWorkbook();
            ISheet sh = wb1.CreateSheet();
            int picIdx = wb1.AddPicture(GetSamplePng(), PictureType.PNG);

            byte[] data1 = Encoding.UTF8.GetBytes("File one content");
            byte[] data2 = Encoding.UTF8.GetBytes("File two content");
            byte[] data3 = Encoding.UTF8.GetBytes("File three content");
            int oleIdx1 = wb1.AddOlePackage(data1, "one.txt", "one.txt", "one.txt");
            int oleIdx2 = wb1.AddOlePackage(data2, "two.txt", "two.txt", "two.txt");
            int oleIdx3 = wb1.AddOlePackage(data3, "three.txt", "three.txt", "three.txt");

            IDrawing<IShape> pat = sh.CreateDrawingPatriarch();
            pat.CreateObjectData(pat.CreateAnchor(0, 0, 0, 0, 0, 1, 2, 4), oleIdx1, picIdx);
            pat.CreateObjectData(pat.CreateAnchor(0, 0, 0, 0, 3, 1, 5, 4), oleIdx2, picIdx);
            pat.CreateObjectData(pat.CreateAnchor(0, 0, 0, 0, 6, 1, 8, 4), oleIdx3, picIdx);

            IWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);

            // All three files extractable
            EmbeddedExtractor ee = new EmbeddedExtractor();
            IList<EmbeddedData> all = ee.ExtractAll(wb2.GetSheetAt(0));
            ClassicAssert.AreEqual(3, all.Count, "Should have 3 embedded objects");
            foreach (EmbeddedData ed in all)
                ClassicAssert.IsNotNull(ed.GetEmbeddedData(), "Each embedded object should have data");

            // VML has 3 OLE shapes
            XSSFSheet xssfSheet = (XSSFSheet)wb2.GetSheetAt(0);
            int oleShapeCount = 0;
            foreach (object item in xssfSheet.GetVMLDrawing(false).GetItems())
            {
                if (item is CT_Shape shape && shape.sizeOfClientDataArray() > 0
                    && shape.GetClientDataArray(0).ObjectType == ST_ObjectType.Pict)
                    oleShapeCount++;
            }
            ClassicAssert.AreEqual(3, oleShapeCount, "VML should contain 3 OLE shapes");

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

