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

    using TestCases.HSSF;
    using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.POIFS.FileSystem;
using System.IO;
    using System.Text;
    using NPOI.Util;

    /**
     * 
     */
    [TestFixture]
    public class TestOLE2Embeding
    {
        [Test]
        public void TestEmbeding()
        {
            // This used to break, until bug #43116 was fixed
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("ole2-embedding.xls");

            // Check we can get at the Escher layer still
            workbook.GetAllPictures();
        }
        [Test]
        public void TestEmbeddedObjects()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("ole2-embedding.xls");

            IList<HSSFObjectData> objects = workbook.GetAllEmbeddedObjects();
            Assert.AreEqual(2, objects.Count, "Wrong number of objects");
            Assert.AreEqual("MBD06CAB431",
                ((HSSFObjectData)objects[0]).GetDirectory().Name,
                    "Wrong name for first object");
            Assert.AreEqual("MBD06CAC85A",
                    ((HSSFObjectData)
                    objects[1]).GetDirectory().Name, "Wrong name for second object");
        }

        [Test]
        public void TestReallyEmbedSomething()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch() as HSSFPatriarch;

            byte[] pictureData = HSSFTestDataSamples.GetTestDataFileContent("logoKarmokar4.png");
            byte[] picturePPT = POIDataSamples.GetSlideShowInstance().ReadFile("clock.jpg");
            int imgIdx = wb.AddPicture(pictureData, PictureType.PNG);
            POIFSFileSystem pptPoifs = GetSamplePPT();
            int pptIdx = wb.AddOlePackage(pptPoifs, "Sample-PPT", "sample.ppt", "sample.ppt");
            POIFSFileSystem xlsPoifs = GetSampleXLS();
            int imgPPT = wb.AddPicture(picturePPT, PictureType.JPEG);
            int xlsIdx = wb.AddOlePackage(xlsPoifs, "Sample-XLS", "sample.xls", "sample.xls");
            int txtIdx = wb.AddOlePackage(GetSampleTXT(), "Sample-TXT", "sample.txt", "sample.txt");

            int rowoffset = 5;
            int coloffset = 5;

            ICreationHelper ch = wb.GetCreationHelper();
            HSSFClientAnchor anchor = (HSSFClientAnchor)ch.CreateClientAnchor();
            anchor.SetAnchor((short)(2 + coloffset), 1 + rowoffset, 0, 0, (short)(3 + coloffset), 5 + rowoffset, 0, 0);
            anchor.AnchorType = (/*setter*/AnchorType.DontMoveAndResize);

            patriarch.CreateObjectData(anchor, pptIdx, imgPPT);

            anchor = (HSSFClientAnchor)ch.CreateClientAnchor();
            anchor.SetAnchor((short)(5 + coloffset), 1 + rowoffset, 0, 0, (short)(6 + coloffset), 5 + rowoffset, 0, 0);
            anchor.AnchorType = (/*setter*/AnchorType.DontMoveAndResize);

            patriarch.CreateObjectData(anchor, xlsIdx, imgIdx);

            anchor = (HSSFClientAnchor)ch.CreateClientAnchor();
            anchor.SetAnchor((short)(3 + coloffset), 10 + rowoffset, 0, 0, (short)(5 + coloffset), 11 + rowoffset, 0, 0);
            anchor.AnchorType = (/*setter*/AnchorType.DontMoveAndResize);

            patriarch.CreateObjectData(anchor, txtIdx, imgIdx);

            anchor = (HSSFClientAnchor)ch.CreateClientAnchor();
            anchor.SetAnchor((short)(1 + coloffset), -2 + rowoffset, 0, 0, (short)(7 + coloffset), 14 + rowoffset, 0, 0);
            anchor.AnchorType = (/*setter*/AnchorType.DontMoveAndResize);

            HSSFSimpleShape circle = patriarch.CreateSimpleShape(anchor);
            circle.ShapeType = (/*setter*/HSSFSimpleShape.OBJECT_TYPE_OVAL);
            circle.IsNoFill = (/*setter*/true);

            if (false)
            {
                FileStream fos = new FileStream("embed.xls", FileMode.Create);
                wb.Write(fos);
                fos.Close();
            }

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb as HSSFWorkbook);

            MemoryStream bos = new MemoryStream();
            HSSFObjectData od = wb.GetAllEmbeddedObjects()[0];
            Ole10Native ole10 = Ole10Native.CreateFromEmbeddedOleObject((DirectoryNode)od.GetDirectory());
            bos = new MemoryStream();
            pptPoifs.WriteFileSystem(bos);
            Assert.IsTrue(Arrays.Equals(ole10.DataBuffer, bos.ToArray()));

            od = wb.GetAllEmbeddedObjects()[1];
            ole10 = Ole10Native.CreateFromEmbeddedOleObject((DirectoryNode)od.GetDirectory());
            bos = new MemoryStream();
            xlsPoifs.WriteFileSystem(bos);
            Assert.IsTrue(Arrays.Equals(ole10.DataBuffer, bos.ToArray()));

            od = wb.GetAllEmbeddedObjects()[2];
            ole10 = Ole10Native.CreateFromEmbeddedOleObject((DirectoryNode)od.GetDirectory());
            Assert.IsTrue(Arrays.Equals(ole10.DataBuffer, GetSampleTXT()));

        }

        static POIFSFileSystem GetSamplePPT()
        {
            // scratchpad classes are not available, so we use something pre-cooked
            Stream is1 = POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("with_textbox.ppt");
            POIFSFileSystem poifs = new POIFSFileSystem(is1);
            is1.Close();

            return poifs;
        }

        static POIFSFileSystem GetSampleXLS()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            sheet.CreateRow(5).CreateCell(2).SetCellValue("yo dawg i herd you like embeddet objekts, so we Put a ole in your ole so you can save a file while you save a file");

            MemoryStream bos = new MemoryStream();
            wb.Write(bos);
            POIFSFileSystem poifs = new POIFSFileSystem(new MemoryStream(bos.ToArray()));

            return poifs;
        }

        static byte[] GetSampleTXT()
        {
            return Encoding.Default.GetBytes("All your base are belong to us");
        }
    }
}

