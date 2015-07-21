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

namespace NPOI.XWPF.UserModel
{
    using System;
    using NUnit.Framework;
    using NPOI.OpenXml4Net.OPC;
    using System.IO;
    using System.Collections.Generic;
    using NPOI.Util;
    using System.Xml.Serialization;
    using NPOI.OpenXmlFormats.Wordprocessing;

    [TestFixture]
    public class TestXWPFDocument
    {

        [Test]
        public void TestContainsMainContentType()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            OPCPackage pack = doc.Package;

            bool found = false;
            foreach (PackagePart part in pack.GetParts())
            {
                if (part.ContentType.Equals(XWPFRelation.DOCUMENT.ContentType))
                {
                    found = true;
                }
                if (false == found)
                {
                    // successful tests should be silent
                    System.Console.WriteLine(part);
                }
            }
            Assert.IsTrue(found);
        }

        [Test]
        public void TestOpen()
        {
            XWPFDocument xml;

            // Simple file
            xml = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            // Check it has key parts
            Assert.IsNotNull(xml.Document);
            Assert.IsNotNull(xml.Document.body);
            Assert.IsNotNull(xml.GetStyles());

            // Complex file
            xml = XWPFTestDataSamples.OpenSampleDocument("IllustrativeCases.docx");
            Assert.IsNotNull(xml.Document);
            Assert.IsNotNull(xml.Document.body);
            Assert.IsNotNull(xml.GetStyles());
        }

        [Test]
        public void TestMetadataBasics()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            Assert.IsNotNull(xml.GetProperties().CoreProperties);
            Assert.IsNotNull(xml.GetProperties().ExtendedProperties);

            Assert.AreEqual("Microsoft Office Word", xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Application);
            Assert.AreEqual(1315, xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Characters);
            Assert.AreEqual(10, xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Lines);

            Assert.AreEqual(null, xml.GetProperties().CoreProperties.Title);
            Assert.AreEqual(null, xml.GetProperties().CoreProperties.GetUnderlyingProperties().GetSubjectProperty());
        }

        [Test]
        public void TestMetadataComplex()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("IllustrativeCases.docx");
            Assert.IsNotNull(xml.GetProperties().CoreProperties);
            Assert.IsNotNull(xml.GetProperties().ExtendedProperties);

            Assert.AreEqual("Microsoft Office Outlook", xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Application);
            Assert.AreEqual(5184, xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Characters);
            Assert.AreEqual(0, xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Lines);

            Assert.AreEqual(" ", xml.GetProperties().CoreProperties.Title);
            Assert.AreEqual(" ", xml.GetProperties().CoreProperties.GetUnderlyingProperties().GetSubjectProperty());
        }

        [Test]
        public void TestWorkbookProperties()
        {
            XWPFDocument doc = new XWPFDocument();
            POIXMLProperties props = doc.GetProperties();
            Assert.IsNotNull(props);
            Assert.AreEqual("NPOI", props.ExtendedProperties.GetUnderlyingProperties().Application);
        }

        [Test]
        public void TestAddParagraph()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            Assert.AreEqual(3, doc.Paragraphs.Count);

            XWPFParagraph p = doc.CreateParagraph();
            Assert.AreEqual(p, doc.Paragraphs[(3)]);
            Assert.AreEqual(4, doc.Paragraphs.Count);

            Assert.AreEqual(3, doc.GetParagraphPos(3));
            Assert.AreEqual(3, doc.GetPosOfParagraph(p));

            //CTP ctp = p.CTP;
            //XWPFParagraph newP = doc.GetParagraph(ctp);
            //Assert.AreSame(p, newP);
            //XmlCursor cursor = doc.Document.Body.GetPArray(0).newCursor();
            //XWPFParagraph cP = doc.InsertNewParagraph(cursor);
            //Assert.AreSame(cP, doc.Paragraphs[(0)]);
            //Assert.AreEqual(5, doc.Paragraphs.Count);
        }
        [Test]
        public void TestAddPicture()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            byte[] jpeg = XWPFTestDataSamples.GetImage("nature1.jpg");
            String relationId = doc.AddPictureData(jpeg, (int)PictureType.JPEG);

            byte[] newJpeg = ((XWPFPictureData)doc.GetRelationById(relationId)).Data;
            Assert.AreEqual(newJpeg.Length, jpeg.Length);
            for (int i = 0; i < jpeg.Length; i++)
            {
                Assert.AreEqual(newJpeg[i], jpeg[i]);
            }
        }
        [Test]
        public void TestAllPictureFormats()
        {
            XWPFDocument doc = new XWPFDocument();

            doc.AddPictureData(new byte[10], (int)PictureType.EMF);
            doc.AddPictureData(new byte[11], (int)PictureType.WMF);
            doc.AddPictureData(new byte[12], (int)PictureType.PICT);
            doc.AddPictureData(new byte[13], (int)PictureType.JPEG);
            doc.AddPictureData(new byte[14], (int)PictureType.PNG);
            doc.AddPictureData(new byte[15], (int)PictureType.DIB);
            doc.AddPictureData(new byte[16], (int)PictureType.GIF);
            doc.AddPictureData(new byte[17], (int)PictureType.TIFF);
            doc.AddPictureData(new byte[18], (int)PictureType.EPS);
            doc.AddPictureData(new byte[19], (int)PictureType.BMP);
            doc.AddPictureData(new byte[20], (int)PictureType.WPG);

            Assert.AreEqual(11, doc.AllPictures.Count);

            doc = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            Assert.AreEqual(11, doc.AllPictures.Count);

        }
        [Test]
        public void TestRemoveBodyElement()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            Assert.AreEqual(3, doc.Paragraphs.Count);
            Assert.AreEqual(3, doc.BodyElements.Count);

            XWPFParagraph p1 = doc.Paragraphs[(0)];
            XWPFParagraph p2 = doc.Paragraphs[(1)];
            XWPFParagraph p3 = doc.Paragraphs[(2)];

            Assert.AreEqual(p1, doc.BodyElements[(0)]);
            Assert.AreEqual(p1, doc.Paragraphs[(0)]);
            Assert.AreEqual(p2, doc.BodyElements[(1)]);
            Assert.AreEqual(p2, doc.Paragraphs[(1)]);
            Assert.AreEqual(p3, doc.BodyElements[(2)]);
            Assert.AreEqual(p3, doc.Paragraphs[(2)]);

            // Add another
            XWPFParagraph p4 = doc.CreateParagraph();

            Assert.AreEqual(4, doc.Paragraphs.Count);
            Assert.AreEqual(4, doc.BodyElements.Count);
            Assert.AreEqual(p1, doc.BodyElements[(0)]);
            Assert.AreEqual(p1, doc.Paragraphs[(0)]);
            Assert.AreEqual(p2, doc.BodyElements[(1)]);
            Assert.AreEqual(p2, doc.Paragraphs[(1)]);
            Assert.AreEqual(p3, doc.BodyElements[(2)]);
            Assert.AreEqual(p3, doc.Paragraphs[(2)]);
            Assert.AreEqual(p4, doc.BodyElements[(3)]);
            Assert.AreEqual(p4, doc.Paragraphs[(3)]);

            // Remove the 2nd
            Assert.AreEqual(true, doc.RemoveBodyElement(1));
            Assert.AreEqual(3, doc.Paragraphs.Count);
            Assert.AreEqual(3, doc.BodyElements.Count);

            Assert.AreEqual(p1, doc.BodyElements[(0)]);
            Assert.AreEqual(p1, doc.Paragraphs[(0)]);
            Assert.AreEqual(p3, doc.BodyElements[(1)]);
            Assert.AreEqual(p3, doc.Paragraphs[(1)]);
            Assert.AreEqual(p4, doc.BodyElements[(2)]);
            Assert.AreEqual(p4, doc.Paragraphs[(2)]);

            // Remove the 1st
            Assert.AreEqual(true, doc.RemoveBodyElement(0));
            Assert.AreEqual(2, doc.Paragraphs.Count);
            Assert.AreEqual(2, doc.BodyElements.Count);

            Assert.AreEqual(p3, doc.BodyElements[(0)]);
            Assert.AreEqual(p3, doc.Paragraphs[(0)]);
            Assert.AreEqual(p4, doc.BodyElements[(1)]);
            Assert.AreEqual(p4, doc.Paragraphs[(1)]);

            // Remove the last
            Assert.AreEqual(true, doc.RemoveBodyElement(1));
            Assert.AreEqual(1, doc.Paragraphs.Count);
            Assert.AreEqual(1, doc.BodyElements.Count);

            Assert.AreEqual(p3, doc.BodyElements[(0)]);
            Assert.AreEqual(p3, doc.Paragraphs[(0)]);
        }

        [Test]
        public void TestRegisterPackagePictureData()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_1.docx");

            /* manually assemble a new image package part*/
            OPCPackage opcPckg = doc.Package;
            XWPFRelation jpgRelation = XWPFRelation.IMAGE_JPEG;
            PackagePartName partName = PackagingUriHelper.CreatePartName(jpgRelation.DefaultFileName.Replace('#', '2'));
            PackagePart newImagePart = opcPckg.CreatePart(partName, jpgRelation.ContentType);
            byte[] nature1 = XWPFTestDataSamples.GetImage("abstract4.jpg");
            Stream os = newImagePart.GetOutputStream();
            os.Write(nature1, 0, nature1.Length);
            os.Close();
            XWPFHeader xwpfHeader = doc.HeaderList[(0)];
            PackageRelationship relationship = xwpfHeader.GetPackagePart().AddRelationship(partName, TargetMode.Internal, jpgRelation.Relation);
            XWPFPictureData newPicData = new XWPFPictureData(newImagePart, relationship);
            /* new part is now Ready to rumble */

            Assert.IsFalse(xwpfHeader.AllPictures.Contains(newPicData));
            Assert.IsFalse(doc.AllPictures.Contains(newPicData));
            Assert.IsFalse(doc.AllPackagePictures.Contains(newPicData));

            doc.RegisterPackagePictureData(newPicData);

            Assert.IsFalse(xwpfHeader.AllPictures.Contains(newPicData));
            Assert.IsFalse(doc.AllPictures.Contains(newPicData));
            Assert.IsTrue(doc.AllPackagePictures.Contains(newPicData));

            doc.Package.Revert();
        }

        [Test]
        public void TestFindPackagePictureData()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_1.docx");
            byte[] nature1 = XWPFTestDataSamples.GetImage("nature1.gif");
            XWPFPictureData part = doc.FindPackagePictureData(nature1, (int)PictureType.GIF);
            Assert.IsNotNull(part);
            Assert.IsTrue(doc.AllPictures.Contains(part));
            Assert.IsTrue(doc.AllPackagePictures.Contains(part));
            doc.Package.Revert();
        }

        [Test]
        public void TestGetAllPictures()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_3.docx");
            IList<XWPFPictureData> allPictures = doc.AllPictures;
            IList<XWPFPictureData> allPackagePictures = doc.AllPackagePictures;

            Assert.IsNotNull(allPictures);
            Assert.AreEqual(3, allPictures.Count);
            foreach (XWPFPictureData xwpfPictureData in allPictures)
            {
                Assert.IsTrue(allPackagePictures.Contains(xwpfPictureData));
            }

            try
            {
                allPictures.Add(allPictures[0]);
                Assert.Fail("This list must be unmodifiable!");
            }
            catch (NotSupportedException)
            {
                // all ok
            }

            doc.Package.Revert();
        }

        [Test]
        public void TestGetAllPackagePictures()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_3.docx");
            IList<XWPFPictureData> allPackagePictures = doc.AllPackagePictures;

            Assert.IsNotNull(allPackagePictures);
            Assert.AreEqual(5, allPackagePictures.Count);

            try
            {
                allPackagePictures.Add(allPackagePictures[0]);
                Assert.Fail("This list must be unmodifiable!");
            }
            catch (NotSupportedException)
            {
                // all ok
            }

            doc.Package.Revert();
        }

        [Test]
        public void TestPictureHandlingSimpleFile()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_1.docx");
            Assert.AreEqual(1, doc.AllPackagePictures.Count);
            byte[] newPic = XWPFTestDataSamples.GetImage("abstract4.jpg");
            String id1 = doc.AddPictureData(newPic, (int)PictureType.JPEG);
            Assert.AreEqual(2, doc.AllPackagePictures.Count);
            /* copy data, to avoid instance-Equality */
            byte[] newPicCopy = Arrays.CopyOf(newPic, newPic.Length);
            String id2 = doc.AddPictureData(newPicCopy, (int)PictureType.JPEG);
            Assert.AreEqual(id1, id2);
            doc.Package.Revert();
        }

        [Test]
        public void TestPictureHandlingHeaderDocumentImages()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_2.docx");
            Assert.AreEqual(1, doc.AllPictures.Count);
            Assert.AreEqual(1, doc.AllPackagePictures.Count);
            Assert.AreEqual(1, doc.HeaderList[(0)].AllPictures.Count);
            doc.Package.Revert();
        }

        [Test]
        public void TestPictureHandlingComplex()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_3.docx");
            XWPFHeader xwpfHeader = doc.HeaderList[(0)];

            Assert.AreEqual(3, doc.AllPictures.Count);
            Assert.AreEqual(3, xwpfHeader.AllPictures.Count);
            Assert.AreEqual(5, doc.AllPackagePictures.Count);

            byte[] nature1 = XWPFTestDataSamples.GetImage("nature1.jpg");
            String id = doc.AddPictureData(nature1, (int)PictureType.JPEG);
            POIXMLDocumentPart part1 = xwpfHeader.GetRelationById("rId1");
            XWPFPictureData part2 = (XWPFPictureData)doc.GetRelationById(id);
            Assert.AreSame(part1, part2);

            doc.Package.Revert();
        }

        [Test]
        public void TestZeroLengthLibreOfficeDocumentWithWaterMarkHeader()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("zero-length.docx");
            POIXMLProperties properties = doc.GetProperties();

            Assert.IsNotNull(properties.CoreProperties);

            XWPFHeader headerArray = doc.GetHeaderArray(0);
            Assert.AreEqual(1, headerArray.AllPictures.Count);
            Assert.AreEqual("image1.png", headerArray.AllPictures[0].FileName);
            Assert.AreEqual("", headerArray.Text);

            ExtendedProperties extendedProperties = properties.ExtendedProperties;
            Assert.IsNotNull(extendedProperties);
            Assert.AreEqual(0, extendedProperties.GetUnderlyingProperties().Characters);
        }

        [Test]
        public void TestSettings()
        {
            XWPFSettings settings = new XWPFSettings();
            settings.SetZoomPercent(50);
            Assert.AreEqual(50, settings.GetZoomPercent());
        }
    }

}