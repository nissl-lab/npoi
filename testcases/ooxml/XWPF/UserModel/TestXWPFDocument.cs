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





    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.OpenXml4Net.OPC;

    [TestClass]
    public class TestXWPFDocument
    {

        [TestMethod]
        public void TestContainsMainContentType(){
		XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
		OPCPackage pack = doc.Package;

		bool found = false;
		foreach(PackagePart part in pack.Parts) {
			if(part.ContentType.Equals(XWPFRelation.DOCUMENT.ContentType)) {
				found = true;
			}
			if (false) {
				// successful tests should be silent
				System.out.Println(part);
			}
		}
		Assert.IsTrue(found);
	}

        [TestMethod]
        public void TestOpen()
        {
            XWPFDocument xml;

            // Simple file
            xml = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            // Check it has key parts
            Assert.IsNotNull(xml.Document);
            Assert.IsNotNull(xml.Document.Body);
            Assert.IsNotNull(xml.Style);

            // Complex file
            xml = XWPFTestDataSamples.OpenSampleDocument("IllustrativeCases.docx");
            Assert.IsNotNull(xml.Document);
            Assert.IsNotNull(xml.Document.Body);
            Assert.IsNotNull(xml.Style);
        }

        [TestMethod]
        public void TestMetadataBasics()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            Assert.IsNotNull(xml.Properties.CoreProperties);
            Assert.IsNotNull(xml.Properties.ExtendedProperties);

            Assert.AreEqual("Microsoft Office Word", xml.Properties.ExtendedProperties.UnderlyingProperties.Application);
            Assert.AreEqual(1315, xml.Properties.ExtendedProperties.UnderlyingProperties.Characters);
            Assert.AreEqual(10, xml.Properties.ExtendedProperties.UnderlyingProperties.Lines);

            Assert.AreEqual(null, xml.Properties.CoreProperties.Title);
            Assert.AreEqual(null, xml.Properties.CoreProperties.UnderlyingProperties.SubjectProperty.Value);
        }

        [TestMethod]
        public void TestMetadataComplex()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("IllustrativeCases.docx");
            Assert.IsNotNull(xml.Properties.CoreProperties);
            Assert.IsNotNull(xml.Properties.ExtendedProperties);

            Assert.AreEqual("Microsoft Office Outlook", xml.Properties.ExtendedProperties.UnderlyingProperties.Application);
            Assert.AreEqual(5184, xml.Properties.ExtendedProperties.UnderlyingProperties.Characters);
            Assert.AreEqual(0, xml.Properties.ExtendedProperties.UnderlyingProperties.Lines);

            Assert.AreEqual(" ", xml.Properties.CoreProperties.Title);
            Assert.AreEqual(" ", xml.Properties.CoreProperties.UnderlyingProperties.SubjectProperty.Value);
        }

        [TestMethod]
        public void TestWorkbookProperties()
        {
            XWPFDocument doc = new XWPFDocument();
            POIXMLProperties props = doc.Properties;
            Assert.IsNotNull(props);
            Assert.AreEqual("Apache POI", props.ExtendedProperties.UnderlyingProperties.Application);
        }

        [TestMethod]
        public void TestAddParagraph()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            Assert.AreEqual(3, doc.Paragraphs.Size());

            XWPFParagraph p = doc.CreateParagraph();
            Assert.AreEqual(p, doc.Paragraphs.Get(3));
            Assert.AreEqual(4, doc.Paragraphs.Size());

            Assert.AreEqual(3, doc.GetParagraphPos(3));
            Assert.AreEqual(3, doc.GetPosOfParagraph(p));

            CTP ctp = p.CTP;
            XWPFParagraph newP = doc.GetParagraph(ctp);
            Assert.AreSame(p, newP);
            XmlCursor cursor = doc.Document.Body.GetPArray(0).newCursor();
            XWPFParagraph cP = doc.InsertNewParagraph(cursor);
            Assert.AreSame(cP, doc.Paragraphs.Get(0));
            Assert.AreEqual(5, doc.Paragraphs.Size());
        }

        public void testAddPicture()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            byte[] jpeg = XWPFTestDataSamples.GetImage("nature1.jpg");
            String relationId = doc.AddPictureData(jpeg, XWPFDocument.PICTURE_TYPE_JPEG);

            byte[] newJpeg = ((XWPFPictureData)doc.GetRelationById(relationId)).Data;
            Assert.AreEqual(newJpeg.Length, jpeg.Length);
            for (int i = 0; i < jpeg.Length; i++)
            {
                Assert.AreEqual(newJpeg[i], jpeg[i]);
            }
        }

        [TestMethod]
        public void TestRemoveBodyElement()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            Assert.AreEqual(3, doc.Paragraphs.Size());
            Assert.AreEqual(3, doc.BodyElements.Size());

            XWPFParagraph p1 = doc.Paragraphs.Get(0);
            XWPFParagraph p2 = doc.Paragraphs.Get(1);
            XWPFParagraph p3 = doc.Paragraphs.Get(2);

            Assert.AreEqual(p1, doc.BodyElements.Get(0));
            Assert.AreEqual(p1, doc.Paragraphs.Get(0));
            Assert.AreEqual(p2, doc.BodyElements.Get(1));
            Assert.AreEqual(p2, doc.Paragraphs.Get(1));
            Assert.AreEqual(p3, doc.BodyElements.Get(2));
            Assert.AreEqual(p3, doc.Paragraphs.Get(2));

            // Add another
            XWPFParagraph p4 = doc.CreateParagraph();

            Assert.AreEqual(4, doc.Paragraphs.Size());
            Assert.AreEqual(4, doc.BodyElements.Size());
            Assert.AreEqual(p1, doc.BodyElements.Get(0));
            Assert.AreEqual(p1, doc.Paragraphs.Get(0));
            Assert.AreEqual(p2, doc.BodyElements.Get(1));
            Assert.AreEqual(p2, doc.Paragraphs.Get(1));
            Assert.AreEqual(p3, doc.BodyElements.Get(2));
            Assert.AreEqual(p3, doc.Paragraphs.Get(2));
            Assert.AreEqual(p4, doc.BodyElements.Get(3));
            Assert.AreEqual(p4, doc.Paragraphs.Get(3));

            // Remove the 2nd
            Assert.AreEqual(true, doc.RemoveBodyElement(1));
            Assert.AreEqual(3, doc.Paragraphs.Size());
            Assert.AreEqual(3, doc.BodyElements.Size());

            Assert.AreEqual(p1, doc.BodyElements.Get(0));
            Assert.AreEqual(p1, doc.Paragraphs.Get(0));
            Assert.AreEqual(p3, doc.BodyElements.Get(1));
            Assert.AreEqual(p3, doc.Paragraphs.Get(1));
            Assert.AreEqual(p4, doc.BodyElements.Get(2));
            Assert.AreEqual(p4, doc.Paragraphs.Get(2));

            // Remove the 1st
            Assert.AreEqual(true, doc.RemoveBodyElement(0));
            Assert.AreEqual(2, doc.Paragraphs.Size());
            Assert.AreEqual(2, doc.BodyElements.Size());

            Assert.AreEqual(p3, doc.BodyElements.Get(0));
            Assert.AreEqual(p3, doc.Paragraphs.Get(0));
            Assert.AreEqual(p4, doc.BodyElements.Get(1));
            Assert.AreEqual(p4, doc.Paragraphs.Get(1));

            // Remove the last
            Assert.AreEqual(true, doc.RemoveBodyElement(1));
            Assert.AreEqual(1, doc.Paragraphs.Size());
            Assert.AreEqual(1, doc.BodyElements.Size());

            Assert.AreEqual(p3, doc.BodyElements.Get(0));
            Assert.AreEqual(p3, doc.Paragraphs.Get(0));
        }

        [TestMethod]
        public void TestRegisterPackagePictureData()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_1.docx");

            /* manually assemble a new image package part*/
            OPCPackage opcPckg = doc.Package;
            XWPFRelation jpgRelation = XWPFRelation.IMAGE_JPEG;
            PackagePartName partName = PackagingURIHelper.CreatePartName(jpgRelation.DefaultFileName.Replace('#', '2'));
            PackagePart newImagePart = opcPckg.CreatePart(partName, jpgRelation.ContentType);
            byte[] nature1 = XWPFTestDataSamples.GetImage("abstract4.jpg");
            OutputStream os = newImagePart.OutputStream;
            os.Write(nature1);
            os.Close();
            XWPFHeader xwpfHeader = doc.HeaderList.Get(0);
            PackageRelationship relationship = xwpfHeader.PackagePart.AddRelationship(partName, TargetMode.INTERNAL, jpgRelation.Relation);
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

        [TestMethod]
        public void TestFindPackagePictureData()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_1.docx");
            byte[] nature1 = XWPFTestDataSamples.GetImage("nature1.gif");
            XWPFPictureData part = doc.FindPackagePictureData(nature1, Document.PICTURE_TYPE_GIF);
            Assert.IsNotNull(part);
            Assert.IsTrue(doc.AllPictures.Contains(part));
            Assert.IsTrue(doc.AllPackagePictures.Contains(part));
            doc.Package.Revert();
        }

        [TestMethod]
        public void TestGetAllPictures()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_3.docx");
            List<XWPFPictureData> allPictures = doc.AllPictures;
            List<XWPFPictureData> allPackagePictures = doc.AllPackagePictures;

            Assert.IsNotNull(allPictures);
            Assert.AreEqual(3, allPictures.Size());
            foreach (XWPFPictureData xwpfPictureData in allPictures)
            {
                Assert.IsTrue(allPackagePictures.Contains(xwpfPictureData));
            }

            try
            {
                allPictures.Add(allPictures.Get(0));
                Assert.Fail("This list must be unmodifiable!");
            }
            catch (UnsupportedOperationException e)
            {
                // all ok
            }

            doc.Package.Revert();
        }

        [TestMethod]
        public void TestGetAllPackagePictures()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_3.docx");
            List<XWPFPictureData> allPackagePictures = doc.AllPackagePictures;

            Assert.IsNotNull(allPackagePictures);
            Assert.AreEqual(5, allPackagePictures.Size());

            try
            {
                allPackagePictures.Add(allPackagePictures.Get(0));
                Assert.Fail("This list must be unmodifiable!");
            }
            catch (UnsupportedOperationException e)
            {
                // all ok
            }

            doc.Package.Revert();
        }

        [TestMethod]
        public void TestPictureHandlingSimpleFile()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_1.docx");
            Assert.AreEqual(1, doc.AllPackagePictures.Size());
            byte[] newPic = XWPFTestDataSamples.GetImage("abstract4.jpg");
            String id1 = doc.AddPictureData(newPic, Document.PICTURE_TYPE_JPEG);
            Assert.AreEqual(2, doc.AllPackagePictures.Size());
            /* copy data, to avoid instance-Equality */
            byte[] newPicCopy = ArrayUtil.CopyOf(newPic, newPic.Length);
            String id2 = doc.AddPictureData(newPicCopy, Document.PICTURE_TYPE_JPEG);
            Assert.AreEqual(id1, id2);
            doc.Package.Revert();
        }

        [TestMethod]
        public void TestPictureHandlingHeaderDocumentImages()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_2.docx");
            Assert.AreEqual(1, doc.AllPictures.Size());
            Assert.AreEqual(1, doc.AllPackagePictures.Size());
            Assert.AreEqual(1, doc.HeaderList.Get(0).AllPictures.Size());
            doc.Package.Revert();
        }

        [TestMethod]
        public void TestPictureHandlingComplex()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_3.docx");
            XWPFHeader xwpfHeader = doc.HeaderList.Get(0);

            Assert.AreEqual(3, doc.AllPictures.Size());
            Assert.AreEqual(3, xwpfHeader.AllPictures.Size());
            Assert.AreEqual(5, doc.AllPackagePictures.Size());

            byte[] nature1 = XWPFTestDataSamples.GetImage("nature1.jpg");
            String id = doc.AddPictureData(nature1, Document.PICTURE_TYPE_JPEG);
            POIXMLDocumentPart part1 = xwpfHeader.GetRelationById("rId1");
            XWPFPictureData part2 = (XWPFPictureData)doc.GetRelationById(id);
            Assert.AreSame(part1, part2);

            doc.Package.Revert();
        }
    }

}