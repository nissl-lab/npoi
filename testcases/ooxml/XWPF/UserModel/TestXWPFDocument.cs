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

namespace TestCases.XWPF.UserModel
{
    using NPOI;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;
    using NPOI.XWPF;
    using NPOI.XWPF.Extractor;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using TestCases;

    [TestFixture]
    public class TestXWPFDocument
    {

        [Test]
        public void TestContainsMainContentType()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            OPCPackage pack = doc.Package;

            bool found = false;
            foreach(PackagePart part in pack.GetParts())
            {
                if(part.ContentType.Equals(XWPFRelation.DOCUMENT.ContentType))
                {
                    found = true;
                }
                if(false == found)
                {
                    // successful tests should be silent
                    System.Console.WriteLine(part);
                }
            }
            ClassicAssert.IsTrue(found);
        }

        [Test]
        public void TestOpen()
        {
            XWPFDocument xml;

            // Simple file
            xml = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            // Check it has key parts
            ClassicAssert.IsNotNull(xml.Document);
            ClassicAssert.IsNotNull(xml.Document.body);
            ClassicAssert.IsNotNull(xml.GetStyles());

            // Complex file
            xml = XWPFTestDataSamples.OpenSampleDocument("IllustrativeCases.docx");
            ClassicAssert.IsNotNull(xml.Document);
            ClassicAssert.IsNotNull(xml.Document.body);
            ClassicAssert.IsNotNull(xml.GetStyles());
        }

        [Test]
        public void TestMetadataBasics()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            ClassicAssert.IsNotNull(xml.GetProperties().CoreProperties);
            ClassicAssert.IsNotNull(xml.GetProperties().ExtendedProperties);

            ClassicAssert.AreEqual("Microsoft Office Word", xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Application);
            ClassicAssert.AreEqual(1315, xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Characters);
            ClassicAssert.AreEqual(10, xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Lines);

            ClassicAssert.AreEqual(null, xml.GetProperties().CoreProperties.Title);
            ClassicAssert.AreEqual(null, xml.GetProperties().CoreProperties.GetUnderlyingProperties().GetSubjectProperty());
        }

        [Test]
        public void TestMetadataComplex()
        {
            XWPFDocument xml = XWPFTestDataSamples.OpenSampleDocument("IllustrativeCases.docx");
            ClassicAssert.IsNotNull(xml.GetProperties().CoreProperties);
            ClassicAssert.IsNotNull(xml.GetProperties().ExtendedProperties);

            ClassicAssert.AreEqual("Microsoft Office Outlook", xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Application);
            ClassicAssert.AreEqual(5184, xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Characters);
            ClassicAssert.AreEqual(0, xml.GetProperties().ExtendedProperties.GetUnderlyingProperties().Lines);

            ClassicAssert.AreEqual(" ", xml.GetProperties().CoreProperties.Title);
            ClassicAssert.AreEqual(" ", xml.GetProperties().CoreProperties.GetUnderlyingProperties().GetSubjectProperty());
        }

        [Test]
        public void TestWorkbookProperties()
        {
            XWPFDocument doc = new XWPFDocument();
            POIXMLProperties props = doc.GetProperties();
            ClassicAssert.IsNotNull(props);
            ClassicAssert.AreEqual("NPOI", props.ExtendedProperties.GetUnderlyingProperties().Application);
        }

        [Test]
        public void TestAddParagraph()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            ClassicAssert.AreEqual(3, doc.Paragraphs.Count);

            XWPFParagraph p = doc.CreateParagraph();
            ClassicAssert.AreEqual(p, doc.Paragraphs[(3)]);
            ClassicAssert.AreEqual(4, doc.Paragraphs.Count);

            ClassicAssert.AreEqual(3, doc.GetParagraphPos(3));
            ClassicAssert.AreEqual(3, doc.GetPosOfParagraph(p));

            //CTP ctp = p.CTP;
            //XWPFParagraph newP = doc.GetParagraph(ctp);
            //ClassicAssert.AreSame(p, newP);
            //XmlCursor cursor = doc.Document.Body.GetPArray(0).newCursor();
            //XWPFParagraph cP = doc.InsertNewParagraph(cursor);
            //ClassicAssert.AreSame(cP, doc.Paragraphs[(0)]);
            //ClassicAssert.AreEqual(5, doc.Paragraphs.Count);
        }

        [Test]
        public void ReplaceParagraphText()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("WordReplaceCRLF.docx");

            //Find and replace text in document body
            doc.FindAndReplaceText("$replace_text$", "Regel1\nRegel2\nRegel3");

            //Find and replace text io tabel cell
            doc.FindAndReplaceText("$replace_cell_text$", "Regel1\nRegel2\nRegel3");

            //Save Word Document
            XWPFDocument outputDocument = outputDocument = XWPFTestDataSamples.WriteOutAndReadBack(doc);

            //Combine all runs of all paragraphs
            StringBuilder builder = new StringBuilder();
            foreach (var paragraph in outputDocument.Paragraphs)
            {
                foreach (var run in  paragraph.Runs)
                {
                    builder.Append(run.GetText(0));
                }
            }

            //Check 
            ClassicAssert.AreEqual("Regel1\nRegel2\nRegel3", builder.ToString());

            //Check text was replaced correctly in table cell
            var table = outputDocument.Tables.FirstOrDefault();
            ClassicAssert.IsNotNull(table);

            var dataRow = table.Rows[1];
            builder.Clear();
            foreach (var tableCellParagraph in dataRow.GetCell(0).Paragraphs)
            {
                foreach(var run in tableCellParagraph.Runs)
                {
                    builder.Append(run.GetText(0));
                }
            }

            //Check 
            ClassicAssert.AreEqual("Table replace multiple enters Regel1\nRegel2\nRegel3 text after last enter", builder.ToString());

        }

        [Test]
        public void FindAndReplaceTextInParagraph()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("WordFindAndReplaceTextInParagraph.docx");
            string initialText = string.Concat(doc.Paragraphs.Select(t => t.Text));

            Dictionary<string, string> replacers = new Dictionary<string, string>
            {
                {"$FINALQUALIFYINGWORK_QUESTION_1_ASKING_SHORT$", "Asking1" },
                {"$FINALQUALIFYINGWORK_QUESTION_1_QUESTION$", "Question1" },
                {"$FINALQUALIFYINGWORK_QUESTION_2_ASKING_SHORT$", "Asking2" },
                {"$FINALQUALIFYINGWORK_QUESTION_2_QUESTION$", "Question2" },
                {"$STUDENT_FULL$", "Ќа русском с пробелами" },
                {"$FINALQUALIFYINGWORK_GRADE$", "5" },
                {"$SIMPLE$", "Last text" },

            };

            //This is calling FindAndReplaceTextInParagraph for each paragraph in document
            foreach(var replacer in replacers)
                doc.FindAndReplaceText(replacer.Key, replacer.Value);
            
            //Save Word Document
            XWPFDocument outputDocument = outputDocument = XWPFTestDataSamples.WriteOutAndReadBack(doc);

            foreach(var replacer in replacers)
                initialText = initialText.Replace(replacer.Key, replacer.Value);

            string savedText = string.Concat(outputDocument.Paragraphs.Select(t => t.Text));

            //Check that at least replacing in string equal to result file
            ClassicAssert.AreEqual(initialText, savedText);

            //Check
            ClassicAssert.AreEqual("Some initial text на разный манер (inserted) and so on:Asking1: Question1Asking2: Question2Result on:1. Say that Ќа русском с пробелами with a very long sentence and one more replacer in the end for (русский €зык) sure 5Last text", 
                savedText);
        }

        [Test]
        public void TestAddPicture()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            byte[] jpeg = XWPFTestDataSamples.GetImage("nature1.jpg");
            String relationId = doc.AddPictureData(jpeg, (int)PictureType.JPEG);

            byte[] newJpeg = ((XWPFPictureData)doc.GetRelationById(relationId)).Data;
            ClassicAssert.AreEqual(newJpeg.Length, jpeg.Length);
            for(int i = 0; i < jpeg.Length; i++)
            {
                ClassicAssert.AreEqual(newJpeg[i], jpeg[i]);
            }
        }
        [Test]
        public void TestAllPictureFormats()
        {
            XWPFDocument doc = new XWPFDocument();

            doc.AddPictureData(new byte[10], (int) PictureType.EMF);
            doc.AddPictureData(new byte[11], (int) PictureType.WMF);
            doc.AddPictureData(new byte[12], (int) PictureType.PICT);
            doc.AddPictureData(new byte[13], (int) PictureType.JPEG);
            doc.AddPictureData(new byte[14], (int) PictureType.PNG);
            doc.AddPictureData(new byte[15], (int) PictureType.DIB);
            doc.AddPictureData(new byte[16], (int) PictureType.GIF);
            doc.AddPictureData(new byte[17], (int) PictureType.TIFF);
            doc.AddPictureData(new byte[18], (int) PictureType.EPS);
            doc.AddPictureData(new byte[19], (int) PictureType.BMP);
            doc.AddPictureData(new byte[20], (int) PictureType.WPG);

            ClassicAssert.AreEqual(11, doc.AllPictures.Count);

            doc = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            ClassicAssert.AreEqual(11, doc.AllPictures.Count);

        }
        [Test]
        public void TestRemoveBodyElement()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");
            ClassicAssert.AreEqual(3, doc.Paragraphs.Count);
            ClassicAssert.AreEqual(3, doc.BodyElements.Count);

            XWPFParagraph p1 = doc.Paragraphs[(0)];
            XWPFParagraph p2 = doc.Paragraphs[(1)];
            XWPFParagraph p3 = doc.Paragraphs[(2)];

            ClassicAssert.AreEqual(p1, doc.BodyElements[(0)]);
            ClassicAssert.AreEqual(p1, doc.Paragraphs[(0)]);
            ClassicAssert.AreEqual(p2, doc.BodyElements[(1)]);
            ClassicAssert.AreEqual(p2, doc.Paragraphs[(1)]);
            ClassicAssert.AreEqual(p3, doc.BodyElements[(2)]);
            ClassicAssert.AreEqual(p3, doc.Paragraphs[(2)]);

            // Add another
            XWPFParagraph p4 = doc.CreateParagraph();

            ClassicAssert.AreEqual(4, doc.Paragraphs.Count);
            ClassicAssert.AreEqual(4, doc.BodyElements.Count);
            ClassicAssert.AreEqual(p1, doc.BodyElements[(0)]);
            ClassicAssert.AreEqual(p1, doc.Paragraphs[(0)]);
            ClassicAssert.AreEqual(p2, doc.BodyElements[(1)]);
            ClassicAssert.AreEqual(p2, doc.Paragraphs[(1)]);
            ClassicAssert.AreEqual(p3, doc.BodyElements[(2)]);
            ClassicAssert.AreEqual(p3, doc.Paragraphs[(2)]);
            ClassicAssert.AreEqual(p4, doc.BodyElements[(3)]);
            ClassicAssert.AreEqual(p4, doc.Paragraphs[(3)]);

            // Remove the 2nd
            ClassicAssert.AreEqual(true, doc.RemoveBodyElement(1));
            ClassicAssert.AreEqual(3, doc.Paragraphs.Count);
            ClassicAssert.AreEqual(3, doc.BodyElements.Count);

            ClassicAssert.AreEqual(p1, doc.BodyElements[(0)]);
            ClassicAssert.AreEqual(p1, doc.Paragraphs[(0)]);
            ClassicAssert.AreEqual(p3, doc.BodyElements[(1)]);
            ClassicAssert.AreEqual(p3, doc.Paragraphs[(1)]);
            ClassicAssert.AreEqual(p4, doc.BodyElements[(2)]);
            ClassicAssert.AreEqual(p4, doc.Paragraphs[(2)]);

            // Remove the 1st
            ClassicAssert.AreEqual(true, doc.RemoveBodyElement(0));
            ClassicAssert.AreEqual(2, doc.Paragraphs.Count);
            ClassicAssert.AreEqual(2, doc.BodyElements.Count);

            ClassicAssert.AreEqual(p3, doc.BodyElements[(0)]);
            ClassicAssert.AreEqual(p3, doc.Paragraphs[(0)]);
            ClassicAssert.AreEqual(p4, doc.BodyElements[(1)]);
            ClassicAssert.AreEqual(p4, doc.Paragraphs[(1)]);

            // Remove the last
            ClassicAssert.AreEqual(true, doc.RemoveBodyElement(1));
            ClassicAssert.AreEqual(1, doc.Paragraphs.Count);
            ClassicAssert.AreEqual(1, doc.BodyElements.Count);

            ClassicAssert.AreEqual(p3, doc.BodyElements[(0)]);
            ClassicAssert.AreEqual(p3, doc.Paragraphs[(0)]);
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

            ClassicAssert.IsFalse(xwpfHeader.AllPictures.Contains(newPicData));
            ClassicAssert.IsFalse(doc.AllPictures.Contains(newPicData));
            ClassicAssert.IsFalse(doc.AllPackagePictures.Contains(newPicData));

            doc.RegisterPackagePictureData(newPicData);

            ClassicAssert.IsFalse(xwpfHeader.AllPictures.Contains(newPicData));
            ClassicAssert.IsFalse(doc.AllPictures.Contains(newPicData));
            ClassicAssert.IsTrue(doc.AllPackagePictures.Contains(newPicData));

            doc.Package.Revert();
        }

        [Test]
        public void TestFindPackagePictureData()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_1.docx");
            byte[] nature1 = XWPFTestDataSamples.GetImage("nature1.gif");
            XWPFPictureData part = doc.FindPackagePictureData(nature1, (int)PictureType.GIF);
            ClassicAssert.IsNotNull(part);
            ClassicAssert.IsTrue(doc.AllPictures.Contains(part));
            ClassicAssert.IsTrue(doc.AllPackagePictures.Contains(part));
            doc.Package.Revert();
        }

        [Test]
        public void TestGetAllPictures()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_3.docx");
            IList<XWPFPictureData> allPictures = doc.AllPictures;
            IList<XWPFPictureData> allPackagePictures = doc.AllPackagePictures;

            ClassicAssert.IsNotNull(allPictures);
            ClassicAssert.AreEqual(3, allPictures.Count);
            foreach(XWPFPictureData xwpfPictureData in allPictures)
            {
                ClassicAssert.IsTrue(allPackagePictures.Contains(xwpfPictureData));
            }

            try
            {
                allPictures.Add(allPictures[0]);
                Assert.Fail("This list must be unmodifiable!");
            }
            catch(NotSupportedException)
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

            ClassicAssert.IsNotNull(allPackagePictures);
            ClassicAssert.AreEqual(5, allPackagePictures.Count);

            try
            {
                allPackagePictures.Add(allPackagePictures[0]);
                Assert.Fail("This list must be unmodifiable!");
            }
            catch(NotSupportedException)
            {
                // all ok
            }

            doc.Package.Revert();
        }

        [Test]
        public void TestPictureHandlingSimpleFile()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_1.docx");
            ClassicAssert.AreEqual(1, doc.AllPackagePictures.Count);
            byte[] newPic = XWPFTestDataSamples.GetImage("abstract4.jpg");
            String id1 = doc.AddPictureData(newPic, (int)PictureType.JPEG);
            ClassicAssert.AreEqual(2, doc.AllPackagePictures.Count);
            /* copy data, to avoid instance-Equality */
            byte[] newPicCopy = Arrays.CopyOf(newPic, newPic.Length);
            String id2 = doc.AddPictureData(newPicCopy, (int)PictureType.JPEG);
            ClassicAssert.AreEqual(id1, id2);
            doc.Package.Revert();
        }

        [Test]
        public void TestPictureHandlingHeaderDocumentImages()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_2.docx");
            ClassicAssert.AreEqual(1, doc.AllPictures.Count);
            ClassicAssert.AreEqual(1, doc.AllPackagePictures.Count);
            ClassicAssert.AreEqual(1, doc.HeaderList[(0)].AllPictures.Count);
            doc.Package.Revert();
        }

        [Test]
        public void TestPictureHandlingComplex()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("issue_51265_3.docx");
            XWPFHeader xwpfHeader = doc.HeaderList[(0)];

            ClassicAssert.AreEqual(3, doc.AllPictures.Count);
            ClassicAssert.AreEqual(3, xwpfHeader.AllPictures.Count);
            ClassicAssert.AreEqual(5, doc.AllPackagePictures.Count);

            byte[] nature1 = XWPFTestDataSamples.GetImage("nature1.jpg");
            String id = doc.AddPictureData(nature1, (int)PictureType.JPEG);
            POIXMLDocumentPart part1 = xwpfHeader.GetRelationById("rId1");
            XWPFPictureData part2 = (XWPFPictureData)doc.GetRelationById(id);
            ClassicAssert.AreSame(part1, part2);

            doc.Package.Revert();
        }

        [Test]
        public void TestZeroLengthLibreOfficeDocumentWithWaterMarkHeader()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("zero-length.docx");
            POIXMLProperties properties = doc.GetProperties();

            ClassicAssert.IsNotNull(properties.CoreProperties);

            XWPFHeader headerArray = doc.GetHeaderArray(0);
            ClassicAssert.AreEqual(1, headerArray.AllPictures.Count);
            ClassicAssert.AreEqual("image1.png", headerArray.AllPictures[0].FileName);
            ClassicAssert.AreEqual("", headerArray.Text);

            ExtendedProperties extendedProperties = properties.ExtendedProperties;
            ClassicAssert.IsNotNull(extendedProperties);
            ClassicAssert.AreEqual(0, extendedProperties.GetUnderlyingProperties().Characters);
        }

        [Test]
        public void TestSettings()
        {
            XWPFSettings settings = new XWPFSettings();
            settings.SetZoomPercent(50);
            ClassicAssert.AreEqual(50, settings.GetZoomPercent());
        }

        [Test]
        public void TestEnforcedWith()
        {
            XWPFDocument docx = XWPFTestDataSamples.OpenSampleDocument("EnforcedWith.docx");
            ClassicAssert.IsTrue(docx.IsEnforcedProtection());
            docx.Close();
        }

        [Test]
        [Ignore("XWPF should be able to write to a new Stream when opened Read-Only")]
        public void TestWriteFromReadOnlyOPC()
        {
            OPCPackage opc = OPCPackage.Open(
                    POIDataSamples.GetDocumentInstance().GetFileInfo("SampleDoc.docx"),
                    PackageAccess.READ
            );
            XWPFDocument doc = new XWPFDocument(opc);
            XWPFWordExtractor ext = new XWPFWordExtractor(doc);
            String origText = ext.Text;

            doc = XWPFTestDataSamples.WriteOutAndReadBack(doc);
            ext = new XWPFWordExtractor(doc);

            ClassicAssert.AreEqual(origText, ext.Text);
        }

        [Test]
        public void TestDocVars()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("docVars.docx");

            foreach(var part in doc.RelationParts)
            {
                var relationType = part.Relationship.RelationshipType;
                if(relationType == XWPFRelation.SETTINGS.Relation)
                {
                    var settingsPart = (XWPFSettings)part.DocumentPart;

                    using(var stream = settingsPart.GetPackagePart().GetInputStream())
                    {
                        var xmldoc = POIXMLDocumentPart.ConvertStreamToXml(stream);
                        var ctSettings = SettingsDocument.Parse(xmldoc, POIXMLDocumentPart.NamespaceManager).Settings;
                        var variables = ctSettings.docVars;

                        ClassicAssert.IsNotNull(variables);
                        ClassicAssert.AreEqual(5, variables.docVar.Count);

                        for(int i = 0; i<variables.docVar.Count; i++)
                            ClassicAssert.AreEqual((i+1).ToString(), variables.docVar[i].val);
                    }
                }
            }
        }
    }
}