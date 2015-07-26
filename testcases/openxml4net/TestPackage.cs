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

using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.OPC.Internal;
using System.IO;
using System.Collections.Generic;
using System;
using TestCases.OpenXml4Net;
using NPOI.Util;
using System.Reflection;
using System.Text.RegularExpressions;
using NUnit.Framework;
using System.Xml;
using System.Text;
namespace TestCases.OPC
{
    [TestFixture]
    public class TestPackage
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(TestPackage));

        /**
         * Test that just opening and closing the file doesn't alter the document.
         */
        [Test]
        public void TestOpenSave()
        {
            String originalFile = OpenXml4NetTestDataSamples.GetSampleFileName("TestPackageCommon.docx");
            FileInfo targetFile = OpenXml4NetTestDataSamples.GetOutputFile("TestPackageOpenSaveTMP.docx");

            OPCPackage p = OPCPackage.Open(originalFile, PackageAccess.READ_WRITE);
            try
            {
                p.Save(targetFile.FullName);

                // Compare the original and newly saved document
                Assert.IsTrue(File.Exists(targetFile.FullName));
                ZipFileAssert.AssertEqual(new FileInfo(originalFile), targetFile);
                File.Delete(targetFile.FullName);
            }
            finally
            {
                // use revert to not re-write the input file
                p.Revert();
            }
        }

        /**
         * Test that when we create a new Package, we give it
         *  the correct default content types
         */
        [Test]
        public void TestCreateGetsContentTypes()
        {
            FileInfo targetFile = OpenXml4NetTestDataSamples.GetOutputFile("TestCreatePackageTMP.docx");

            // Zap the target file, in case of an earlier run
            if (File.Exists(targetFile.FullName))
                File.Delete(targetFile.FullName);

            OPCPackage pkg = OPCPackage.Create(targetFile.FullName);

            // Check it has content types for rels and xml
            ContentTypeManager ctm = GetContentTypeManager(pkg);
            Assert.AreEqual(
                    "application/xml",
                    ctm.GetContentType(
                            PackagingUriHelper.CreatePartName("/foo.xml")
                    )
            );
            Assert.AreEqual(
                    ContentTypes.RELATIONSHIPS_PART,
                    ctm.GetContentType(
                            PackagingUriHelper.CreatePartName("/foo.rels")
                    )
            );
            Assert.IsNull(
                    ctm.GetContentType(
                            PackagingUriHelper.CreatePartName("/foo.txt")
                    )
            );
        }

        /**
         * Test namespace creation.
         */
        [Test]
        public void TestCreatePackageAddPart()
        {
            FileInfo targetFile = OpenXml4NetTestDataSamples.GetOutputFile("TestCreatePackageTMP.docx");

            FileInfo expectedFile = OpenXml4NetTestDataSamples.GetSampleFile("TestCreatePackageOUTPUT.docx");

            // Zap the target file, in case of an earlier run
            if (targetFile.Exists) targetFile.Delete();

            // Create a namespace
            OPCPackage pkg = OPCPackage.Create(targetFile.FullName);
            PackagePartName corePartName = PackagingUriHelper
                    .CreatePartName("/word/document.xml");

            pkg.AddRelationship(corePartName, TargetMode.Internal,
                    PackageRelationshipTypes.CORE_DOCUMENT, "rId1");

            PackagePart corePart = pkg
                    .CreatePart(
                            corePartName,
                            "application/vnd.openxmlformats-officedocument.wordProcessingml.document.main+xml");

            XmlDocument doc = new XmlDocument();

            XmlNamespaceManager mgr = new XmlNamespaceManager(doc.NameTable);
            string wuri = "http://schemas.openxmlformats.org/wordProcessingml/2006/main";
            mgr.AddNamespace("w", wuri);
            XmlElement elDocument = doc.CreateElement("w:document", wuri);
            doc.AppendChild(elDocument);
            XmlElement elBody = doc.CreateElement("w:body", wuri);
            elDocument.AppendChild(elBody);
            XmlElement elParagraph = doc.CreateElement("w:p", wuri);
            elBody.AppendChild(elParagraph);
            XmlElement elRun = doc.CreateElement("w:r", wuri);
            elParagraph.AppendChild(elRun);
            XmlElement elText = doc.CreateElement("w:t", wuri);
            elRun.AppendChild(elText);
            elText.InnerText = ("Hello Open XML !");

            StreamHelper.SaveXmlInStream(doc, corePart.GetOutputStream());
            pkg.Close();

            ZipFileAssert.AssertEqual(expectedFile, targetFile);
            File.Delete(targetFile.FullName);
        }

        /**
         * Tests that we can create a new namespace, add a core
         *  document and another part, save and re-load and
         *  have everything Setup as expected
         */
        [Test]
        //[Ignore("add relation Uri #Sheet1!A1")]
        public void TestCreatePackageWithCoreDocument()
        {
            MemoryStream baos = new MemoryStream();
            OPCPackage pkg = OPCPackage.Create(baos);

            // Add a core document
            PackagePartName corePartName = PackagingUriHelper.CreatePartName("/xl/workbook.xml");
            // Create main part relationship
            pkg.AddRelationship(corePartName, TargetMode.Internal, PackageRelationshipTypes.CORE_DOCUMENT, "rId1");
            // Create main document part
            PackagePart corePart = pkg.CreatePart(corePartName, "application/vnd.Openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            // Put in some dummy content
            Stream coreOut = corePart.GetOutputStream();
            byte[] buffer = Encoding.UTF8.GetBytes("<dummy-xml />");
            coreOut.Write(buffer, 0, buffer.Length);
            coreOut.Close();

            // And another bit
            PackagePartName sheetPartName = PackagingUriHelper.CreatePartName("/xl/worksheets/sheet1.xml");
            PackageRelationship rel =
                 corePart.AddRelationship(sheetPartName, TargetMode.Internal, "http://schemas.Openxmlformats.org/officeDocument/2006/relationships/worksheet", "rSheet1");
            PackagePart part = pkg.CreatePart(sheetPartName, "application/vnd.Openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            Assert.IsNotNull(part);

            // Dummy content again
            coreOut = corePart.GetOutputStream();
            buffer = Encoding.UTF8.GetBytes("<dummy-xml2 />");
            coreOut.Write(buffer, 0, buffer.Length);
            coreOut.Close();

            //add a relationship with internal target: "#Sheet1!A1"
            corePart.AddRelationship(PackagingUriHelper.ToUri("#Sheet1!A1"), TargetMode.Internal, "http://schemas.Openxmlformats.org/officeDocument/2006/relationships/hyperlink", "rId2");

            // Check things are as expected
            PackageRelationshipCollection coreRels =
                pkg.GetRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT);
            Assert.AreEqual(1, coreRels.Size);
            PackageRelationship coreRel = coreRels.GetRelationship(0);
            Assert.AreEqual("/", coreRel.SourceUri.ToString());
            Assert.AreEqual("/xl/workbook.xml", coreRel.TargetUri.ToString());
            Assert.IsNotNull(pkg.GetPart(coreRel));


            // Save and re-load
            pkg.Close();
            FileInfo tmp = TempFile.CreateTempFile("testCreatePackageWithCoreDocument", ".zip");
            FileStream fout = new FileStream(tmp.FullName, FileMode.Create, FileAccess.ReadWrite);
            try
            {
                buffer = baos.ToArray();
                fout.Write(buffer, 0 , buffer.Length);
            }
            finally
            {
                fout.Close();
            }
            pkg = OPCPackage.Open(tmp.FullName);
            //tmp.Delete();

            try
            {
                // Check still right
                coreRels = pkg.GetRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT);
                Assert.AreEqual(1, coreRels.Size);
                coreRel = coreRels.GetRelationship(0);

                Assert.AreEqual("/", coreRel.SourceUri.ToString());
                Assert.AreEqual("/xl/workbook.xml", coreRel.TargetUri.ToString());
                corePart = pkg.GetPart(coreRel);
                Assert.IsNotNull(corePart);

                PackageRelationshipCollection rels = corePart.GetRelationshipsByType("http://schemas.Openxmlformats.org/officeDocument/2006/relationships/hyperlink");
                Assert.AreEqual(1, rels.Size);
                rel = rels.GetRelationship(0);
                //Assert.AreEqual("Sheet1!A1", rel.TargetUri.Fragment);

                assertMSCompatibility(pkg);
            }
            finally
            {
                pkg.Close();
            }

        }

        private void assertMSCompatibility(OPCPackage pkg)
        {
            PackagePartName relName = PackagingUriHelper.CreatePartName(PackageRelationship.ContainerPartRelationship);
            PackagePart relPart = pkg.GetPart(relName);

            XmlDocument xmlRelationshipsDoc = DocumentHelper.LoadDocument(relPart.GetInputStream());

            XmlElement root = xmlRelationshipsDoc.DocumentElement;
            XmlNodeList nodeList = root.GetElementsByTagName(PackageRelationship.RELATIONSHIP_TAG_NAME);
            int nodeCount = nodeList.Count;
            for (int i = 0; i < nodeCount; i++)
            {
                XmlElement element = (XmlElement)nodeList.Item(i);
                String value = element.GetAttribute(PackageRelationship.TARGET_ATTRIBUTE_NAME);
                Assert.IsTrue(value[0] != '/', "Root target must not start with a leading slash ('/'): " + value);
            }

        }

        /**
         * Test namespace opening.
         */
        [Test]
        public void TestOpenPackage()
        {
            FileInfo targetFile = OpenXml4NetTestDataSamples.GetOutputFile("TestOpenPackageTMP.docx");

            FileInfo inputFile = OpenXml4NetTestDataSamples.GetSampleFile("TestOpenPackageINPUT.docx");

            FileInfo expectedFile = OpenXml4NetTestDataSamples.GetSampleFile("TestOpenPackageOUTPUT.docx");

            // Copy the input file in the output directory
            FileHelper.CopyFile(inputFile.FullName, targetFile.FullName);

            // Create a namespace
            OPCPackage pkg = OPCPackage.Open(targetFile.FullName);

            // Modify core part
            PackagePartName corePartName = PackagingUriHelper
                    .CreatePartName("/word/document.xml");

            PackagePart corePart = pkg.GetPart(corePartName);

            // Delete some part to have a valid document
            foreach (PackageRelationship rel in corePart.Relationships)
            {
                corePart.RemoveRelationship(rel.Id);
                pkg.RemovePart(PackagingUriHelper.CreatePartName(PackagingUriHelper
                        .ResolvePartUri(corePart.PartName.URI, rel
                                .TargetUri)));
            }

            // Create a content
            XmlDocument doc = new XmlDocument();

            XmlNamespaceManager mgr = new XmlNamespaceManager(doc.NameTable);
            string wuri = "http://schemas.openxmlformats.org/wordProcessingml/2006/main";
            mgr.AddNamespace("w", wuri);
            XmlElement elDocument = doc.CreateElement("w:document", wuri);
            doc.AppendChild(elDocument);
            XmlElement elBody = doc.CreateElement("w:body", wuri);
            elDocument.AppendChild(elBody);
            XmlElement elParagraph = doc.CreateElement("w:p", wuri);
            elBody.AppendChild(elParagraph);
            XmlElement elRun = doc.CreateElement("w:r", wuri);
            elParagraph.AppendChild(elRun);
            XmlElement elText = doc.CreateElement("w:t", wuri);
            elRun.AppendChild(elText);
            elText.InnerText = ("Hello Open XML !");

            StreamHelper.SaveXmlInStream(doc, corePart.GetOutputStream());
            // Save and close
            try
            {
                pkg.Close();
            }
            catch (IOException)
            {
                Assert.Fail();
            }

            ZipFileAssert.AssertEqual(expectedFile, targetFile);
            File.Delete(targetFile.FullName);
        }

        /**
         * Checks that we can write a namespace to a simple
         *  OutputStream, in Addition to the normal writing
         *  to a file
         */
        [Test]
        public void TestSaveToOutputStream()
        {
            String originalFile = OpenXml4NetTestDataSamples.GetSampleFileName("TestPackageCommon.docx");
            FileInfo targetFile = OpenXml4NetTestDataSamples.GetOutputFile("TestPackageOpenSaveTMP.docx");

            OPCPackage p = OPCPackage.Open(originalFile, PackageAccess.READ_WRITE);
            try
            {
                FileStream fs = targetFile.OpenWrite();
                try
                {
                    p.Save(fs);
                }
                finally
                {
                    fs.Close();
                }
                
                // Compare the original and newly saved document
                Assert.IsTrue(File.Exists(targetFile.FullName));
                ZipFileAssert.AssertEqual(new FileInfo(originalFile), targetFile);
                File.Delete(targetFile.FullName);
            }
            finally
            {
                p.Revert();
            }
        }

        /**
         * Checks that we can open+read a namespace from a
         *  simple InputStream, in Addition to the normal
         *  Reading from a file
         */
        [Test]
        public void TestOpenFromInputStream()
        {
            String originalFile = OpenXml4NetTestDataSamples.GetSampleFileName("TestPackageCommon.docx");

            FileStream finp = new FileStream(originalFile, FileMode.Open, FileAccess.Read);

            OPCPackage p = OPCPackage.Open(finp);

            Assert.IsNotNull(p);
            Assert.IsNotNull(p.Relationships);
            Assert.AreEqual(12, p.GetParts().Count);

            // Check it has the usual bits
            Assert.IsTrue(p.HasRelationships);
            Assert.IsTrue(p.ContainPart(PackagingUriHelper.CreatePartName("/_rels/.rels")));
        }

        /**
         * TODO: fix and enable
         */
        //[Test]
        public void disabled_testRemovePartRecursive()
        {
            String originalFile = OpenXml4NetTestDataSamples.GetSampleFileName("TestPackageCommon.docx");
            FileInfo targetFile = OpenXml4NetTestDataSamples.GetOutputFile("TestPackageRemovePartRecursiveOUTPUT.docx");
            FileInfo tempFile = OpenXml4NetTestDataSamples.GetOutputFile("TestPackageRemovePartRecursiveTMP.docx");

            OPCPackage p = OPCPackage.Open(originalFile, PackageAccess.READ_WRITE);
            p.RemovePartRecursive(PackagingUriHelper.CreatePartName(new Uri(
                    "/word/document.xml", UriKind.Relative)));
            p.Save(tempFile.FullName);

            // Compare the original and newly saved document
            Assert.IsTrue(File.Exists(targetFile.FullName));
            ZipFileAssert.AssertEqual(targetFile, tempFile);
            File.Delete(targetFile.FullName);
        }
        [Test]
        public void TestDeletePart()
        {
            Dictionary<PackagePartName, String> expectedValues;
            Dictionary<PackagePartName, String> values;

            values = new Dictionary<PackagePartName, String>();

            // Expected values
            expectedValues = new Dictionary<PackagePartName, String>();
            expectedValues.Add(PackagingUriHelper.CreatePartName("/_rels/.rels"),
                    "application/vnd.openxmlformats-package.relationships+xml");

            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/docProps/app.xml"),
                            "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            expectedValues.Add(PackagingUriHelper
                    .CreatePartName("/docProps/core.xml"),
                    "application/vnd.openxmlformats-package.core-properties+xml");
            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/word/fontTable.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml");
            expectedValues.Add(PackagingUriHelper
                    .CreatePartName("/word/media/image1.gif"), "image/gif");
            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/word/settings.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");
            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/word/styles.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
            expectedValues.Add(PackagingUriHelper
                    .CreatePartName("/word/theme/theme1.xml"),
                    "application/vnd.openxmlformats-officedocument.theme+xml");
            expectedValues
                    .Add(
                            PackagingUriHelper
                                    .CreatePartName("/word/webSettings.xml"),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml");

            String filepath = OpenXml4NetTestDataSamples.GetSampleFileName("sample.docx");

            OPCPackage p = OPCPackage.Open(filepath, PackageAccess.READ_WRITE);
            // Remove the core part
            p.DeletePart(PackagingUriHelper.CreatePartName("/word/document.xml"));

            foreach (PackagePart part in p.GetParts())
            {
                values.Add(part.PartName, part.ContentType);
                logger.Log(POILogger.DEBUG, part.PartName);
            }

            // Compare expected values with values return by the namespace
            foreach (PackagePartName partName in expectedValues.Keys)
            {
                Assert.IsNotNull(values[partName]);
                Assert.AreEqual(expectedValues[partName], values[partName]);
            }
            // Don't save modifications
            p.Revert();
        }
        [Test]
        public void TestDeletePartRecursive()
        {
            Dictionary<PackagePartName, String> expectedValues;
            Dictionary<PackagePartName, String> values;

            values = new Dictionary<PackagePartName, String>();

            // Expected values
            expectedValues = new Dictionary<PackagePartName, String>();
            expectedValues.Add(PackagingUriHelper.CreatePartName("/_rels/.rels"),
                    "application/vnd.openxmlformats-package.relationships+xml");

            expectedValues
                    .Add(PackagingUriHelper.CreatePartName("/docProps/app.xml"),
                            "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            expectedValues.Add(PackagingUriHelper
                    .CreatePartName("/docProps/core.xml"),
                    "application/vnd.openxmlformats-package.core-properties+xml");

            String filepath = OpenXml4NetTestDataSamples.GetSampleFileName("sample.docx");

            OPCPackage p = OPCPackage.Open(filepath, PackageAccess.READ_WRITE);
            // Remove the core part
            p.DeletePartRecursive(PackagingUriHelper.CreatePartName("/word/document.xml"));

            foreach (PackagePart part in p.GetParts())
            {
                values.Add(part.PartName, part.ContentType);
                logger.Log(POILogger.DEBUG, part.PartName);
            }

            // Compare expected values with values return by the namespace
            foreach (PackagePartName partName in expectedValues.Keys)
            {
                Assert.IsNotNull(values[partName]);
                Assert.AreEqual(expectedValues[partName], values[partName]);
            }
            // Don't save modifications
            p.Revert();
        }

        /**
         * Test that we can open a file by path, and then
         *  write Changes to it.
         */
        [Test]
        public void TestOpenFileThenOverWrite()
        {
            string tempFile = TempFile.GetTempFilePath("poiTesting", "tmp");
            FileInfo origFile = OpenXml4NetTestDataSamples.GetSampleFile("TestPackageCommon.docx");
            FileHelper.CopyFile(origFile.FullName, tempFile);

            // Open the temp file
            OPCPackage p = OPCPackage.Open(tempFile, PackageAccess.READ_WRITE);
            // Close it
            p.Close();
            // Delete it
            File.Delete(tempFile);

            // Reset
            FileHelper.CopyFile(origFile.FullName, tempFile);
            p = OPCPackage.Open(tempFile, PackageAccess.READ_WRITE);

            // Save it to the same file - not allowed
            try
            {
                p.Save(tempFile);
                Assert.Fail("You shouldn't be able to call save(File) to overwrite the current file");
            }
            catch (IOException) { }

            p.Close();
            // Delete it
            File.Delete(tempFile);


            // Open it read only, then close and delete - allowed
            FileHelper.CopyFile(origFile.FullName, tempFile);
            p = OPCPackage.Open(tempFile, PackageAccess.READ);
            p.Close();
            File.Delete(tempFile);
        }
        /**
         * Test that we can open a file by path, save it
         *  to another file, then delete both
         */
        [Test]
        public void TestOpenFileThenSaveDelete()
        {
            string tempFile = TempFile.GetTempFilePath("poiTesting", "tmp");
            string tempFile2 = TempFile.GetTempFilePath("poiTesting", "tmp");
            FileInfo origFile = OpenXml4NetTestDataSamples.GetSampleFile("TestPackageCommon.docx");
            FileHelper.CopyFile(origFile.FullName, tempFile);

            // Open the temp file
            OPCPackage p = OPCPackage.Open(tempFile, PackageAccess.READ_WRITE);

            // Save it to a different file
            p.Save(tempFile2);
            p.Close();

            // Delete both the files
            File.Delete(tempFile);
            File.Delete(tempFile2);
        }

        private static ContentTypeManager GetContentTypeManager(OPCPackage pkg)
        {
            FieldInfo f = typeof(OPCPackage).GetField("contentTypeManager", BindingFlags.NonPublic | BindingFlags.Instance);
            //f.SetAccessible(true);
            return (ContentTypeManager)f.GetValue(pkg);
        }
        [Test]
        public void TestGetPartsByName()
        {
            String filepath = OpenXml4NetTestDataSamples.GetSampleFileName("sample.docx");

            OPCPackage pkg = OPCPackage.Open(filepath, PackageAccess.READ_WRITE);
            try
            {
                List<PackagePart> rs = pkg.GetPartsByName(new Regex("^/word/.*?\\.xml$"));
                Dictionary<String, PackagePart> selected = new Dictionary<String, PackagePart>();

                foreach (PackagePart p in rs)
                    selected.Add(p.PartName.Name, p);

                Assert.AreEqual(6, selected.Count);
                Assert.IsTrue(selected.ContainsKey("/word/document.xml"));
                Assert.IsTrue(selected.ContainsKey("/word/fontTable.xml"));
                Assert.IsTrue(selected.ContainsKey("/word/settings.xml"));
                Assert.IsTrue(selected.ContainsKey("/word/styles.xml"));
                Assert.IsTrue(selected.ContainsKey("/word/theme/theme1.xml"));
                Assert.IsTrue(selected.ContainsKey("/word/webSettings.xml"));
            }
            finally
            {
                pkg.Revert();
            }

        }
        [Test]
        public void TestGetPartSize()
        {
            String filepath = OpenXml4NetTestDataSamples.GetSampleFileName("sample.docx");
            OPCPackage pkg = OPCPackage.Open(filepath, PackageAccess.READ);
            try
            {
                int checked1 = 0;
                foreach (PackagePart part in pkg.GetParts())
                {
                    // Can get the size of zip parts
                    if (part.PartName.Name.Equals("/word/document.xml"))
                    {
                        checked1++;
                        Assert.AreEqual(typeof(ZipPackagePart), part.GetType());
                        Assert.AreEqual(6031L, part.Size);
                    }
                    if (part.PartName.Name.Equals("/word/fontTable.xml"))
                    {
                        checked1++;
                        Assert.AreEqual(typeof(ZipPackagePart), part.GetType());
                        Assert.AreEqual(1312L, part.Size);
                    }

                    // But not from the others
                    if (part.PartName.Name.Equals("/docProps/core.xml"))
                    {
                        checked1++;
                        Assert.AreEqual(typeof(PackagePropertiesPart), part.GetType());
                        Assert.AreEqual(-1, part.Size);
                    }
                }
                // Ensure we actually found the parts we want to check
                Assert.AreEqual(3, checked1);
            }
            finally
            {
                pkg.Revert();
            }
        }
        [Test]
        public void TestReplaceContentType()
        {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("sample.xlsx");
            OPCPackage p = OPCPackage.Open(is1);

            ContentTypeManager mgr = GetContentTypeManager(p);

            Assert.True(mgr.IsContentTypeRegister("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
            Assert.False(mgr.IsContentTypeRegister("application/vnd.ms-excel.sheet.macroEnabled.main+xml"));

            Assert.True(
                    p.ReplaceContentType(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
            );

            Assert.False(mgr.IsContentTypeRegister("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
            Assert.True(mgr.IsContentTypeRegister("application/vnd.ms-excel.sheet.macroEnabled.main+xml"));
        }
    }
}






