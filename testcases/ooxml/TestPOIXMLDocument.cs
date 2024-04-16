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

namespace TestCases.OOXML
{
    using NPOI;
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text.RegularExpressions;
    using TestCases;
    using TestCases.Util;

    /**
     * Test recursive read and write of OPC namespaces
     */
    [TestFixture]
    public class TestPOIXMLDocument 
    {

        private class OPCParser : POIXMLDocument
        {

            public OPCParser(OPCPackage pkg)
                : base(pkg)
            {

            }

            public override List<PackagePart> GetAllEmbedds()
            {
                throw new NotSupportedException();
            }

            public void Parse(POIXMLFactory factory)
            {
                Load(factory);
            }
        }

        private class TestFactory : POIXMLFactory
        {

            public TestFactory()
            {
                //
            }
            protected override POIXMLRelation GetDescriptor(String relationshipType)
            {
                return null;
            }
            protected override POIXMLDocumentPart CreateDocumentPart(Type cls, Type[] classes, Object[] values)
            {
                return null;
            }

        }

        /**
         * Recursively Traverse a OOXML document and assert that same logical parts have the same physical instances
         */
        private void Traverse(POIXMLDocumentPart part, Dictionary<String, POIXMLDocumentPart> context)
        {
            Assert.AreEqual(part.GetPackageRelationship().TargetUri.ToString(), part.GetPackagePart().PartName.Name);

            context[part.GetPackagePart().PartName.Name] = part;
            foreach (POIXMLDocumentPart p in part.GetRelations())
            {
                Assert.IsNotNull(p.ToString());
                String uri = p.GetPackagePart().PartName.URI.ToString();
                Assert.AreEqual(uri, p.GetPackageRelationship().TargetUri.ToString());

                if (!context.ContainsKey(uri))
                {
                    Traverse(p, context);
                }
                else
                {
                    POIXMLDocumentPart prev = context[uri];
                    Assert.AreSame(prev, p, "Duplicate POIXMLDocumentPart instance for targetURI=" + uri);
                }
            }
        }

        public void AssertReadWrite(OPCPackage pkg1)
        {

            OPCParser doc = new OPCParser(pkg1);
            doc.Parse(new TestFactory());

            Dictionary<String, POIXMLDocumentPart> context = new Dictionary<String, POIXMLDocumentPart>();
            Traverse(doc, context);
            context.Clear();

            string tmp = TempFile.GetTempFilePath("poi-ooxml", ".tmp");
            FileStream out1 = new FileStream(tmp, FileMode.CreateNew);
            doc.Write(out1);
            out1.Close();

            // Should not be able to write to an output stream that has been closed
            try
            {
                doc.Write(out1);
                Assert.Fail("Should not be able to write to an output stream that has been closed.");
            }
            catch (OpenXML4NetRuntimeException e) {
                //OpenXml4NetRuntimeException
                // FIXME: A better exception class (IOException?) and message should be raised
                // indicating that the document could not be written because the output stream is closed.
                // see {@link org.apache.poi.openxml4j.opc.ZipPackage#saveImpl(java.io.OutputStream)}
                if (Regex.IsMatch(e.Message, "Fail to save: an error occurs while saving the package : Must support writing.+"))
                {
                    // expected
                }
                else
                {
                    throw e;
                }
            }

            // Should not be able to write a document that has been closed
            doc.Close();
            try
            {
                doc.Write(new NullOutputStream());
                Assert.Fail("Should not be able to write a document that has been closed.");
            }
            catch (IOException e) {
                if (e.Message.Equals("Cannot write data, document seems to have been closed already"))
                {
                    // expected
                }
                else
                {
                    throw e;
                }
            }

            // Should be able to close a document multiple times, though subsequent closes will have no effect.
            doc.Close();

            OPCPackage pkg2 = OPCPackage.Open(tmp);
            doc = new OPCParser(pkg1);
            try
            {
                doc = new OPCParser(pkg1);
                doc.Parse(new TestFactory());
                context = new Dictionary<String, POIXMLDocumentPart>();
                Traverse(doc, context);
                context.Clear();

                Assert.AreEqual(pkg1.Relationships.Size, pkg2.Relationships.Size);

                List<PackagePart> l1 = pkg1.GetParts();
                List<PackagePart> l2 = pkg2.GetParts();

                Assert.AreEqual(l1.Count, l2.Count);
                for (int i = 0; i < l1.Count; i++)
                {
                    PackagePart p1 = l1[i];
                    PackagePart p2 = l2[i];

                    Assert.AreEqual(p1.ContentType, p2.ContentType);
                    Assert.AreEqual(p1.HasRelationships, p2.HasRelationships);
                    if (p1.HasRelationships)
                    {
                        Assert.AreEqual(p1.Relationships.Size, p2.Relationships.Size);
                    }
                    Assert.AreEqual(p1.PartName, p2.PartName);
                }
            }
            finally
            {
                doc.Close();
                pkg1.Close();
                pkg2.Close();
            }
        }

        [Test]
        public void TestPPTX()
        {
            POIDataSamples pds = POIDataSamples.GetSlideShowInstance();
            AssertReadWrite(
                    PackageHelper.Open(pds.OpenResourceAsStream("PPTWithAttachments.pptm"))
            );
        }
        [Test]
        public void TestXLSX()
        {
            POIDataSamples pds = POIDataSamples.GetSpreadSheetInstance();
            AssertReadWrite(
                    PackageHelper.Open(pds.OpenResourceAsStream("ExcelWithAttachments.xlsm"))
                    );
        }
        [Test]
        public void TestDOCX()
        {
            POIDataSamples pds = POIDataSamples.GetDocumentInstance();
            AssertReadWrite(
                    PackageHelper.Open(pds.OpenResourceAsStream("WordWithAttachments.docx"))
                    );
        }
        [Test]
        public void TestRelationOrder()
        {
            POIDataSamples pds = POIDataSamples.GetDocumentInstance();
            OPCPackage pkg = PackageHelper.Open(pds.OpenResourceAsStream("WordWithAttachments.docx"));
            OPCParser doc = new OPCParser(pkg);
            try
            {
                doc.Parse(new TestFactory());

                foreach (POIXMLDocumentPart rel in doc.GetRelations())
                {
                    //TODO finish me
                    Assert.IsNotNull(rel);
                }
            }
            finally
            {
                doc.Close();
            }

        }
        [Test]
        public void TestGetNextPartNumber()
        {

            POIDataSamples pds = POIDataSamples.GetDocumentInstance();
            OPCPackage pkg = PackageHelper.Open(pds.OpenResourceAsStream("WordWithAttachments.docx"));
            OPCParser doc = new OPCParser(pkg);
            try
            {
                doc.Parse(new TestFactory());

                // Non-indexed parts: Word is taken, Excel is not
                Assert.AreEqual(-1, doc.GetNextPartNumber(XWPFRelation.DOCUMENT, 0));
                Assert.AreEqual(-1, doc.GetNextPartNumber(XWPFRelation.DOCUMENT, -1));
                Assert.AreEqual(-1, doc.GetNextPartNumber(XWPFRelation.DOCUMENT, 99));
                Assert.AreEqual(0, doc.GetNextPartNumber(XSSFRelation.WORKBOOK, 0));
                Assert.AreEqual(0, doc.GetNextPartNumber(XSSFRelation.WORKBOOK, -1));
                Assert.AreEqual(0, doc.GetNextPartNumber(XSSFRelation.WORKBOOK, 99));

                // Indexed parts:
                // Has 2 headers
                Assert.AreEqual(0, doc.GetNextPartNumber(XWPFRelation.HEADER, 0));
                Assert.AreEqual(3, doc.GetNextPartNumber(XWPFRelation.HEADER, -1));
                Assert.AreEqual(3, doc.GetNextPartNumber(XWPFRelation.HEADER, 1));
                Assert.AreEqual(8, doc.GetNextPartNumber(XWPFRelation.HEADER, 8));

                // Has no Excel Sheets
                Assert.AreEqual(0, doc.GetNextPartNumber(XSSFRelation.WORKSHEET, 0));
                Assert.AreEqual(1, doc.GetNextPartNumber(XSSFRelation.WORKSHEET, -1));
                Assert.AreEqual(1, doc.GetNextPartNumber(XSSFRelation.WORKSHEET, 1));
            }
            finally
            {
                doc.Close();
            }
        }
        [Test]
        public void TestCommitNullPart()
        {
            POIXMLDocumentPart part = new POIXMLDocumentPart();
            part.PrepareForCommit();
            part.Commit();
            part.OnSave(new List<PackagePart>());

            Assert.IsNull(part.GetRelationById(null));
            Assert.IsNull(part.GetRelationId(null));
            Assert.IsFalse(part.RemoveRelation(null, true));
            part.RemoveRelation(null);
            Assert.AreEqual(string.Empty, part.ToString());
            part.OnDocumentCreate();
            //part.GetTargetPart(null);
        }
    }
}



