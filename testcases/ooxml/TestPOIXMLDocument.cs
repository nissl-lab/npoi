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

namespace NPOI.OOXML
{
    using System;
    using NPOI.Util;
    using System.Collections.Generic;
    using NUnit.Framework;
    using System.IO;
    using NPOI.OpenXml4Net.OPC;
    using NPOI;
    using TestCases;

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
            public override POIXMLDocumentPart CreateDocumentPart(POIXMLDocumentPart parent, PackageRelationship rel, PackagePart part)
            {
                return new POIXMLDocumentPart(part, rel);
            }

            public override POIXMLDocumentPart CreateDocumentPart(POIXMLRelation descriptor)
            {
                throw new NotSupportedException();
            }

        }

        /**
         * Recursively Traverse a OOXML document and assert that same logical parts have the same physical instances
         */
        private void Traverse(POIXMLDocumentPart part, Dictionary<String, POIXMLDocumentPart> context)
        {
            context[part.GetPackageRelationship().TargetUri.ToString()] = part;
            foreach (POIXMLDocumentPart p in part.GetRelations())
            {
                Assert.IsNotNull(p.ToString());
                String uri = p.GetPackageRelationship().TargetUri.ToString();
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

            OPCPackage pkg2 = OPCPackage.Open(tmp);
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
                pkg2.Revert();
            }
        }

        [Test]
        public void TestPPTX()
        {
            AssertReadWrite(
                    PackageHelper.Open(POIDataSamples.GetSlideShowInstance().OpenResourceAsStream("PPTWithAttachments.pptm"))
            );
        }
        [Test]
        public void TestXLSX()
        {
            AssertReadWrite(
                    PackageHelper.Open(POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("ExcelWithAttachments.xlsm"))
                    );
        }
        [Test]
        public void TestDOCX()
        {
            AssertReadWrite(
                    PackageHelper.Open(POIDataSamples.GetDocumentInstance().OpenResourceAsStream("WordWithAttachments.docx"))
                    );
        }
        [Test]
        public void TestRelationOrder()
        {
            OPCPackage pkg = PackageHelper.Open(POIDataSamples.GetDocumentInstance().OpenResourceAsStream("WordWithAttachments.docx"));
            OPCParser doc = new OPCParser(pkg);
            doc.Parse(new TestFactory());

            foreach (POIXMLDocumentPart rel in doc.GetRelations())
            {
                //TODO finish me
                Assert.IsNotNull(rel);
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



