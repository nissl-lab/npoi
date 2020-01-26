using System;
using System.Text;
using System.Collections.Generic;
using NUnit.Framework;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.OpenXml4Net.OPC;

namespace TestCases.OpenXml4Net.OPC.Internal
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class TestContentTypeManager
    {
        /**
         * Test the properties part content parsing.
         */
        [Test]
        public void TestContentType()
        {
            String filepath = OpenXml4NetTestDataSamples.GetSampleFileName("sample.docx");
            // Retrieves core properties part
            OPCPackage p = OPCPackage.Open(filepath, PackageAccess.READ);
            PackageRelationshipCollection rels = p.GetRelationshipsByType(PackageRelationshipTypes.CORE_PROPERTIES);
            PackageRelationship corePropertiesRelationship = rels.GetRelationship(0);
            PackagePart coreDocument = p.GetPart(corePropertiesRelationship);

            Assert.AreEqual("application/vnd.openxmlformats-package.core-properties+xml", coreDocument.ContentType);
            // TODO - finish writing this test
            Assume.That(false,"finish writing this test");

            ContentTypeManager ctm = new ZipContentTypeManager(coreDocument.GetInputStream(), p);
        }


        /**
         * Test the addition of several default and override content types.
         */
        [Test]
        public void TestContentTypeAddition()
        {
            ContentTypeManager ctm = new ZipContentTypeManager(null, null);

            PackagePartName name1 = PackagingUriHelper.CreatePartName("/foo/foo.XML");
            PackagePartName name2 = PackagingUriHelper.CreatePartName("/foo/foo2.xml");
            PackagePartName name3 = PackagingUriHelper.CreatePartName("/foo/doc.rels");
            PackagePartName name4 = PackagingUriHelper.CreatePartName("/foo/doc.RELS");

            // Add content types
            ctm.AddContentType(name1, "foo-type1");
            ctm.AddContentType(name2, "foo-type2");
            ctm.AddContentType(name3, "text/xml+rel");
            ctm.AddContentType(name4, "text/xml+rel");

            Assert.AreEqual(ctm.GetContentType(name1), "foo-type1");
            Assert.AreEqual(ctm.GetContentType(name2), "foo-type2");
            Assert.AreEqual(ctm.GetContentType(name3), "text/xml+rel");
            Assert.AreEqual(ctm.GetContentType(name3), "text/xml+rel");
        }
        /**
         * Test the addition then removal of content types.
         */
        [Test]
        public void TestContentTypeRemoval()
        {
            ContentTypeManager ctm = new ZipContentTypeManager(null, null);

            PackagePartName name1 = PackagingUriHelper.CreatePartName("/foo/foo.xml");
            PackagePartName name2 = PackagingUriHelper.CreatePartName("/foo/foo2.xml");
            PackagePartName name3 = PackagingUriHelper.CreatePartName("/foo/doc.rels");
            PackagePartName name4 = PackagingUriHelper.CreatePartName("/foo/doc.RELS");

            // Add content types
            ctm.AddContentType(name1, "foo-type1");
            ctm.AddContentType(name2, "foo-type2");
            ctm.AddContentType(name3, "text/xml+rel");
            ctm.AddContentType(name4, "text/xml+rel");
            ctm.RemoveContentType(name2);
            ctm.RemoveContentType(name3);

            Assert.AreEqual(ctm.GetContentType(name1), "foo-type1");
            Assert.AreEqual(ctm.GetContentType(name2), "foo-type1");
            Assert.AreEqual(ctm.GetContentType(name3), null);

            ctm.RemoveContentType(name1);
            Assert.AreEqual(ctm.GetContentType(name1), null);
            Assert.AreEqual(ctm.GetContentType(name2), null);
        }

        /**
         * Test the addition then removal of content types in a package.
         */
        [Ignore("")]
        [Test]
        public void TestContentTypeRemovalPackage()
        {
            // TODO
            Assert.Fail("test not written");
        }

    }
}
