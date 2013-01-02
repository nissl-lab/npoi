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
        public TestContentTypeManager()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

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
    }
}
