using NPOI.OpenXmlFormats.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ooxml.Testcases
{
    
    
    /// <summary>
    ///This is a test class for CT_FontTest and is intended
    ///to contain all CT_FontTest Unit Tests
    ///</summary>
    [TestClass()]
    public class CT_FontTest
    {


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
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for Parse
        ///</summary>
        [TestMethod()]
        public void ParseTest()
        {
            string fontName = "Tahoma";
            CT_Font expected = new CT_Font();
            var name = expected.AddNewName();
            name.val = fontName;
            var b = expected.AddNewB();
            b.val = true;
            string firstXml = CT_Font.GetString(expected);

            CT_Font actual = CT_Font.Parse(firstXml);
            string secondXml = CT_Font.GetString(actual);
            Assert.AreEqual(firstXml, secondXml);
            Assert.AreEqual(expected.name.val, actual.name.val);
            Assert.AreEqual(fontName, actual.name.val);           
            Assert.AreEqual(expected.nameSpecified, actual.nameSpecified);
            Assert.AreEqual(expected.b.val, actual.b.val);
            Assert.AreEqual(true, actual.b.val);
            Assert.AreEqual(expected.bSpecified, actual.bSpecified);
            Assert.AreEqual(expected.iSpecified, actual.iSpecified);
            Assert.AreEqual(false, actual.iSpecified);
        }
    }
}
