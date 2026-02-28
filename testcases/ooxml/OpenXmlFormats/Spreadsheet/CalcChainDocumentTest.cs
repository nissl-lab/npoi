using System.IO;
using NUnit.Framework;
using NPOI.OpenXmlFormats.Spreadsheet;

namespace ooxml.Testcases
{
    /// <summary>
    ///This is a test class for serialization and deserialization of the CT_CalcChain and CT_CalcCell.
    ///</summary>
    [TestFixture]
    public class CalcChainDocumentTest
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
        ///A test for the Serialization of CT_CalcChain.
        ///</summary>
        [Test]
        public void SerializeCalcChainDocumentTest()
        {
            var calcChain = new CT_CalcChain();
            {
                var cell1 = new CT_CalcCell();
                cell1.r = "E1";
                cell1.i = 1;
                calcChain.AddC(cell1);
            }
            {
                var cell1 = new CT_CalcCell();
                cell1.r = "D1";
                calcChain.AddC(cell1);
            }
            {
                var cell1 = new CT_CalcCell();
                cell1.r = "C1";
                calcChain.AddC(cell1);
            }

            using (StringWriter stream = new StringWriter())
            {
                CT_CalcChain_Accessor.serializer.Serialize(stream, calcChain, CT_CalcChain_Accessor.namespaces);
                string expected = @"<?xml version=""1.0"" encoding=""utf-16""?>
<calcChain xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <c r=""E1"" i=""1"" />
  <c r=""D1"" />
  <c r=""C1"" />
</calcChain>";
                Assert.AreEqual(expected, stream.ToString());
            }
        }


        /// <summary>
        ///A test for Deserialize
        ///</summary>
        [Test]
        public void DeserializeCalcChainDocumentTest()
        {
            // The following input is the excerpt of the Excel file (49966.xlsx).
            string input =
@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<calcChain xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><c r=""E1"" i=""1""/><c r=""D1""/><c r=""C1""/></calcChain>";
            CT_CalcChain result;
            {
                using (StringReader stream = new StringReader(input))
                {
                    result = (CT_CalcChain)CT_CalcChain_Accessor.serializer.Deserialize(stream);
                }
            }
            Assert.AreEqual(3, result.c.Count);
            Assert.AreEqual("E1", result.c[0].r);
            Assert.AreEqual(1, result.c[0].i);

            Assert.AreEqual("D1", result.c[1].r);
            Assert.AreEqual(0, result.c[1].i); // the default value
            Assert.AreEqual(false, result.c[1].iSpecified);

            Assert.AreEqual("C1", result.c[2].r);
            Assert.AreEqual(0, result.c[2].i); // the default value
            Assert.AreEqual(false, result.c[2].iSpecified);
        }



    }
}
