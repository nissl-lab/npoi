using System;
using System.Text;
using System.Collections.Generic;

using NUnit.Framework;
using NPOI.SS;

namespace TestCases.SS
{
    /// <summary>
    /// Summary description for TestSpreadsheetVersion
    /// </summary>
    [TestFixture]
    public class TestSpreadsheetVersion
    {
        public TestSpreadsheetVersion()
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
        // [SetUp]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [Test]
        public void TestExcel97()
        {
            SpreadsheetVersion v = SpreadsheetVersion.EXCEL97;
            Assert.AreEqual(1 << 8, v.MaxColumns);
            Assert.AreEqual(v.MaxColumns - 1, v.LastColumnIndex);
            Assert.AreEqual(1 << 16, v.MaxRows);
            Assert.AreEqual(v.MaxRows - 1, v.LastRowIndex);
            Assert.AreEqual(30, v.MaxFunctionArgs);
            Assert.AreEqual(3, v.MaxConditionalFormats);
            Assert.AreEqual("IV", v.LastColumnName);
        }
        [Test]
        public void TestExcel2007()
        {
            SpreadsheetVersion v = SpreadsheetVersion.EXCEL2007;
            Assert.AreEqual(1 << 14, v.MaxColumns);
            Assert.AreEqual(v.MaxColumns - 1, v.LastColumnIndex);
            Assert.AreEqual(1 << 20, v.MaxRows);
            Assert.AreEqual(v.MaxRows - 1, v.LastRowIndex);
            Assert.AreEqual(255, v.MaxFunctionArgs);
            Assert.AreEqual(Int32.MaxValue, v.MaxConditionalFormats);
            Assert.AreEqual("XFD", v.LastColumnName);
        }
    }
}
