using System;
using System.Text;
using System.Collections.Generic;

using NUnit.Framework;using NUnit.Framework.Legacy;
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
            ClassicAssert.AreEqual(1 << 8, v.MaxColumns);
            ClassicAssert.AreEqual(v.MaxColumns - 1, v.LastColumnIndex);
            ClassicAssert.AreEqual(1 << 16, v.MaxRows);
            ClassicAssert.AreEqual(v.MaxRows - 1, v.LastRowIndex);
            ClassicAssert.AreEqual(30, v.MaxFunctionArgs);
            ClassicAssert.AreEqual(3, v.MaxConditionalFormats);
            ClassicAssert.AreEqual("IV", v.LastColumnName);
        }
        [Test]
        public void TestExcel2007()
        {
            SpreadsheetVersion v = SpreadsheetVersion.EXCEL2007;
            ClassicAssert.AreEqual(1 << 14, v.MaxColumns);
            ClassicAssert.AreEqual(v.MaxColumns - 1, v.LastColumnIndex);
            ClassicAssert.AreEqual(1 << 20, v.MaxRows);
            ClassicAssert.AreEqual(v.MaxRows - 1, v.LastRowIndex);
            ClassicAssert.AreEqual(255, v.MaxFunctionArgs);
            ClassicAssert.AreEqual(Int32.MaxValue, v.MaxConditionalFormats);
            ClassicAssert.AreEqual("XFD", v.LastColumnName);
        }
    }
}
