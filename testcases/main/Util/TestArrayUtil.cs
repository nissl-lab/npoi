using System;
using System.Text;

using NPOI.Util;

using NUnit.Framework;

namespace TestCases.Util
{
    /// <summary>
    /// This class is used to test NPOI.Util.Arrays class
    /// </summary>
    [TestFixture]
    public class TestArrayUtil
    {
        public TestArrayUtil()
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

        [Test]
        public void TestCharFill()
        {
            char[] a=new char[100];
            Arrays.Fill(a, 'a');
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != 'a')
                {
                    Assert.Fail("incorrect default value");
                }
            }
        }

        [Test]
        public void TestByteFill()
        {
            byte[] a = new byte[100];
            Arrays.Fill(a, 0x01);
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != 0x01)
                {
                    Assert.Fail("incorrect default value");
                }
            }
        }

        [Test]
        public void TestIntFill()
        {
            int[] a = new int[100];
            Arrays.Fill(a, 5);
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != 5)
                {
                    Assert.Fail("incorrect default value");
                }
            }
        }
        [Test]
        public void TestEquals()
        {
            byte[] a = { 1, 2, 3, 4, 0x56, 0x2A,0xff };
            byte[] b = { 1, 2, 3, 4, 0x56, 0x2A, 0xff };
            byte[] c = { 1, 0, 3, 4, 0x56, 0x2B, 0xff };

            Assert.IsTrue(Arrays.Equals(a, b));
            Assert.IsFalse(Arrays.Equals(a, c));
        }
    }
}
