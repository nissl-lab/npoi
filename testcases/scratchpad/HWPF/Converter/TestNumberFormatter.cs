namespace TestCases.HWPF.Converter
{
    using NPOI.HWPF;
    using NPOI.HWPF.UserModel;
    using System;
    
    using NPOI.HWPF.Converter;
    using NUnit.Framework;
    [TestFixture]
    public class TestNumberFormatter
    {
        [Test]
        public void TestRoman()
        {
            Assert.AreEqual("i", NumberFormatter.GetNumber(1, 2));
            Assert.AreEqual("ii", NumberFormatter.GetNumber(2, 2));
            Assert.AreEqual("iii", NumberFormatter.GetNumber(3, 2));
            Assert.AreEqual("iv", NumberFormatter.GetNumber(4, 2));
            Assert.AreEqual("v", NumberFormatter.GetNumber(5, 2));
            Assert.AreEqual("vi", NumberFormatter.GetNumber(6, 2));
            Assert.AreEqual("vii", NumberFormatter.GetNumber(7, 2));
            Assert.AreEqual("viii", NumberFormatter.GetNumber(8, 2));
            Assert.AreEqual("ix", NumberFormatter.GetNumber(9, 2));
            Assert.AreEqual("x", NumberFormatter.GetNumber(10, 2));

            Assert.AreEqual("mdcvi", NumberFormatter.GetNumber(1606, 2));
            Assert.AreEqual("mcmx", NumberFormatter.GetNumber(1910, 2));
            Assert.AreEqual("mcmliv", NumberFormatter.GetNumber(1954, 2));
        }
        [Test]
        public void TestEnglish()
        {
            Assert.AreEqual("a", NumberFormatter.GetNumber(1, 4));
            Assert.AreEqual("z", NumberFormatter.GetNumber(26, 4));

            Assert.AreEqual("aa", NumberFormatter.GetNumber(1 * 26 + 1, 4));
            Assert.AreEqual("az", NumberFormatter.GetNumber(1 * 26 + 26, 4));

            Assert.AreEqual("za", NumberFormatter.GetNumber(26 * 26 + 1, 4));
            Assert.AreEqual("zz", NumberFormatter.GetNumber(26 * 26 + 26, 4));

            Assert.AreEqual("aaa",
                       NumberFormatter.GetNumber(26 * 26 + 1 * 26 + 1, 4));
            Assert.AreEqual("aaz",
                       NumberFormatter.GetNumber(26 * 26 + 1 * 26 + 26, 4));

            Assert.AreEqual("aba",
                       NumberFormatter.GetNumber(1 * 26 * 26 + 2 * 26 + 1, 4));
            Assert.AreEqual("aza",
                       NumberFormatter.GetNumber(1 * 26 * 26 + 26 * 26 + 1, 4));

            Assert.AreEqual("azz",
                       NumberFormatter.GetNumber(26 * 26 + 26 * 26 + 26, 4));
            Assert.AreEqual("baa",
                       NumberFormatter.GetNumber(2 * 26 * 26 + 1 * 26 + 1, 4));
            Assert.AreEqual("zaa",
                       NumberFormatter.GetNumber(26 * 26 * 26 + 1 * 26 + 1, 4));
            Assert.AreEqual("zzz",
                       NumberFormatter.GetNumber(26 * 26 * 26 + 26 * 26 + 26, 4));

            Assert.AreEqual(
                       "aaaa",
                       NumberFormatter.GetNumber(1 * 26 * 26 * 26 + 1 * 26 * 26 + 1
                               * 26 + 1, 4));
            Assert.AreEqual(
                       "azzz",
                       NumberFormatter.GetNumber(1 * 26 * 26 * 26 + 26 * 26 * 26 + 26
                               * 26 + 26, 4));
            Assert.AreEqual(
                       "zzzz",
                       NumberFormatter.GetNumber(26 * 26 * 26 * 26 + 26 * 26 * 26
                               + 26 * 26 + 26, 4));

            for (int i = 1; i < 1000000; i++)
            {
                // make sure there is no exceptions
                NumberFormatter.GetNumber(i, 4);
            }
        }
    }
}
