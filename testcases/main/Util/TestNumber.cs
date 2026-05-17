using NPOI.Util;
using NUnit.Framework;

namespace TestCases.Util
{
    internal class TestNumber
    {
        [TestCase(1, ExpectedResult = 1)]
        [TestCase(16, ExpectedResult = 1)]
        [TestCase(31, ExpectedResult = 5)]
        [TestCase(-1, ExpectedResult = 32)]
        [TestCase(int.MaxValue, ExpectedResult = 31)]
        [TestCase(int.MinValue, ExpectedResult = 1)]
        public static int TestBitCount(int value) => Number.BitCount(value);
    }
}
