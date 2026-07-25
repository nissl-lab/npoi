using NPOI.Util;
using NUnit.Framework;

namespace TestCases.Util
{
    [TestFixture]
    internal static class TestNumber
    {
        [TestCase(0, ExpectedResult = 0)]
        [TestCase(1, ExpectedResult = 1)]
        [TestCase(16, ExpectedResult = 1)]
        [TestCase(31, ExpectedResult = 5)]
        [TestCase(-1, ExpectedResult = 32)]
        [TestCase(int.MaxValue, ExpectedResult = 31)]
        [TestCase(int.MinValue, ExpectedResult = 1)]
        public static int TestBitCount(int value) => Number.BitCount(value);
    }
}
