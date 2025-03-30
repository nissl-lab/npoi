using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;
using NPOI.XSSF.UserModel.Extensions;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace TestCases.XSSF.UserModel.Extensions
{
    [TestFixture]
    public sealed class TestBarGroupingMapper
    {
        [TestCase(BarGrouping.Clustered, ST_BarGrouping.clustered)]
        [TestCase(BarGrouping.Stacked, ST_BarGrouping.stacked)]
        [TestCase(BarGrouping.Standard, ST_BarGrouping.standard)]
        [TestCase(BarGrouping.PercentStacked, ST_BarGrouping.percentStacked)]
        public void TestMappingBarGroupingToST_BarGrouping(BarGrouping barGrouping, ST_BarGrouping expected)
        {
            ST_BarGrouping actual = barGrouping.ToST_BarGrouping();
            
            ClassicAssert.AreEqual(expected, actual);
        }
    }
}