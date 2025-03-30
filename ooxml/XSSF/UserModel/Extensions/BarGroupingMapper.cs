using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;
using System;

namespace NPOI.XSSF.UserModel.Extensions
{
    internal static class BarGroupingMapper
    {
        public static ST_BarGrouping ToST_BarGrouping(this BarGrouping barGrouping)
        {
            return barGrouping switch {
                BarGrouping.Clustered => ST_BarGrouping.clustered,
                BarGrouping.Standard => ST_BarGrouping.standard,
                BarGrouping.PercentStacked => ST_BarGrouping.percentStacked,
                BarGrouping.Stacked => ST_BarGrouping.stacked,
                _ => throw new NotSupportedException()
            };
        }
    }
}