using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using System;

namespace NPOI.OOXML.XSSF.UserModel
{
    public class XSSFConditionFilterData:IConditionFilterData
    {
        private readonly CT_CfRule _cfRule;
        public XSSFConditionFilterData(CT_CfRule cfRule)
        {
            _cfRule = cfRule;
        }

        public bool AboveAverage => _cfRule.aboveAverage;

        public bool Bottom => _cfRule.bottom;

        public bool EqualAverage => _cfRule.equalAverage;

        public bool Percent => _cfRule.percent;

        public long Rank => _cfRule.rank;

        public int StdDev => _cfRule.stdDev;
    }
}
