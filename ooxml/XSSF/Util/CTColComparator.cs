using System;
using System.Collections.Generic;
using System.Text;
using NPOI.OpenXmlFormats.Spreadsheet;

namespace NPOI.XSSF.Util
{
    public class CTColComparator:Comparer<CT_Col>
    {
        public override int Compare(CT_Col o1, CT_Col o2)
        {
            if (o1.min < o2.min)
            {
                return -1;
            }
            else if (o1.min > o2.min)
            {
                return 1;
            }
            else
            {
                if (o1.max < o2.max)
                {
                    return -1;
                }
                if (o1.max > o2.max)
                {
                    return 1;
                }
                return 0;
            }
        }
    }
}
