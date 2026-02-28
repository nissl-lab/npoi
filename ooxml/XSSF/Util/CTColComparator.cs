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
    
        public static IComparer<CT_Col> BY_MAX = new CTColComparatorByMax();
        public static IComparer<CT_Col> BY_MIN_MAX = new CTColComparatorByMinMax();
        public class CTColComparatorByMinMax : CTColComparator
        {
            public override int Compare(CT_Col col1, CT_Col col2)
            {
                long col11min = col1.min;
                long col2min = col2.min;
                return col11min < col2min ? -1 : col11min > col2min ? 1 : BY_MAX.Compare(col1, col2);
            }
        }
        public class CTColComparatorByMax : CTColComparator
        {
            public override int Compare(CT_Col col1, CT_Col col2)
            {
                long col1max = col1.max;
                long col2max = col2.max;
                return col1max < col2max ? -1 : col1max > col2max ? 1 : 0;
            }
        }
    }
}
