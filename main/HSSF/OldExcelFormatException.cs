using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSSF
{
    public class OldExcelFormatException:Exception
    {
        public OldExcelFormatException(String s)
            : base(s)
        { }

    }
}
