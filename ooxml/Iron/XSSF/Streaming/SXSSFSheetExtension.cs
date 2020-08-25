using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.XSSF.Streaming
{
    public partial class SXSSFSheet
    {
        public int LastPhysicalRowNumber
        {
            get
            {
                return _rows.Max(r => r.Value.RowNum);
            }
        }
    }
}
