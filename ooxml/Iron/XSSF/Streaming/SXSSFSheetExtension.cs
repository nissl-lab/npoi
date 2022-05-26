using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

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

        IEnumerator<IRow> IEnumerable<IRow>.GetEnumerator()
        {
            return ((IEnumerable<IRow>)_sh).GetEnumerator();
        }
    }
}
