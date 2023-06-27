using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.XSSF.UserModel
{
    public partial class XSSFSheet
    {
        public int LastPhysicalRowNumber
        {
            get
            {
                return _rows.Count == 0
                    ? 0
                    :_rows.Max(r => r.Value.RowNum);
            }
        }

        IEnumerator<IRow> IEnumerable<IRow>.GetEnumerator()
        {
            return _rows.Values.GetEnumerator();
        }
    }
}
