using NPOI.HSSF.Model;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.HSSF.UserModel
{
    public partial class HSSFSheet
    {
        public int LastPhysicalRowNumber
        {
            get
            {
                return rows.Max(r => r.Value.RowNum);
            }
        }

        IEnumerator<IRow> IEnumerable<IRow>.GetEnumerator()
        {
            return rows.Values.GetEnumerator();
        }

        public InternalSheet GetSheet()
        {
            return _sheet;
        }
    }
}
