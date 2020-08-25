using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
