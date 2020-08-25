using System.Linq;

namespace NPOI.XSSF.UserModel
{
    public partial class XSSFSheet
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
