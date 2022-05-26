using NPOI.SS.UserModel;

namespace NPOI.XSSF.UserModel
{
    public partial class XSSFCellStyle : ICellStyle
    {
        public bool IsBuiltInStyle => Index == 0;
    }
}
