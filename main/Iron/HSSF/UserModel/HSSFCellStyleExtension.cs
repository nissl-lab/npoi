namespace NPOI.HSSF.UserModel
{
    public partial class HSSFCellStyle
    {
        public bool IsBuiltInStyle
        {
            get
            {
                var sr = _workbook.GetStyleRecord(index);

                return sr == null || sr.IsBuiltin;
            }
        }
    }
}
