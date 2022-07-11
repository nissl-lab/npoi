using NPOI.SS.UserModel;

namespace NPOI.XSSF.UserModel
{
    public partial class XSSFCell : ICell
    {
        /// <summary>
        /// By default NPOI will return new style if there's no style attributed to a cell,
        /// it doesn't consider that even if the cell has no style, the column that it belongs to might
        /// have a style. This method will check first if there is a specific style for the cell
        /// and if there isn't will check for a style in the column. It will attribute that style to a cell then.
        /// </summary>
        public void EnsureStyleConsideringColumnStyle()
        {
            if ((_stylesSource != null) && (_stylesSource.NumCellStyles > 0))
            {
                long idx = 0;

                if (_cell.IsSetS())
                {
                    idx = _cell.s;
                }
                else if (Sheet.GetColumnStyle(ColumnIndex) != null)
                {
                    idx = Sheet.GetColumnStyle(ColumnIndex).Index;
                }

                if (_cell.IsSetS())
                {
                    _cell.unsetS();
                }

                _cell.s = (uint)idx;
            }
        }
    }
}
