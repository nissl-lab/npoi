using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.XSSF.Streaminging;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFEvaluationSheet : XSSFEvaluationSheet
    {
        private SXSSFSheet _xs;

    public SXSSFEvaluationSheet(SXSSFSheet sheet)
        {
            _xs = sheet;
        }

        public SXSSFSheet getSXSSFSheet()
        {
            return _xs;
        }

    public SXSSFEvaluationCell getCell(int rowIndex, int columnIndex)
        {
            //TODO: maybe why we want a sorted dict.
            SXSSFRow row = _xs._rows[rowIndex];
            if (row == null)
            {
                if (rowIndex <= _xs.lastFlushedRowNumber)
                {
                    throw new RowFlushedException(rowIndex);
                }
                return null;
            }
            SXSSFCell cell = (SXSSFCell)row.Cells[columnIndex];
            if (cell == null)
            {
                return null;
            }
            return new SXSSFEvaluationCell(cell, this);
        }

        /* (non-JavaDoc), inherit JavaDoc from EvaluationSheet
         * @since POI 3.15 beta 3
         */

    public void clearAllCachedResultValues()
        {
            // nothing to do
        }
    }
}
