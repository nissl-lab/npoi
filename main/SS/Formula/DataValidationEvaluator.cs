using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula
{
    public class DataValidationEvaluator
    {
        /**
    * Note that this assumes the cell cached value is up to date and in sync with data edits
     *
    * @param cell The {@link Cell} to check.
    * @param type The {@link CellType} to check for.
    * @return true if the cell or cached cell formula result type match the given type
    */
        public static bool IsType(ICell cell, CellType type)
        {
            CellType cellType = cell.CellType;
            return cellType == type
                  || (cellType == CellType.Formula
                      && cell.CachedFormulaResultType == type
                     );
        }

    }
}
