using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula
{
    public class DataValidationEvaluator
    {
        /// <summary>
        /// Note that this assumes the cell cached value is up to date and in sync with data edits
        /// </summary>
        /// <param name="cell">The <see cref="Cell"/> to check.</param>
        /// <param name="type">The <see cref="CellType"/> to check for.</param>
        /// <returns>true if the cell or cached cell formula result type match the given type</returns>
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
