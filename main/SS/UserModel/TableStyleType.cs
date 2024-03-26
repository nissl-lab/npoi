using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public enum TableStyleType
    {
        wholeTable,
        pageFieldLabels,// pivot only
        pageFieldValues,// pivot only
        firstColumnStripe,
        secondColumnStripe,
        firstRowStripe,
        secondRowStripe,
        lastColumn,
        firstColumn,
        headerRow,
        totalRow,
        firstHeaderCell,
        lastHeaderCell,
        firstTotalCell,
        lastTotalCell,
        /* these are for pivot tables only */
        /// <summary>
        /// </summary>
        firstSubtotalColumn,
        /// <summary>
        /// </summary>
        secondSubtotalColumn,
        /// <summary>
        /// </summary>
        thirdSubtotalColumn,
        /// <summary>
        /// </summary>
        blankRow,
        /// <summary>
        /// </summary>
        firstSubtotalRow,
        /// <summary>
        /// </summary>
        secondSubtotalRow,
        /// <summary>
        /// </summary>
        thirdSubtotalRow,
        /// <summary>
        /// </summary>
        firstColumnSubheading,
        /// <summary>
        /// </summary>
        secondColumnSubheading,
        /// <summary>
        /// </summary>
        thirdColumnSubheading,
        /// <summary>
        /// </summary>
        firstRowSubheading,
        /// <summary>
        /// </summary>
        secondRowSubheading,
        /// <summary>
        /// </summary>
        thirdRowSubheading
    }
}
