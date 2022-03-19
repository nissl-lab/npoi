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
        /***/
        firstSubtotalColumn,
        /***/
        secondSubtotalColumn,
        /***/
        thirdSubtotalColumn,
        /***/
        blankRow,
        /***/
        firstSubtotalRow,
        /***/
        secondSubtotalRow,
        /***/
        thirdSubtotalRow,
        /***/
        firstColumnSubheading,
        /***/
        secondColumnSubheading,
        /***/
        thirdColumnSubheading,
        /***/
        firstRowSubheading,
        /***/
        secondRowSubheading,
        /***/
        thirdRowSubheading
    }
}
