/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using NPOI.SS.Util;
using System;

namespace NPOI.SS.UserModel
{
    /// <summary>
    /// <para>
    /// Ordered list of table style elements, for both data tables and pivot tables.
    /// Some elements only apply to pivot tables, but any style definition can omit any number,
    /// so having them in one list should not be a problem.
    /// </para>
    /// <para>
    /// The order is the specification order of application, with later elements overriding previous
    /// ones, if style properties conflict.
    /// </para>
    /// <para>
    /// Processing could iterate bottom-up if looking for specific properties, and stop when the
    /// first style is found defining a value for that property.
    /// </para>
    /// <para>
    /// Enum names match the OOXML spec values exactly, so <see cref="valueOf(String)" /> will work.
    /// </para>
    /// </summary>
    /// @since 3.17 beta 1
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

    public static class TableStyleTypeExtension
    {
        public static CellRangeAddressBase GetRange(this TableStyleType styleType, ITable table, ICell cell)
        {
            switch(styleType)
            {
                case TableStyleType.wholeTable:
                    return new CellRangeAddress(table.StartRowIndex, table.EndRowIndex, table.StartColIndex, table.EndColIndex);
                case TableStyleType.firstColumnStripe:
                    return GetFirstColumnStripeRange(table, cell);
                case TableStyleType.secondColumnStripe:
                    return GetSecondColumnStripeRange(table, cell);
                case TableStyleType.firstRowStripe:
                    return GetFirstRowStripeRange(table, cell);
                case TableStyleType.secondRowStripe:
                    return GetSecondRowStripeRange(table, cell);
                case TableStyleType.lastColumn:
                    if (! table.Style.IsShowLastColumn) return null;
                    return new CellRangeAddress(table.StartRowIndex, table.EndRowIndex, table.EndColIndex, table.EndColIndex);
                case TableStyleType.firstColumn:
                    if (! table.Style.IsShowFirstColumn) return null;
                    return new CellRangeAddress(table.StartRowIndex, table.EndRowIndex, table.StartColIndex, table.StartColIndex);
                case TableStyleType.headerRow:
                    if (table.HeaderRowCount < 1) return null;
                    return new CellRangeAddress(table.StartRowIndex, table.StartRowIndex + table.HeaderRowCount -1, table.StartColIndex, table.EndColIndex);
                case TableStyleType.totalRow:
                    if (table.TotalsRowCount < 1) return null;
                    return new CellRangeAddress(table.EndRowIndex - table.TotalsRowCount +1, table.EndRowIndex, table.StartColIndex, table.EndColIndex);
                case TableStyleType.firstHeaderCell:
                    if (table.HeaderRowCount < 1) return null;
                    return new CellRangeAddress(table.StartRowIndex, table.StartRowIndex, table.StartColIndex, table.StartColIndex);
                case TableStyleType.lastHeaderCell:
                    if (table.HeaderRowCount < 1) return null;
                    return new CellRangeAddress(table.StartRowIndex, table.StartRowIndex, table.EndColIndex, table.EndColIndex);
                case TableStyleType.firstTotalCell:
                    if (table.TotalsRowCount < 1) return null;
                    return new CellRangeAddress(table.EndRowIndex - table.TotalsRowCount +1, table.EndRowIndex, table.StartColIndex, table.StartColIndex);
                case TableStyleType.lastTotalCell:
                    if (table.TotalsRowCount < 1) return null;
                    return new CellRangeAddress(table.EndRowIndex - table.TotalsRowCount +1, table.EndRowIndex, table.EndColIndex, table.EndColIndex);
                default:
                    return null;
            }
        }

        private static CellRangeAddress GetFirstColumnStripeRange(ITable table, ICell cell)
        {
            ITableStyleInfo info = table.Style;
            if (! info.IsShowColumnStripes) return null;
            IDifferentialStyleProvider c1Style = info.Style.GetStyle(TableStyleType.firstColumnStripe);
            IDifferentialStyleProvider c2Style = info.Style.GetStyle(TableStyleType.secondColumnStripe);
            int c1Stripe = c1Style == null ? 1 : Math.Max(1, c1Style.StripeSize);
            int c2Stripe = c2Style == null ? 1 : Math.Max(1, c2Style.StripeSize);
            
            int firstStart = table.StartColIndex;
            int secondStart = firstStart + c1Stripe;
            int c = cell.ColumnIndex;
            
            // look for the stripe containing c, accounting for the style element stripe size
            // could do fancy math, but tables can't be that wide, a simple loop is fine
            // if not in this type of stripe, return null
            while (true) {
                if (firstStart > c) break;
                if (c >= firstStart && c <= secondStart -1) return new CellRangeAddress(table.StartRowIndex, table.EndRowIndex, firstStart, secondStart - 1);
                firstStart = secondStart + c2Stripe;
                secondStart = firstStart + c1Stripe;
            }
            return null;
        }

        private static CellRangeAddress GetSecondColumnStripeRange(ITable table, ICell cell)
        {
            ITableStyleInfo info = table.Style;
            if (! info.IsShowColumnStripes) return null;
            
            IDifferentialStyleProvider c1Style = info.Style.GetStyle(TableStyleType.firstColumnStripe);
            IDifferentialStyleProvider c2Style = info.Style.GetStyle(TableStyleType.secondColumnStripe);
            int c1Stripe = c1Style == null ? 1 : Math.Max(1, c1Style.StripeSize);
            int c2Stripe = c2Style == null ? 1 : Math.Max(1, c2Style.StripeSize);

            int firstStart = table.StartColIndex;
            int secondStart = firstStart + c1Stripe;
            int c = cell.ColumnIndex;
            
            // look for the stripe containing c, accounting for the style element stripe size
            // could do fancy math, but tables can't be that wide, a simple loop is fine
            // if not in this type of stripe, return null
            while (true) {
                if (firstStart > c) break;
                if (c >= secondStart && c <= secondStart + c2Stripe -1) return new CellRangeAddress(table.StartRowIndex, table.EndRowIndex, secondStart, secondStart + c2Stripe - 1);
                firstStart = secondStart + c2Stripe;
                secondStart = firstStart + c1Stripe;
            }
            return null;
        }

        private static CellRangeAddress GetFirstRowStripeRange(ITable table, ICell cell)
        {
            ITableStyleInfo info = table.Style;
            if (! info.IsShowRowStripes) return null;
            
            IDifferentialStyleProvider c1Style = info.Style.GetStyle(TableStyleType.firstRowStripe);
            IDifferentialStyleProvider c2Style = info.Style.GetStyle(TableStyleType.secondRowStripe);
            int c1Stripe = c1Style == null ? 1 : Math.Max(1, c1Style.StripeSize);
            int c2Stripe = c2Style == null ? 1 : Math.Max(1, c2Style.StripeSize);

            int firstStart = table.StartRowIndex + table.HeaderRowCount;
            int secondStart = firstStart + c1Stripe;
            int c = cell.RowIndex;
            
            // look for the stripe containing c, accounting for the style element stripe size
            // could do fancy math, but tables can't be that wide, a simple loop is fine
            // if not in this type of stripe, return null
            while (true) {
                if (firstStart > c) break;
                if (c >= firstStart && c <= secondStart -1)
                    return new CellRangeAddress(firstStart, secondStart - 1, table.StartColIndex, table.EndColIndex);
                firstStart = secondStart + c2Stripe;
                secondStart = firstStart + c1Stripe;
            }
            return null;
        }

        private static CellRangeAddress GetSecondRowStripeRange(ITable table, ICell cell)
        {
            ITableStyleInfo info = table.Style;
            if(!info.IsShowRowStripes)
                return null;

            IDifferentialStyleProvider c1Style = info.Style.GetStyle(TableStyleType.firstRowStripe);
            IDifferentialStyleProvider c2Style = info.Style.GetStyle(TableStyleType.secondRowStripe);
            int c1Stripe = c1Style == null ? 1 : Math.Max(1, c1Style.StripeSize);
            int c2Stripe = c2Style == null ? 1 : Math.Max(1, c2Style.StripeSize);

            int firstStart = table.StartRowIndex + table.HeaderRowCount;
            int secondStart = firstStart + c1Stripe;
            int c = cell.RowIndex;

            // look for the stripe containing c, accounting for the style element stripe size
            // could do fancy math, but tables can't be that wide, a simple loop is fine
            // if not in this type of stripe, return null
            while(true)
            {
                if(firstStart > c)
                    break;
                if(c >= secondStart && c <= secondStart +c2Stripe -1)
                    return new CellRangeAddress(secondStart, secondStart + c2Stripe - 1, table.StartColIndex, table.EndColIndex);
                firstStart = secondStart + c2Stripe;
                secondStart = firstStart + c1Stripe;
            }
            return null;
        }

        /// <summary>
        /// A range is returned only for the part of the table matching this enum instance and containing the given cell.
        /// Null is returned for all other cases, such as:
        /// <list type="bullet">
        /// <item><description>Cell on a different sheet than the table</description></item>
        /// <item><description>Cell outside the table</description></item>
        /// <item><description>this Enum part is not included in the table (i.e. no header/totals row)</description></item>
        /// <item><description>this Enum is for a table part not yet implemented in POI, such as pivot table elements</description></item>
        /// </list>
        /// The returned range can be used to determine how style options may or may not apply to this cell.
        /// For example, <see cref="wholeTable"/> borders only apply to the outer boundary of a table, while the
        /// rest of the styling, such as font and color, could apply to all the interior cells as well.
        /// </summary>
        /// <param name="table">table to evaluate</param>
        /// <param name="cell">to evaluate</param>
        /// <returns>range in the table representing this class of cells, if it contains the given cell, or null if not applicable.
        /// Stripe style types return only the stripe range containing the given cell, or null.
        /// </returns>
        public static CellRangeAddressBase AppliesTo(this TableStyleType styleType, ITable table, ICell cell)
        {
            if (table == null || cell == null) return null;
            if ( ! cell.Sheet.SheetName.Equals(table.SheetName)) return null;
            if ( ! table.Contains(cell)) return null;
        
            CellRangeAddressBase range = styleType.GetRange(table, cell);
            if (range != null && range.IsInRange(cell.RowIndex, cell.ColumnIndex)) return range;
            // else
            return null;
        }
    }
}
