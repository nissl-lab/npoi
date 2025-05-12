/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.SS.UserModel
{
    using System;
    using System.Text.RegularExpressions;


    public static class Table
    {
        /**
         * Regular expression matching a Structured Reference (Table syntax) for XSSF table expressions.
         * Public for unit tests
         * @see <a href="https://support.office.com/en-us/article/Using-structured-references-with-Excel-tables-F5ED2452-2337-4F71-BED3-C8AE6D2B276E">
         *         Excel Structured Reference Syntax
         *      </a>
         */
        public static Regex IsStructuredReference = new Regex("[a-zA-Z_\\\\][a-zA-Z0-9._]*\\[.*\\]");
    }
    /**
     * XSSF Only!
     * High level abstraction of table in a workbook.
     */
    public interface ITable
    {
        /// <summary>
        ///  Get the top-left column index relative to the sheet
        /// </summary>
        int StartColIndex { get; }
        /// <summary>
        /// Get the top-left row index on the sheet
        /// </summary>
        int StartRowIndex { get; }
        /// <summary>
        /// Get the bottom-right column index on the sheet
        /// </summary>
        int EndColIndex { get; }
        /// <summary>
        /// Get the bottom-right row index
        /// </summary>
        int EndRowIndex { get; }
        /// <summary>
        /// Get the name of the table.
        /// </summary>
        String Name { get; }
        /// <summary>
        /// Get the name of the table style, if there is one. 
        /// May be a built-in name or user-defined.
        /// </summary>
        string StyleName { get; }
        /// <summary>
        /// Returns the index of a given named column in the table (names are case insensitive in XSSF).
        /// Note this list is lazily loaded and cached for performance.
        /// Changes to the underlying table structure are not reflected in later calls
        /// unless <c>XSSFTable.updateHeaders()</c> is called to reset the cache.
        /// </summary>
        /// <param name="columnHeader">the column header name to Get the table column index of</param>
        /// <returns>column index corresponding to <c>columnHeader</c></returns>
        int FindColumnIndex(String columnHeader);
        /// <summary>
        /// Returns the sheet name that the table belongs to.
        /// </summary>
        String SheetName { get; }
        /// <summary>
        /// Note: This is misleading.  The OOXML spec indicates this is true if the totals row
        /// has <b><i>ever</i></b> been shown, not whether or not it is currently displayed.
        /// Use <see cref="getTotalsRowCount()" /> > 0 to decide whether or not the totals row is visible.
        /// </summary>
        /// <returns>true if a totals row has ever been shown for this table</returns>
        /// @see #getTotalsRowCount()
        /// <remarks>
        /// @since 3.15 beta 2
        /// </remarks>
        bool IsHasTotalsRow { get; }

        /// <summary>
        /// </summary>
        /// <returns>0 for no totals rows, 1 for totals row shown.
        /// Values > 1 are not currently used by Excel up through 2016, and the OOXML spec
        /// doesn't define how they would be implemented.
        /// </returns>
        /// <remarks>
        /// @since 3.17 beta 1
        /// </remarks>
        int TotalsRowCount { get; }
    
        /// <summary>
        /// </summary>
        /// <returns>0 for no header rows, 1 for table headers shown.
        /// Values > 1 might be used by Excel for pivot tables?
        /// </returns>
        /// <remarks>
        /// @since 3.17 beta 1
        /// </remarks>
        int HeaderRowCount { get; }
        /// <summary>
        /// TableStyleInfo for this instance
        /// </summary>
        ITableStyleInfo Style { get; }
        /// <summary>
        /// checks if the given cell is part of the table.  Includes checking that they are on the same sheet.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>true if the table and cell are on the same sheet and the cell is within the table range.</returns>
        /// <remarks>
        /// @since 3.17 beta 1
        /// </remarks>
        bool Contains(ICell cell);
    }

}