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
        /// <summary>
        /// Regular expression matching a Structured Reference (Table syntax) for XSSF table expressions.
        /// Public for unit tests
        /// <a href="https://support.office.com/en-us/article/Using-structured-references-with-Excel-tables-F5ED2452-2337-4F71-BED3-C8AE6D2B276E">Excel Structured Reference Syntax</a>
        /// </summary>
        public static Regex IsStructuredReference = new Regex("[a-zA-Z_\\\\][a-zA-Z0-9._]*\\[.*\\]");
    }
    /// <summary>
    /// XSSF Only!
    /// High level abstraction of table in a workbook.
    /// </summary>
    public interface ITable
    {
        /// <summary>
        ///  Get the top-left column index relative to the sheet
        /// </summary>
        /// <returns>table start column index on sheet</returns>
        int StartColIndex { get; }
        /// <summary>
        ///  Get the top-left row index on the sheet
        /// </summary>
        /// <returns>table start row index on sheet</returns>
        int StartRowIndex { get; }
        /// <summary>
        ///  Get the bottom-right column index on the sheet
        /// </summary>
        /// <returns>table end column index on sheet</returns>
        int EndColIndex { get; }
        /// <summary>
        ///  Get the bottom-right row index
        /// </summary>
        /// <returns>table end row index on sheet</returns>
        int EndRowIndex { get; }
        /// <summary>
        /// Get the name of the table.
        /// </summary>
        /// <returns>table name</returns>
        String Name { get; }

        /// <summary>
        /// Returns the index of a given named column in the table (names are case insensitive in XSSF).
        /// Note this list is lazily loaded and cached for performance.
        /// Changes to the underlying table structure are not reflected in later calls
        /// unless <c>XSSFTable.UpdateHeaders()</c> is called to reset the cache.
        /// </summary>
        /// <param name="columnHeader">the column header name to Get the table column index of</param>
        /// <returns>column index corresponding to <c>columnHeader</c></returns>
        int FindColumnIndex(String columnHeader);
        /// <summary>
        /// Returns the sheet name that the table belongs to.
        /// </summary>
        String SheetName { get; }
        /// <summary>
        /// Returns true iff the table has a 'Totals' row
        /// </summary>
        bool IsHasTotalsRow { get; }


    }

}