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
    using NPOI.SS.Formula;
    using NPOI.SS.Util;

    public enum CellType : int
    {
        Unknown = -1,
        NUMERIC = 0,
        STRING = 1,
        FORMULA = 2,
        BLANK = 3,
        BOOLEAN = 4,
        ERROR = 5
    }

    /**
     * High level representation of a cell in a row of a spreadsheet.
     * <p>
     * Cells can be numeric, formula-based or string-based (text).  The cell type
     * specifies this.  String cells cannot conatin numbers and numeric cells cannot
     * contain strings (at least according to our model).  Client apps should do the
     * conversions themselves.  Formula cells have the formula string, as well as
     * the formula result, which can be numeric or string.
     * </p>
     * <p>
     * Cells should have their number (0 based) before being Added to a row.
     * </p>
     */
    public interface ICell
    {

        /**
         * Returns column index of this cell
         *
         * @return zero-based column index of a column in a sheet.
         */
        int ColumnIndex
        {
            get;
            set;
        }

        /**
         * Returns row index of a row in the sheet that Contains this cell
         *
         * @return zero-based row index of a row in the sheet that Contains this cell
         */
        int RowIndex { get; }

        /**
         * Returns the sheet this cell belongs to
         *
         * @return the sheet this cell belongs to
         */
        ISheet Sheet { get; }

        /**
         * Returns the Row this cell belongs to
         *
         * @return the Row that owns this cell
         */
        IRow Row { get; }

        /**
         * Set the cells type (numeric, formula or string)
         *
         * @throws ArgumentException if the specified cell type is invalid
         * @see #CELL_TYPE_NUMERIC
         * @see #CELL_TYPE_STRING
         * @see #CELL_TYPE_FORMULA
         * @see #CELL_TYPE_BLANK
         * @see #CELL_TYPE_BOOLEAN
         * @see #CELL_TYPE_ERROR
         */
        CellType CellType
        {
            get;
        }
        /**
 * Set the cells type (numeric, formula or string)
 *
 * @throws IllegalArgumentException if the specified cell type is invalid
 * @see #CELL_TYPE_NUMERIC
 * @see #CELL_TYPE_STRING
 * @see #CELL_TYPE_FORMULA
 * @see #CELL_TYPE_BLANK
 * @see #CELL_TYPE_BOOLEAN
 * @see #CELL_TYPE_ERROR
 */
        /**
* Set a error value for the cell
*
* @param value the error value to set this cell to.  For formulas we'll set the
*        precalculated value , for errors we'll set
*        its value. For other types we will change the cell to an error
*        cell and set its value.
* @see FormulaError
*/

        void SetCellType(CellType cellType);
        /**
         * Only valid for formula cells
         * @return one of ({@link #CELL_TYPE_NUMERIC}, {@link #CELL_TYPE_STRING},
         *     {@link #CELL_TYPE_BOOLEAN}, {@link #CELL_TYPE_ERROR}) depending
         * on the cached value of the formula
         */



        NPOI.SS.UserModel.CellType CachedFormulaResultType { get; }

        /**
         * Set a numeric value for the cell
         *
         * @param value  the numeric value to set this cell to.  For formulas we'll set the
         *        precalculated value, for numerics we'll set its value. For other types we
         *        will change the cell to a numeric cell and set its value.
         */
        void SetCellValue(double value);
        /**
         * Set a error value for the cell
         *
         * @param value the error value to set this cell to.  For formulas we'll set the
         *        precalculated value , for errors we'll set
         *        its value. For other types we will change the cell to an error
         *        cell and set its value.
         * @see FormulaError
         */
        void SetCellErrorValue(byte value);
        /**
         * Converts the supplied date to its equivalent Excel numeric value and Sets
         * that into the cell.
         * <p/>
         * <b>Note</b> - There is actually no 'DATE' cell type in Excel. In many
         * cases (when entering date values), Excel automatically adjusts the
         * <i>cell style</i> to some date format, creating the illusion that the cell
         * data type is now something besides {@link Cell#CELL_TYPE_NUMERIC}.  POI
         * does not attempt to replicate this behaviour.  To make a numeric cell
         * display as a date, use {@link #SetCellStyle(CellStyle)} etc.
         *
         * @param value the numeric value to set this cell to.  For formulas we'll set the
         *        precalculated value, for numerics we'll set its value. For other types we
         *        will change the cell to a numerics cell and set its value.
         */
        void SetCellValue(DateTime value);

        /**
         * Set a rich string value for the cell.
         *
         * @param value  value to set the cell to.  For formulas we'll set the formula
         * string, for String cells we'll set its value.  For other types we will
         * change the cell to a string cell and set its value.
         * If value is null then we will change the cell to a Blank cell.
         */
        void SetCellValue(IRichTextString value);

        /**
         * Set a string value for the cell.
         *
         * @param value  value to set the cell to.  For formulas we'll set the formula
         * string, for String cells we'll set its value.  For other types we will
         * change the cell to a string cell and set its value.
         * If value is null then we will change the cell to a Blank cell.
         */
        void SetCellValue(String value);

        /**
         * Return a formula for the cell, for example, <code>SUM(C4:E4)</code>
         *
         * @return a formula for the cell
         * @throws InvalidOperationException if the cell type returned by {@link #GetCellType()} is not CELL_TYPE_FORMULA
         */
        String CellFormula { get; set; }

        /**
         * Get the value of the cell as a number.
         * <p>
         * For strings we throw an exception. For blank cells we return a 0.
         * For formulas or error cells we return the precalculated value;
         * </p>
         * @return the value of the cell as a number
         * @throws InvalidOperationException if the cell type returned by {@link #GetCellType()} is CELL_TYPE_STRING
         * @exception FormatException if the cell value isn't a parsable <code>double</code>.
         * @see DataFormatter for turning this number into a string similar to that which Excel would render this number as.
         */
        double NumericCellValue { get; }

        /**
         * Get the value of the cell as a date.
         * <p>
         * For strings we throw an exception. For blank cells we return a null.
         * </p>
         * @return the value of the cell as a date
         * @throws InvalidOperationException if the cell type returned by {@link #GetCellType()} is CELL_TYPE_STRING
         * @exception FormatException if the cell value isn't a parsable <code>double</code>.
         * @see DataFormatter for formatting  this date into a string similar to how excel does.
         */
        DateTime DateCellValue { get; }

        /**
         * Get the value of the cell as a XSSFRichTextString
         * <p>
         * For numeric cells we throw an exception. For blank cells we return an empty string.
         * For formula cells we return the pre-calculated value.
         * </p>
         * @return the value of the cell as a XSSFRichTextString
         */
        IRichTextString RichStringCellValue { get; }
        /**
 * Get the value of the cell as an error code.
 * <p>
 * For strings, numbers, and booleans, we throw an exception.
 * For blank cells we return a 0.
 * </p>
 *
 * @return the value of the cell as an error code
 * @throws IllegalStateException if the cell type returned by {@link #getCellType()} isn't CELL_TYPE_ERROR
 * @see FormulaError for error codes
 */
        byte ErrorCellValue { get; }

        /**
         * Get the value of the cell as a string
         * <p>
         * For numeric cells we throw an exception. For blank cells we return an empty string.
         * For formulaCells that are not string Formulas, we return empty String.
         * </p>
         * @return the value of the cell as a string
         */
        String StringCellValue { get; }

        /**
         * Set a bool value for the cell
         *
         * @param value the bool value to set this cell to.  For formulas we'll set the
         *        precalculated value, for bools we'll set its value. For other types we
         *        will change the cell to a bool cell and set its value.
         */
        void SetCellValue(bool value);

        /**
         * Set a error value for the cell
         *
         * @param value the error value to set this cell to.  For formulas we'll set the
         *        precalculated value , for errors we'll set
         *        its value. For other types we will change the cell to an error
         *        cell and set its value.
         * @see FormulaError
         */
        byte CellErrorValue { get; set; }

        /**
         * Get the value of the cell as a bool.
         * <p>
         * For strings, numbers, and errors, we throw an exception. For blank cells we return a false.
         * </p>
         * @return the value of the cell as a bool
         * @throws InvalidOperationException if the cell type returned by {@link #GetCellType()}
         *   is not CELL_TYPE_BOOLEAN, CELL_TYPE_BLANK or CELL_TYPE_FORMULA
         */
        bool BooleanCellValue { get; }

        /**
         * Return the cell's style.
         *
         * @return the cell's style. Always not-null. Default cell style has zero index and can be obtained as
         * <code>workbook.GetCellStyleAt(0)</code>
         * @see Workbook#GetCellStyleAt(short)
         */
        ICellStyle CellStyle { get; set; }

        /**
         * Sets this cell as the active cell for the worksheet
         */
        void SetAsActiveCell();

        /**
         * Returns comment associated with this cell
         *
         * @return comment associated with this cell or <code>null</code> if not found
         */
        IComment CellComment { get; set; }

        /**
         * Removes the comment for this cell, if there is one.
         */
        void RemoveCellComment();

        /**
         * Returns hyperlink associated with this cell
         *
         * @return hyperlink associated with this cell or <code>null</code> if not found
         */
        IHyperlink Hyperlink { get; set; }

        /**
         * Only valid for array formula cells
         *
         * @return range of the array formula group that the cell belongs to.
         */
        CellRangeAddress GetArrayFormulaRange();

        /**
         * @return <c>true</c> if this cell is part of group of cells having a common array formula.
         */
        bool IsPartOfArrayFormulaGroup { get; }

        bool IsMergedCell { get; }
    }
}

