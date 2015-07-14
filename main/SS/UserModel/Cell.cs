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
    using NPOI.SS.Util;

    public enum CellType : int
    {
        Unknown = -1,
        Numeric = 0,
        String = 1,
        Formula = 2,
        Blank = 3,
        Boolean = 4,
        Error = 5
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

        /// <summary>
        /// zero-based column index of a column in a sheet.
        /// </summary>
        int ColumnIndex
        {
            get;
        }

        /// <summary>
        /// zero-based row index of a row in the sheet that contains this cell
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// the sheet this cell belongs to
        /// </summary>
        ISheet Sheet { get; }

        /// <summary>
        /// the row this cell belongs to
        /// </summary>
        IRow Row { get; }

        /// <summary>
        /// Set the cells type (numeric, formula or string)
        /// </summary>
        /// <p>If the cell currently contains a value, the value will
        /// be converted to match the new type, if possible. Formatting
        /// is generally lost in the process however.</p>
        /// <p>If what you want to do is get a String value for your
        /// numeric cell, <i>stop!</i>. This is not the way to do it.
        /// Instead, for fetching the string value of a numeric or boolean
        /// or date cell, use {@link DataFormatter} instead.</p> 
        CellType CellType
        {
            get;
        }

        /// <summary>
        /// Set the cells type (numeric, formula or string)
        /// </summary>
        /// <param name="cellType"></param>
        void SetCellType(CellType cellType);


        /// <summary>
        /// Only valid for formula cells
        /// </summary>
        CellType CachedFormulaResultType { get; }

   
        /// <summary>
        /// Set a numeric value for the cell
        /// </summary>
        /// <param name="value">the numeric value to set this cell to.  For formulas we'll set the
        ///  precalculated value, for numerics we'll set its value. For other types we will change 
        ///  the cell to a numeric cell and set its value.
        /// </param>
        void SetCellValue(double value);

        /// <summary>
        /// Set a error value for the cell
        /// </summary>
        /// <param name="value">the error value to set this cell to.  For formulas we'll set the 
        /// precalculated value , for errors we'll set its value. For other types we will change 
        /// the cell to an error cell and set its value.
        /// </param>
        void SetCellErrorValue(byte value);

        /// <summary>
        /// Converts the supplied date to its equivalent Excel numeric value and Sets that into the cell.
        /// </summary>
        /// <param name="value">the numeric value to set this cell to.  For formulas we'll set the
        ///  precalculated value, for numerics we'll set its value. For other types we will change 
        ///  the cell to a numerics cell and set its value.
        /// </param>
        void SetCellValue(DateTime value);

        /// <summary>
        /// Set a rich string value for the cell.
        /// </summary>
        /// <param name="value">value to set the cell to.  For formulas we'll set the formula
        /// string, for String cells we'll set its value.  For other types we will
        ///  change the cell to a string cell and set its value.
        ///  If value is null then we will change the cell to a Blank cell.
        ///  </param>
        void SetCellValue(IRichTextString value);

        /// <summary>
        /// Set a string value for the cell.
        /// </summary>
        /// <param name="value">value to set the cell to.  For formulas we'll set the formula 
        /// string, for String cells we'll set its value.  For other types we will 
        /// change the cell to a string cell and set its value. 
        /// If value is null then we will change the cell to a blank cell.
        /// </param>
        void SetCellValue(String value);

        /// <summary>
        /// Copy the cell to the target index. If the target cell exists, a new cell will be inserted before the existing cell.
        /// </summary>
        /// <param name="targetIndex">target index</param>
        /// <returns>the new copied cell object</returns>
        ICell CopyCellTo(int targetIndex);

        /// <summary>
        /// Return a formula for the cell
        /// </summary>
        /// <exception cref="InvalidOperationException">if the cell type returned by GetCellType() is not CELL_TYPE_FORMULA </exception>
        String CellFormula { get; set; }

        /// <summary>
        /// Sets formula for this cell.
        /// </summary>
        /// <param name="formula">the formula to Set, e.g. <code>"SUM(C4:E4)"</code>.</param>
        void SetCellFormula(String formula);        

        /// <summary>
        /// Get the value of the cell as a number.
        /// </summary>
        /// <exception cref="InvalidOperationException">if the cell type returned by GetCellType() is CELL_TYPE_STRING</exception>
        /// <exception cref="FormatException">if the cell value isn't a parsable double</exception>
        double NumericCellValue { get; }

        /// <summary>
        /// Get the value of the cell as a date.
        /// </summary>
        /// <exception cref="InvalidOperationException">if the cell type returned by GetCellType() is CELL_TYPE_STRING</exception>
        /// <exception cref="FormatException">if the cell value isn't a parsable double</exception>
        DateTime DateCellValue { get; }

        /// <summary>
        /// Get the value of the cell RichTextString
        /// </summary>
        IRichTextString RichStringCellValue { get; }

        /// <summary>
        /// Get the value of the cell as an error code.
        /// </summary>
        byte ErrorCellValue { get; }

        /// <summary>
        /// Get the value of the cell as a string
        /// </summary>
        String StringCellValue { get; }

        /// <summary>
        /// Set a bool value for the cell
        /// </summary>
        /// <param name="value"></param>
        void SetCellValue(bool value);

        /// <summary>
        /// Get the value of the cell as a bool.
        /// </summary>
        bool BooleanCellValue { get; }

        /// <summary>
        /// Return the cell's style.
        /// </summary>
        ICellStyle CellStyle { get; set; }

        /// <summary>
        /// Sets this cell as the active cell for the worksheet
        /// </summary>
        void SetAsActiveCell();

        /// <summary>
        /// comment associated with this cell
        /// </summary>
        IComment CellComment { get; set; }

        /// <summary>
        /// Removes the comment for this cell, if there is one.
        /// </summary>
        void RemoveCellComment();

         /// <summary>
        /// hyperlink associated with this cell
        /// </summary>
        IHyperlink Hyperlink { get; set; }

        /// <summary>
        /// Removes the hyperlink for this cell, if there is one.
        /// </summary>
        void RemoveHyperlink();

        /// <summary>
        ///  Only valid for array formula cells
        /// </summary>
        /// <returns>range of the array formula group that the cell belongs to.</returns>
        CellRangeAddress ArrayFormulaRange{ get; }

        /// <summary>
        /// if this cell is part of group of cells having a common array formula.
        /// </summary>
        bool IsPartOfArrayFormulaGroup { get; }

        bool IsMergedCell { get; }
    }
}

