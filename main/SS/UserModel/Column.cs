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

using System.Collections.Generic;

namespace NPOI.SS.UserModel
{
    /// <summary>
    /// High level representation of a column of a spreadsheet.
    /// </summary>  
    public interface IColumn : IEnumerable<ICell>
    {
        /// <summary>
        /// Use this to create new cells within the column and return it.
        /// The cell that is returned is a <see cref="CellType.Blank"/>. The type can be Changed
        /// either through calling <see cref="ICell.SetCellValue"/> or <see cref="ICell.SetCellType"/>.
        /// </summary>
        /// <param name="rowIndex">the row number this cell represents</param>
        /// <returns>a high level representation of the Created cell</returns>
        /// <exception cref="ArgumentOutOfRangeException">if columnIndex is less than 0 or greater than 16384, 
        /// the maximum number of columns supported by the SpreadsheetML format(.xlsx)</exception>
        ICell CreateCell(int rowIndex);

        /// <summary>
        /// Use this to create new cells within the column and return it.
        /// </summary>
        /// <param name="rowIndex">the row number this cell represents</param>
        /// <param name="type">the cell's data type</param>
        /// <returns>a high level representation of the Created cell.</returns>
        /// <exception cref="ArgumentOutOfRangeException">if columnIndex is less than 0 or greater than 16384, 
        /// the maximum number of columns supported by the SpreadsheetML format(.xlsx)</exception>
        ICell CreateCell(int rowIndex, CellType type);

        /// <summary>
        /// Remove the Cell from this column.
        /// </summary>
        /// <param name="cell">the cell to remove</param>
        /// <exception cref="ArgumentException"></exception>
        void RemoveCell(ICell cell);

        /// <summary>
        /// Get column number this column represents
        /// </summary>
        /// <returns>the column number (0 based)</returns>
        int ColumnNum { get; set; }

        /// <summary>
        /// Returns the cell at the given (0 based) index,
        /// with the <see cref="MissingCellPolicy"/> from the parent Workbook.
        /// </summary>
        /// <param name="cellNum"></param>
        /// <returns>the cell at the given (0 based) index</returns>
        ICell GetCell(int cellNum);

        /// <summary>
        /// Returns the cell at the given (0 based) index, with the specified <see cref="MissingCellPolicy"/>
        /// </summary>
        /// <param name="cellNum"></param>
        /// <param name="policy"></param>
        /// <returns>the cell at the given (0 based) index</returns>
        /// <exception cref="ArgumentException">if cellnum is less than 0 or the specified MissingCellPolicy is invalid</exception>
        ICell GetCell(int cellNum, MissingCellPolicy policy);

        /// <summary>
        /// Get the number of the first cell Contained in this column.
        /// </summary>
        /// <returns>short representing the first logical cell in the column,
        /// or -1 if the column does not contain any cells.</returns>
        short FirstCellNum { get; }

        /// <summary>
        /// Gets the index of the last cell Contained in this column <b>PLUS ONE</b>. The result also
        /// happens to be the 1-based column number of the last cell. This value can be used as a
        /// standard upper bound when iterating over cells:
        /// </summary>
        /// <returns>short representing the last logical cell in the column <b>PLUS ONE</b>,
        /// or -1 if the column does not contain any cells.</returns>
        short LastCellNum { get; }

        /// <summary>
        /// Gets the number of defined cells (NOT number of cells in the actual column!).
        /// That is to say if only rows 0,4,5 have values then there would be 3.
        /// </summary>
        /// <returns>int representing the number of defined cells in the column.</returns>
        int PhysicalNumberOfCells { get; }

        /// <summary>
        /// Get whether or not to display this column with 0 width
        /// </summary>
        bool ZeroWidth { get; set; }

        /// <summary>
        /// Get the column's width. 
        /// If the height is not Set, the default worksheet value is returned
        /// </summary>
        /// <returns>column height</returns>
        double Width { get; set; }

        /// <summary>
        /// Is this column formatted? Most aren't, but some columns
        /// do have whole-column styles. For those that do, you
        /// can get the formatting from <see cref="ColumnStyle"/>
        /// </summary>
        bool IsFormatted { get; }

        /// <summary>
        /// Is the column width set to best fit the content?
        /// </summary>
        bool IsBestFit { get; set; }

        /// <summary>
        /// Returns the Sheet this column belongs to
        /// </summary>
        /// <returns>the Sheet that owns this column</returns>
        ISheet Sheet { get; }

        /// <summary>
        /// Returns the whole-column cell style. Most columns won't
        /// have one of these, so will return null. Call
        /// <see cref="IsFormatted"/> to check first.
        /// </summary>
        ICellStyle ColumnStyle { get; set; }

        /// <summary>
        /// Remove the Cell from this column.
        /// </summary>
        /// <param name="cell">the cell to remove</param>
        /// <exception cref="ArgumentException"></exception>
        void MoveCell(ICell cell, int newRow);

        /// <summary>
        /// Copy the current column to the target column
        /// </summary>
        /// <param name="targetIndex">column index of the target column</param>
        /// <returns>the new copied column object</returns>
        IColumn CopyColumnTo(int targetIndex);

        /// <summary>
        /// Copy the source cell to the target cell. If the target cell exists, the new copied cell will be inserted before the existing one
        /// </summary>
        /// <param name="sourceIndex">index of the source cell</param>
        /// <param name="targetIndex">index of the target cell</param>
        /// <returns>the new copied cell object</returns>
        ICell CopyCell(int sourceIndex, int targetIndex);

        /// <summary>
        /// Get cells in the column
        /// </summary>
        List<ICell> Cells { get; }

        /// <summary>
        /// Returns the columns outline level. Increased as you
        /// put it into more groups (outlines), reduced as
        /// you take it out of them.
        /// </summary>
        int OutlineLevel { get; set; }

        bool Hidden { get; set; }

        bool Collapsed { get; set; }
    }
}
