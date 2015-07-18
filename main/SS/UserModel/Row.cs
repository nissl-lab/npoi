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
    using System.Collections.Generic;
    using System.Collections;

    /// <summary>
    /// Used to specify the different possible policies
    /// if for the case of null and blank cells
    /// </summary>    
    public class MissingCellPolicy
    {
        private static int NEXT_ID = 1;
        public int id;
        public MissingCellPolicy()
        {
            this.id = NEXT_ID++;
        }
        /// <summary>Missing cells are returned as null, Blank cells are returned as normal</summary>
        public static readonly MissingCellPolicy RETURN_NULL_AND_BLANK = new MissingCellPolicy();
        /// <summary>Missing cells are returned as null, as are blank cells</summary>
        public static readonly MissingCellPolicy RETURN_BLANK_AS_NULL = new MissingCellPolicy();
        /// <summary>A new, blank cell is Created for missing cells. Blank cells are returned as normal</summary>
        public static readonly MissingCellPolicy CREATE_NULL_AS_BLANK = new MissingCellPolicy();
    }

    /// <summary>
    /// High level representation of a row of a spreadsheet.
    /// </summary>    
    public interface IRow : IEnumerable<ICell>
    {
        /// <summary>
        /// Use this to create new cells within the row and return it.
        /// 
        /// The cell that is returned is a <see cref="ICell"/>/<see cref="CellType.Blank"/>.
        /// The type can be changed either through calling <c>SetCellValue</c> or <c>SetCellType</c>.
        /// </summary>
        /// <param name="column">the column number this cell represents</param>
        /// <returns>Cell a high level representation of the created cell.</returns>
        /// <throws>
        /// ArgumentException if columnIndex &lt; 0 or greater than the maximum number of supported columns
        /// (255 for *.xls, 1048576 for *.xlsx)
        /// </throws>
        ICell CreateCell(int column);

        /// <summary>
        /// Use this to create new cells within the row and return it.
        /// 
        /// The cell that is returned is a <see cref="ICell"/>/<see cref="CellType.Blank"/>. The type can be changed
        /// either through calling <code>SetCellValue</code> or <code>SetCellType</code>.
        /// </summary>
        /// <param name="column">the column number this cell represents</param>
        /// <param name="type"></param>
        /// <returns>Cell a high level representation of the created cell.</returns>
        /// <throws>ArgumentException if columnIndex &lt; 0 or greater than the maximum number of supported columns
        /// (255 for *.xls, 1048576 for *.xlsx)
        /// </throws>
        ICell CreateCell(int column, NPOI.SS.UserModel.CellType type);

        /// <summary>
        /// Remove the Cell from this row.
        /// </summary>
        /// <param name="cell">the cell to remove</param>
        void RemoveCell(ICell cell);

        /// <summary>
        /// Get row number this row represents
        /// </summary>        
        /// <returns>the row number (0 based)</returns>
        int RowNum { get; set; }

        /// <summary>
        /// Get the cell representing a given column (logical cell) 0-based.  If you
        /// ask for a cell that is not defined....you get a null.
        /// </summary>
        /// <param name="cellnum">0 based column number</param>
        /// <returns>Cell representing that column or null if undefined.</returns>
        /// <see cref="GetCell(int, NPOI.SS.UserModel.MissingCellPolicy)"/>
        ICell GetCell(int cellnum);

        /// <summary>
        /// Returns the cell at the given (0 based) index, with the specified {@link NPOI.SS.usermodel.Row.MissingCellPolicy}
        /// </summary>
        /// <returns>the cell at the given (0 based) index</returns>
        /// <throws>ArgumentException if cellnum &lt; 0 or the specified MissingCellPolicy is invalid</throws>
        /// <see cref="MissingCellPolicy.RETURN_NULL_AND_BLANK"/>
        /// <see cref="MissingCellPolicy.RETURN_BLANK_AS_NULL"/>
        /// <see cref="MissingCellPolicy.CREATE_NULL_AS_BLANK"/>
        ICell GetCell(int cellnum, MissingCellPolicy policy);

        /// <summary>
        /// Get the number of the first cell Contained in this row.
        /// </summary>
        /// <returns>
        /// short representing the first logical cell in the row,
        /// or -1 if the row does not contain any cells.
        /// </returns>
        short FirstCellNum { get; }

        /// <summary>
        /// Gets the index of the last cell Contained in this row <b>PLUS ONE</b>. The result also
        /// happens to be the 1-based column number of the last cell.  This value can be used as a
        /// standard upper bound when iterating over cells:
        /// <pre>
        /// short minColIx = row.GetFirstCellNum();
        /// short maxColIx = row.GetLastCellNum();
        /// for(short colIx=minColIx; colIx&lt;maxColIx; colIx++) {
        /// Cell cell = row.GetCell(colIx);
        /// if(cell == null) {
        /// continue;
        /// }
        /// //... do something with cell
        /// }
        /// </pre>
        /// </summary>
        /// <returns>
        /// short representing the last logical cell in the row <b>PLUS ONE</b>,
        /// or -1 if the row does not contain any cells.
        /// </returns>
        short LastCellNum { get; }

        /// <summary>
        /// Gets the number of defined cells (NOT number of cells in the actual row!).
        /// That is to say if only columns 0,4,5 have values then there would be 3.
        /// </summary>
        /// <returns>int representing the number of defined cells in the row.</returns>
        int PhysicalNumberOfCells { get; }

        /// <summary>
        /// Get whether or not to display this row with 0 height
        /// </summary>
        /// <returns>zHeight height is zero or not.</returns>
        bool ZeroHeight { get; set; }

        /// <summary>
        /// Get the row's height measured in twips (1/20th of a point). 
        /// If the height is not set, the default worksheet value is returned,
        /// <see cref="ISheet.DefaultRowHeightInPoints"/>
        /// </summary>
        /// <returns>row height measured in twips (1/20th of a point)</returns>
        short Height { get; set; }

        /// <summary>
        /// Returns row height measured in point size. 
        /// If the height is not set, the default worksheet value is returned,
        /// <see cref="ISheet.DefaultRowHeightInPoints"/>
        /// </summary>
        /// <returns>row height measured in point size
        /// <see cref="ISheet.DefaultRowHeightInPoints"/>
        /// </returns>
        float HeightInPoints { get; set; }
        /// <summary>
        /// Is this row formatted? Most aren't, but some rows
        /// do have whole-row styles. For those that do, you
        /// can get the formatting from <see cref="RowStyle"/>
        /// </summary>
        bool IsFormatted { get; }

        /// <summary>
        /// Returns the Sheet this row belongs to
        /// </summary>
        /// <returns>the Sheet that owns this row</returns>
        ISheet Sheet { get; }

        /// <summary>
        /// Returns the whole-row cell styles. Most rows won't
        /// have one of these, so will return null. Call IsFormmated to check first
        /// </summary>
        /// <value>The row style.</value>
        ICellStyle RowStyle { get; set; }

        /// <summary>
        /// Moves the supplied cell to a new column, which
        /// must not already have a cell there!
        /// </summary>
        /// <param name="cell">The cell to move</param>
        /// <param name="newColumn">The new column number (0 based)</param>
        void MoveCell(ICell cell, int newColumn);
        /// <summary>
        /// Copy the current row to the target row
        /// </summary>
        /// <param name="targetIndex">row index of the target row</param>
        /// <returns>the new copied row object</returns>
        IRow CopyRowTo(int targetIndex);
        /// <summary>
        /// Copy the source cell to the target cell. If the target cell exists, the new copied cell will be inserted before the existing one
        /// </summary>
        /// <param name="sourceIndex">index of the source cell</param>
        /// <param name="targetIndex">index of the target cell</param>
        /// <returns>the new copied cell object</returns>
        ICell CopyCell(int sourceIndex, int targetIndex);
        /// <summary>
        /// Get cells in the row
        /// </summary>
        List<ICell> Cells { get; }

        /// <summary>
        /// Returns the rows outline level. Increased as you
        /// put it into more groups (outlines), reduced as
        /// you take it out of them.
        /// </summary>
        int OutlineLevel { get; }
    }
}

