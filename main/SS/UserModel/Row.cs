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
    using System.Collections.Generic;
    using System.Collections;
    /**
     * Used to specify the different possible policies
     *  if for the case of null and blank cells
     */
    public class MissingCellPolicy
    {
        private static int NEXT_ID = 1;
        public int id;
        public MissingCellPolicy()
        {
            this.id = NEXT_ID++;
        }
        /** Missing cells are returned as null, Blank cells are returned as normal */
        public static MissingCellPolicy RETURN_NULL_AND_BLANK = new MissingCellPolicy();
        /** Missing cells are returned as null, as are blank cells */
        public static MissingCellPolicy RETURN_BLANK_AS_NULL = new MissingCellPolicy();
        /** A new, blank cell is Created for missing cells. Blank cells are returned as normal */
        public static MissingCellPolicy CREATE_NULL_AS_BLANK = new MissingCellPolicy();
    }
    /**
     * High level representation of a row of a spreadsheet.
     */
    public interface IRow
    {

        /**
         * Use this to create new cells within the row and return it.
         * <p>
         * The cell that is returned is a {@link Cell#CELL_TYPE_BLANK}. The type can be Changed
         * either through calling <code>SetCellValue</code> or <code>SetCellType</code>.
         *
         * @param column - the column number this cell represents
         * @return Cell a high level representation of the Created cell.
         * @throws ArgumentException if columnIndex < 0 or greater than the maximum number of supported columns
         * (255 for *.xls, 1048576 for *.xlsx)
         */
        ICell CreateCell(int column);

        /**
         * Use this to create new cells within the row and return it.
         * <p>
         * The cell that is returned is a {@link Cell#CELL_TYPE_BLANK}. The type can be Changed
         * either through calling SetCellValue or SetCellType.
         *
         * @param column - the column number this cell represents
         * @return Cell a high level representation of the Created cell.
         * @throws ArgumentException if columnIndex < 0 or greate than a maximum number of supported columns
         * (255 for *.xls, 1048576 for *.xlsx)
         */
        ICell CreateCell(int column, NPOI.SS.UserModel.CellType type);

        /**
         * Remove the Cell from this row.
         *
         * @param cell the cell to remove
         */
        void RemoveCell(ICell cell);

        /**
         * Get row number this row represents
         *
         * @return the row number (0 based)
         */
        int RowNum { get; set; }

        /**
         * Get the cell representing a given column (logical cell) 0-based.  If you
         * ask for a cell that is not defined....you get a null.
         *
         * @param cellnum  0 based column number
         * @return Cell representing that column or null if undefined.
         * @see #GetCell(int, NPOI.SS.usermodel.Row.MissingCellPolicy)
         */
        ICell GetCell(int cellnum);

        /**
         * Returns the cell at the given (0 based) index, with the specified {@link NPOI.SS.usermodel.Row.MissingCellPolicy}
         *
         * @return the cell at the given (0 based) index
         * @throws ArgumentException if cellnum < 0 or the specified MissingCellPolicy is invalid
         * @see Row#RETURN_NULL_AND_BLANK
         * @see Row#RETURN_BLANK_AS_NULL
         * @see Row#CREATE_NULL_AS_BLANK
         */
        ICell GetCell(int cellnum, MissingCellPolicy policy);

        /**
         * Get the number of the first cell Contained in this row.
         *
         * @return short representing the first logical cell in the row,
         *  or -1 if the row does not contain any cells.
         */
        int FirstCellNum { get; }

        /**
         * Gets the index of the last cell Contained in this row <b>PLUS ONE</b>. The result also
         * happens to be the 1-based column number of the last cell.  This value can be used as a
         * standard upper bound when iterating over cells:
         * <pre>
         * short minColIx = row.GetFirstCellNum();
         * short maxColIx = row.GetLastCellNum();
         * for(short colIx=minColIx; colIx&lt;maxColIx; colIx++) {
         *   Cell cell = row.GetCell(colIx);
         *   if(cell == null) {
         *     continue;
         *   }
         *   //... do something with cell
         * }
         * </pre>
         *
         * @return short representing the last logical cell in the row <b>PLUS ONE</b>,
         *   or -1 if the row does not contain any cells.
         */
        int LastCellNum { get; }

        /**
         * Gets the number of defined cells (NOT number of cells in the actual row!).
         * That is to say if only columns 0,4,5 have values then there would be 3.
         *
         * @return int representing the number of defined cells in the row.
         */
        int PhysicalNumberOfCells { get; }

        /**
         * Get whether or not to display this row with 0 height
         *
         * @return - zHeight height is zero or not.
         */
        bool ZeroHeight { get; set; }

        /**
         * Get the row's height measured in twips (1/20th of a point). If the height is not Set, the default worksheet value is returned,
         * See {@link Sheet#GetDefaultRowHeightInPoints()}
         *
         * @return row height measured in twips (1/20th of a point)
         */
        short Height { get; set; }

        /**
         * Returns row height measured in point size. If the height is not Set, the default worksheet value is returned,
         * See {@link Sheet#GetDefaultRowHeightInPoints()}
         *
         * @return row height measured in point size
         * @see Sheet#GetDefaultRowHeightInPoints()
         */
        float HeightInPoints { get; set; }

        /**
         * @return Cell iterator of the physically defined cells.  Note element 4 may
         * actually be row cell depending on how many are defined!
         */
        IEnumerator GetEnumerator();

        /**
         * Returns the Sheet this row belongs to
         *
         * @return the Sheet that owns this row
         */
        ISheet Sheet { get; }

        ICellStyle RowStyle { get; set; }

        void MoveCell(ICell cell, int idx);

        List<ICell> Cells { get; }
    }
}

