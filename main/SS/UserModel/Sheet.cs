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

    using NPOI.SS.Util;

    public enum MarginType : short
    {
        LeftMargin = 0,
        RightMargin = 1,
        TopMargin = 2,
        BottomMargin = 3
    }
    public enum PanePosition : byte
    {
        LOWER_RIGHT = 0,
        UPPER_RIGHT = 1,
        LOWER_LEFT = 2,
        UPPER_LEFT = 3,
    }

    /// <summary>
    /// High level representation of a Excel worksheet.
    /// </summary>
    /// <remarks>
    /// Sheets are the central structures within a workbook, and are where a user does most of his spreadsheet work.
    /// The most common type of sheet is the worksheet, which is represented as a grid of cells. Worksheet cells can
    /// contain text, numbers, dates, and formulas. Cells can also be formatted.
    /// </remarks>
    public interface ISheet //: IDisposable
    {

        /// <summary>
        /// Create a new row within the sheet and return the high level representation
        /// </summary>
        /// <param name="rownum">The row number.</param>
        /// <returns>high level Row object representing a row in the sheet</returns>
        /// <see>RemoveRow(Row)</see>
        IRow CreateRow(int rownum);

        /// <summary>
        /// Remove a row from this sheet.  All cells Contained in the row are Removed as well
        /// </summary>
        /// <param name="row">a row to Remove.</param>
        void RemoveRow(IRow row);

        /// <summary>
        /// Returns the logical row (not physical) 0-based.  If you ask for a row that is not
        /// defined you get a null.  This is to say row 4 represents the fifth row on a sheet.
        /// </summary>
        /// <param name="rownum">row to get (0-based).</param>
        /// <returns>the rownumber or null if its not defined on the sheet</returns>
        IRow GetRow(int rownum);

        /// <summary>
        /// Returns the number of physically defined rows (NOT the number of rows in the sheet)
        /// </summary>
        /// <value>the number of physically defined rows in this sheet.</value>
        int PhysicalNumberOfRows { get; }

        /// <summary>
        /// Gets the first row on the sheet
        /// </summary>
        /// <value>the number of the first logical row on the sheet (0-based).</value>
        int FirstRowNum { get; }

        /// <summary>
        /// Gets the last row on the sheet
        /// </summary>
        /// <value>last row contained n this sheet (0-based)</value>
        int LastRowNum { get; }

        /// <summary>
        /// whether force formula recalculation.
        /// </summary>
        bool ForceFormulaRecalculation { get; set; }

        /**
         * Get the visibility state for a given column
         *
         * @param columnIndex - the column to get (0-based)
         * @param hidden - the visiblity state of the column
         */
        void SetColumnHidden(int columnIndex, bool hidden);

        /**
         * Get the hidden state for a given column
         *
         * @param columnIndex - the column to set (0-based)
         * @return hidden - <code>false</code> if the column is visible
         */
        bool IsColumnHidden(int columnIndex);

        /**
         * Set the width (in units of 1/256th of a character width)
         * <p>
         * The maximum column width for an individual cell is 255 characters.
         * This value represents the number of characters that can be displayed
         * in a cell that is formatted with the standard font.
         * </p>
         *
         * @param columnIndex - the column to set (0-based)
         * @param width - the width in units of 1/256th of a character width
         */
        void SetColumnWidth(int columnIndex, int width);

        /**
         * get the width (in units of 1/256th of a character width )
         * @param columnIndex - the column to set (0-based)
         * @return width - the width in units of 1/256th of a character width
         */
        int GetColumnWidth(int columnIndex);

        /// <summary>
        /// Get the default column width for the sheet (if the columns do not define their own width)
        /// in characters
        /// </summary>
        /// <value>default column width measured in characters.</value>
        int DefaultColumnWidth { get; set; }

        /// <summary>
        /// Get the default row height for the sheet (if the rows do not define their own height) in
        /// twips (1/20 of  a point)
        /// </summary>
        /// <value>default row height measured in twips (1/20 of  a point)</value>
        int DefaultRowHeight { get; set; }

        /// <summary>
        /// Get the default row height for the sheet (if the rows do not define their own height) in
        /// points.
        /// </summary>
        /// <value>The default row height in points.</value>
        float DefaultRowHeightInPoints { get; set; }

        /// <summary>
        /// Returns the CellStyle that applies to the given
        /// (0 based) column, or null if no style has been
        /// set for that column
        /// </summary>
        /// <param name="column">The column.</param>
        /// <returns></returns>
        ICellStyle GetColumnStyle(int column);

        /// <summary>
        /// Adds a merged region of cells (hence those cells form one)
        /// </summary>
        /// <param name="region">(rowfrom/colfrom-rowto/colto) to merge.</param>
        /// <returns>index of this region</returns>
        int AddMergedRegion(NPOI.SS.Util.CellRangeAddress region);

        /// <summary>
        /// Determine whether printed output for this sheet will be horizontally centered.
        /// </summary>
        bool HorizontallyCenter { get; set; }

        /// <summary>
        /// Determine whether printed output for this sheet will be vertically centered.
        /// </summary>
        bool VerticallyCenter { get; set; }

        /// <summary>
        /// Removes a merged region of cells (hence letting them free)
        /// </summary>
        /// <param name="index">index of the region to unmerge</param>
        void RemoveMergedRegion(int index);

        /// <summary>
        /// Returns the number of merged regions
        /// </summary>
        int NumMergedRegions { get; }

        /// <summary>
        /// Returns the merged region at the specified index
        /// </summary>
        /// <param name="index">The index.</param>      
        NPOI.SS.Util.CellRangeAddress GetMergedRegion(int index);

        /// <summary>
        /// Returns an iterator of the physical rows
        /// </summary>
        System.Collections.IEnumerator GetRowEnumerator();

        /// <summary>
        /// Gets the flag indicating whether the window should show 0 (zero) in cells Containing zero value.
        /// When false, cells with zero value appear blank instead of showing the number zero.
        /// </summary>
        /// <value>whether all zero values on the worksheet are displayed.</value>
        bool DisplayZeros { get; set; }


        /// <summary>
        /// Gets or sets a value indicating whether the sheet displays Automatic Page Breaks.
        /// </summary>
        bool Autobreaks
        {
            get;
            set;
        }


        /// <summary>
        /// Get whether to display the guts or not,
        /// </summary>
        /// <value>default value is true</value>
        bool DisplayGuts { get; set; }

        /// <summary>
        /// Flag indicating whether the Fit to Page print option is enabled.
        /// </summary>
        bool FitToPage { get; set; }

        /**
         * Flag indicating whether summary rows appear below detail in an outline, when applying an outline.
         *
         * <p>
         * When true a summary row is inserted below the detailed data being summarized and a
         * new outline level is established on that row.
         * </p>
         * <p>
         * When false a summary row is inserted above the detailed data being summarized and a new outline level
         * is established on that row.
         * </p>
         * @return <code>true</code> if row summaries appear below detail in the outline
         */
        bool RowSumsBelow { get; set; }

        /**
         * Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline.
         *
         * <p>
         * When true a summary column is inserted to the right of the detailed data being summarized
         * and a new outline level is established on that column.
         * </p>
         * <p>
         * When false a summary column is inserted to the left of the detailed data being
         * summarized and a new outline level is established on that column.
         * </p>
         * @return <code>true</code> if col summaries appear right of the detail in the outline
         */
        bool RowSumsRight { get; set; }

        /**
         * Gets the flag indicating whether this sheet displays the lines
         * between rows and columns to make editing and Reading easier.
         *
         * @return <code>true</code> if this sheet displays gridlines.
         * @see #isPrintGridlines() to check if printing of gridlines is turned on or off
         */
        bool IsPrintGridlines { get; set; }

        /**
         * Gets the print Setup object.
         *
         * @return The user model for the print Setup object.
         */
        IPrintSetup PrintSetup { get; }

        /**
         * Gets the user model for the default document header.
         * <p/>
         * Note that XSSF offers more kinds of document headers than HSSF does
         * </p>
         * @return the document header. Never <code>null</code>
         */
        IHeader Header { get; }

        /**
         * Gets the user model for the default document footer.
         * <p/>
         * Note that XSSF offers more kinds of document footers than HSSF does.
         *
         * @return the document footer. Never <code>null</code>
         */
        IFooter Footer { get; }
        /**
         * Gets the size of the margin in inches.
         *
         * @param margin which margin to get
         * @return the size of the margin
         */
        double GetMargin(MarginType margin);

        /**
         * Sets the size of the margin in inches.
         *
         * @param margin which margin to get
         * @param size the size of the margin
         */
        void SetMargin(MarginType margin, double size);

        /**
         * Answer whether protection is enabled or disabled
         *
         * @return true => protection enabled; false => protection disabled
         */
        bool Protect { get; }

        /**
         * Answer whether scenario protection is enabled or disabled
         *
         * @return true => protection enabled; false => protection disabled
         */
        bool ScenarioProtect { get; }
        short TabColorIndex { get; set; }
        IDrawing DrawingPatriarch { get; }
        /**
         * Sets the zoom magnication for the sheet.  The zoom is expressed as a
         * fraction.  For example to express a zoom of 75% use 3 for the numerator
         * and 4 for the denominator.
         *
         * @param numerator     The numerator for the zoom magnification.
         * @param denominator   The denominator for the zoom magnification.
         */
        void SetZoom(int numerator, int denominator);

        /// <summary>
        /// The top row in the visible view when the sheet is
        /// first viewed after opening it in a viewer
        /// </summary>
        /// <value>the rownum (0 based) of the top row.</value>
        short TopRow { get; set; }

        /// <summary>
        /// The left col in the visible view when the sheet is
        /// first viewed after opening it in a viewer
        /// </summary>
        /// <value>the rownum (0 based) of the top row</value>
        short LeftCol { get; set; }

        /**
         * Sets desktop window pane display area, when the
         * file is first opened in a viewer.
         *
         * @param toprow the top row to show in desktop window pane
         * @param leftcol the left column to show in desktop window pane
         */
        void ShowInPane(short toprow, short leftcol);

        /**
         * Shifts rows between startRow and endRow n number of rows.
         * If you use a negative number, it will shift rows up.
         * Code ensures that rows don't wrap around.
         *
         * Calls shiftRows(startRow, endRow, n, false, false);
         *
         * <p>
         * Additionally shifts merged regions that are completely defined in these
         * rows (ie. merged 2 cells on a row to be shifted).
         * @param startRow the row to start shifting
         * @param endRow the row to end shifting
         * @param n the number of rows to shift
         */
        void ShiftRows(int startRow, int endRow, int n);

        /**
         * Shifts rows between startRow and endRow n number of rows.
         * If you use a negative number, it will shift rows up.
         * Code ensures that rows don't wrap around
         *
         * <p>
         * Additionally shifts merged regions that are completely defined in these
         * rows (ie. merged 2 cells on a row to be shifted).
         * <p>
         * @param startRow the row to start shifting
         * @param endRow the row to end shifting
         * @param n the number of rows to shift
         * @param copyRowHeight whether to copy the row height during the shift
         * @param reSetOriginalRowHeight whether to set the original row's height to the default
         */
        void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool reSetOriginalRowHeight);

        /**
         * Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
         * @param colSplit      Horizonatal position of split.
         * @param rowSplit      Vertical position of split.
         * @param topRow        Top row visible in bottom pane
         * @param leftmostColumn   Left column visible in right pane.
         */
        void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow);

        /**
         * Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
         * @param colSplit      Horizonatal position of split.
         * @param rowSplit      Vertical position of split.
         */
        void CreateFreezePane(int colSplit, int rowSplit);

        /**
         * Creates a split pane. Any existing freezepane or split pane is overwritten.
         * @param xSplitPos      Horizonatal position of split (in 1/20th of a point).
         * @param ySplitPos      Vertical position of split (in 1/20th of a point).
         * @param topRow        Top row visible in bottom pane
         * @param leftmostColumn   Left column visible in right pane.
         * @param activePane    Active pane.  One of: PANE_LOWER_RIGHT,
         *                      PANE_UPPER_RIGHT, PANE_LOWER_LEFT, PANE_UPPER_LEFT
         * @see #PANE_LOWER_LEFT
         * @see #PANE_LOWER_RIGHT
         * @see #PANE_UPPER_LEFT
         * @see #PANE_UPPER_RIGHT
         */
        void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane);

        /// <summary>
        /// Returns the information regarding the currently configured pane (split or freeze)
        /// </summary>
        /// <value>if no pane configured returns <c>null</c> else return the pane information.</value>
        PaneInformation PaneInformation { get; }

        /// <summary>
        /// Returns if gridlines are displayed
        /// </summary>
        bool DisplayGridlines { get; set; }

        /// <summary>
        /// Returns if formulas are displayed
        /// </summary>
        bool DisplayFormulas { get; set; }

        /// <summary>
        /// Returns if RowColHeadings are displayed.
        /// </summary>
        bool DisplayRowColHeadings { get; set; }
        /// <summary>
        /// Returns if RowColHeadings are displayed.
        /// </summary>
        bool IsActive { get; set; }
        /// <summary>
        /// Determines if there is a page break at the indicated row
        /// </summary>
        /// <param name="row">The row.</param>
        bool IsRowBroken(int row);

        /// <summary>
        /// Removes the page break at the indicated row
        /// </summary>
        /// <param name="row">The row index.</param>
        void RemoveRowBreak(int row);

        /// <summary>
        /// Retrieves all the horizontal page breaks
        /// </summary>
        /// <value>all the horizontal page breaks, or null if there are no row page breaks</value>
        int[] RowBreaks { get; }

        /// <summary>
        /// Retrieves all the vertical page breaks
        /// </summary>
        /// <value>all the vertical page breaks, or null if there are no column page breaks.</value>
        int[] ColumnBreaks { get; }

        void SetActiveCell(int row, int column);
        void SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn);
        void SetActiveCellRange(List<CellRangeAddress8Bit> cellranges, int activeRange, int activeRow, int activeColumn);
        /// <summary>
        /// Sets a page break at the indicated column
        /// </summary>
        /// <param name="column">The column.</param>
        void SetColumnBreak(int column);

        /// <summary>
        /// Sets the row break.
        /// </summary>
        /// <param name="row">The row.</param>
        void SetRowBreak(int row);

        /// <summary>
        /// Determines if there is a page break at the indicated column
        /// </summary>
        /// <param name="column">The column index.</param>
        bool IsColumnBroken(int column);

        /// <summary>
        /// Removes a page break at the indicated column
        /// </summary>
        /// <param name="column">The column.</param>
        void RemoveColumnBreak(int column);

        /// <summary>
        /// Expands or collapses a column group.
        /// </summary>
        /// <param name="columnNumber">One of the columns in the group.</param>
        /// <param name="collapsed">if set to <c>true</c>collapse group.<c>false</c>expand group.</param>
        void SetColumnGroupCollapsed(int columnNumber, bool collapsed);

        /// <summary>
        /// Create an outline for the provided column range.
        /// </summary>
        /// <param name="fromColumn">beginning of the column range.</param>
        /// <param name="toColumn">end of the column range.</param>
        void GroupColumn(int fromColumn, int toColumn);

        /// <summary>
        /// Ungroup a range of columns that were previously groupped
        /// </summary>
        /// <param name="fromColumn">start column (0-based).</param>
        /// <param name="toColumn">end column (0-based).</param>
        void UngroupColumn(int fromColumn, int toColumn);

        /// <summary>
        /// Tie a range of rows toGether so that they can be collapsed or expanded
        /// </summary>
        /// <param name="fromRow">start row (0-based)</param>
        /// <param name="toRow">end row (0-based)</param>
        void GroupRow(int fromRow, int toRow);

        /// <summary>
        /// Ungroup a range of rows that were previously groupped
        /// </summary>
        /// <param name="fromRow">start row (0-based)</param>
        /// <param name="toRow">end row (0-based)</param>
        void UngroupRow(int fromRow, int toRow);

        /// <summary>
        /// Set view state of a groupped range of rows
        /// </summary>
        /// <param name="row">start row of a groupped range of rows (0-based).</param>
        /// <param name="collapse">whether to expand/collapse the detail rows.</param>
        void SetRowGroupCollapsed(int row, bool collapse);

        /// <summary>
        /// Sets the default column style for a given column.  POI will only apply this style to new cells Added to the sheet.
        /// </summary>
        /// <param name="column">the column index</param>
        /// <param name="style">the style to set</param>
        void SetDefaultColumnStyle(int column, ICellStyle style);

        /// <summary>
        /// Adjusts the column width to fit the contents.
        /// </summary>
        /// <param name="column">the column index</param>
        /// <remarks>
        /// This process can be relatively slow on large sheets, so this should
        /// normally only be called once per column, at the end of your
        /// processing.
        /// </remarks>
        void AutoSizeColumn(int column);

        /// <summary>
        /// Adjusts the column width to fit the contents.
        /// </summary>
        /// <param name="column">the column index.</param>
        /// <param name="useMergedCells">whether to use the contents of merged cells when calculating the width of the column. Default is to ignore merged cells.</param>
        /// <remarks>
        /// This process can be relatively slow on large sheets, so this should
        /// normally only be called once per column, at the end of your
        /// processing.
        /// </remarks>
        void AutoSizeColumn(int column, bool useMergedCells);

        /// <summary>
        /// Returns cell comment for the specified row and column
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        IComment GetCellComment(int row, int column);

        /// <summary>
        /// Creates the top-level drawing patriarch.
        /// </summary>
        /// <returns></returns>
        IDrawing CreateDrawingPatriarch();


        /// <summary>
        /// Gets the parent workbook.
        /// </summary>
        IWorkbook Workbook { get; }

        /// <summary>
        /// Gets the name of the sheet.
        /// </summary>
        String SheetName { get; }

        /// <summary>
        /// Gets or sets a value indicating whether this sheet is currently selected.
        /// </summary>
        bool IsSelected { get; set; }

        void SetActive(bool sel);

                /**
         * Sets array formula to specified region for result.
         *
         * @param formula text representation of the formula
         * @param range Region of array formula for result.
         * @return the {@link CellRange} of cells affected by this change
         */
        ICellRange<ICell> SetArrayFormula(String formula, CellRangeAddress range);

        /**
         * Remove a Array Formula from this sheet.  All cells contained in the Array Formula range are removed as well
         *
         * @param cell   any cell within Array Formula range
         * @return the {@link CellRange} of cells affected by this change
         */
        ICellRange<ICell> RemoveArrayFormula(ICell cell);

        bool IsMergedRegion(CellRangeAddress mergedRegion);
    }

}