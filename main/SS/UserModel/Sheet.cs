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
    using System.Collections;

    /// <summary>
    /// Indicate the position of the margin. One of left, right, top and bottom.
    /// </summary>
    public enum MarginType : short
    {
        /// <summary>
        /// referes to the left margin
        /// </summary>
        LeftMargin = 0,
        /// <summary>
        /// referes to the right margin
        /// </summary>
        RightMargin = 1,
        /// <summary>
        /// referes to the top margin
        /// </summary>
        TopMargin = 2,
        /// <summary>
        /// referes to the bottom margin
        /// </summary>
        BottomMargin = 3,

        HeaderMargin = 4,
        FooterMargin = 5
    }

    /// <summary>
    /// Define the position of the pane. One of lower/right, upper/right, lower/left and upper/left.
    /// </summary>
    public enum PanePosition : byte
    {
        /// <summary>
        /// referes to the lower/right corner
        /// </summary>
        LowerRight = 0,
        /// <summary>
        /// referes to the upper/right corner
        /// </summary>
        UpperRight = 1,
        /// <summary>
        /// referes to the lower/left corner
        /// </summary>
        LowerLeft = 2,
        /// <summary>
        /// referes to the upper/left corner
        /// </summary>
        UpperLeft = 3,
    }

    /// <summary>
    /// High level representation of a Excel worksheet.
    /// </summary>
    /// <remarks>
    /// Sheets are the central structures within a workbook, and are where a user does most of his spreadsheet work.
    /// The most common type of sheet is the worksheet, which is represented as a grid of cells. Worksheet cells can
    /// contain text, numbers, dates, and formulas. Cells can also be formatted.
    /// </remarks>
    public interface ISheet
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

        /// <summary>
        /// Get the visibility state for a given column
        /// </summary>
        /// <param name="columnIndex">the column to get (0-based)</param>
        /// <param name="hidden">the visiblity state of the column</param>
        void SetColumnHidden(int columnIndex, bool hidden);

        /// <summary>
        /// Get the hidden state for a given column
        /// </summary>
        /// <param name="columnIndex">the column to set (0-based)</param>
        /// <returns>hidden - <c>false</c> if the column is visible</returns>
        bool IsColumnHidden(int columnIndex);

        /// <summary>
        /// Copy the source row to the target row. If the target row exists, the new copied row will be inserted before the existing one
        /// </summary>
        /// <param name="sourceIndex">source index</param>
        /// <param name="targetIndex">target index</param>
        /// <returns>the new copied row object</returns>
        IRow CopyRow(int sourceIndex, int targetIndex);
        /// <summary>
        /// Set the width (in units of 1/256th of a character width)
        /// </summary>
        /// <param name="columnIndex">the column to set (0-based)</param>
        /// <param name="width">the width in units of 1/256th of a character width</param>
        /// <remarks>
        /// The maximum column width for an individual cell is 255 characters.
        /// This value represents the number of characters that can be displayed
        /// in a cell that is formatted with the standard font.
        /// </remarks>
        void SetColumnWidth(int columnIndex, int width);

        /// <summary>
        /// get the width (in units of 1/256th of a character width )
        /// </summary>
        /// <param name="columnIndex">the column to get (0-based)</param>
        /// <returns>the width in units of 1/256th of a character width</returns>
        int GetColumnWidth(int columnIndex);

        /// <summary>
        /// get the width in pixel
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        /// <remarks>
        /// Please note, that this method works correctly only for workbooks
        /// with the default font size (Arial 10pt for .xls and Calibri 11pt for .xlsx).
        /// If the default font is changed the column width can be streched
        /// </remarks>
        float GetColumnWidthInPixels(int columnIndex);

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
        short DefaultRowHeight { get; set; }

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
        /// Gets the row enumerator.
        /// </summary>
        /// <returns>
        /// an iterator of the PHYSICAL rows.  Meaning the 3rd element may not
        /// be the third row if say for instance the second row is undefined.
        /// Call <see cref="NPOI.SS.UserModel.IRow.RowNum"/> on each row 
        /// if you care which one it is.
        /// </returns>
        IEnumerator GetRowEnumerator();


        /// <summary>
        /// Get the row enumerator
        /// </summary>
        /// <returns></returns>
        IEnumerator GetEnumerator();

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

        /// <summary>
        /// Flag indicating whether summary rows appear below detail in an outline, when applying an outline.
        ///
        /// 
        /// When true a summary row is inserted below the detailed data being summarized and a
        /// new outline level is established on that row.
        /// 
        /// 
        /// When false a summary row is inserted above the detailed data being summarized and a new outline level
        /// is established on that row.
        /// 
        /// </summary>
        /// <returns><c>true</c> if row summaries appear below detail in the outline</returns>
        bool RowSumsBelow { get; set; }

        /// <summary>
        /// Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline.
        ///
        /// 
        /// When true a summary column is inserted to the right of the detailed data being summarized
        /// and a new outline level is established on that column.
        /// 
        /// 
        /// When false a summary column is inserted to the left of the detailed data being
        /// summarized and a new outline level is established on that column.
        /// 
        /// </summary>
        /// <returns><c>true</c> if col summaries appear right of the detail in the outline</returns>
        bool RowSumsRight { get; set; }

        /// <summary>
        /// Gets the flag indicating whether this sheet displays the lines
        /// between rows and columns to make editing and reading easier.
        /// </summary>
        /// <returns><c>true</c> if this sheet displays gridlines.</returns>
        bool IsPrintGridlines { get; set; }

        /// <summary>
        /// Gets the print Setup object.
        /// </summary>
        /// <returns>The user model for the print Setup object.</returns>
        IPrintSetup PrintSetup { get; }

        /// <summary>
        /// Gets the user model for the default document header.
        /// <p/>
        /// Note that XSSF offers more kinds of document headers than HSSF does
        /// 
        /// </summary>
        /// <returns>the document header. Never <code>null</code></returns>
        IHeader Header { get; }

        /// <summary>
        /// Gets the user model for the default document footer.
        /// <p/>
        /// Note that XSSF offers more kinds of document footers than HSSF does.
        /// </summary>
        /// <returns>the document footer. Never <code>null</code></returns>
        IFooter Footer { get; }
        /// <summary>
        /// Gets the size of the margin in inches.
        /// </summary>
        /// <param name="margin">which margin to get</param>
        /// <returns>the size of the margin</returns>
        double GetMargin(MarginType margin);

        /// <summary>
        /// Sets the size of the margin in inches.
        /// </summary>
        /// <param name="margin">which margin to get</param>
        /// <param name="size">the size of the margin</param>
        void SetMargin(MarginType margin, double size);

        /// <summary>
        /// Answer whether protection is enabled or disabled
        /// </summary>
        /// <returns>true => protection enabled; false => protection disabled</returns>
        bool Protect { get; }

        /// <summary>
        /// Sets the protection enabled as well as the password
        /// </summary>
        /// <param name="password">to set for protection. Pass <code>null</code> to remove protection</param>
        void ProtectSheet(String password);

        /// <summary>
        /// Answer whether scenario protection is enabled or disabled
        /// </summary>
        /// <returns>true => protection enabled; false => protection disabled</returns>
        bool ScenarioProtect { get; }

        /// <summary>
        /// Gets or sets the tab color of the _sheet
        /// </summary>
        short TabColorIndex { get; set; }

        /// <summary>
        /// Returns the top-level drawing patriach, if there is one.
        /// This will hold any graphics or charts for the _sheet.
        /// WARNING - calling this will trigger a parsing of the
        /// associated escher records. Any that aren't supported
        /// (such as charts and complex drawing types) will almost
        /// certainly be lost or corrupted when written out. Only
        /// use this with simple drawings, otherwise call
        /// HSSFSheet#CreateDrawingPatriarch() and
        /// start from scratch!
        /// </summary>
        /// <value>The drawing patriarch.</value>
        IDrawing DrawingPatriarch { get; }

        /// <summary>
        /// Sets the zoom magnication for the sheet.  The zoom is expressed as a
        /// fraction.  For example to express a zoom of 75% use 3 for the numerator
        /// and 4 for the denominator.
        /// </summary>
        /// <param name="numerator">The numerator for the zoom magnification.</param>
        /// <param name="denominator">denominator for the zoom magnification.</param>
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

        /// <summary>
        /// Sets desktop window pane display area, when the file is first opened in a viewer.
        /// </summary>
        /// <param name="toprow">the top row to show in desktop window pane</param>
        /// <param name="leftcol">the left column to show in desktop window pane</param>
        void ShowInPane(int toprow, int leftcol);
        /// <summary>
        /// Sets desktop window pane display area, when the
        /// file is first opened in a viewer.
        /// </summary>
        /// <param name="toprow"> the top row to show in desktop window pane</param>
        /// <param name="leftcol"> the left column to show in desktop window pane</param>
        void ShowInPane(short toprow, short leftcol);

        /// <summary>
        /// Shifts rows between startRow and endRow n number of rows.
        /// If you use a negative number, it will shift rows up.
        /// Code ensures that rows don't wrap around.
        ///
        /// Calls shiftRows(startRow, endRow, n, false, false);
        ///
        /// 
        /// Additionally shifts merged regions that are completely defined in these
        /// rows (ie. merged 2 cells on a row to be shifted).
        /// </summary>
        /// <param name="startRow">the row to start shifting</param>
        /// <param name="endRow">the row to end shifting</param>
        /// <param name="n">the number of rows to shift</param>
        void ShiftRows(int startRow, int endRow, int n);

        /// <summary>
        /// Shifts rows between startRow and endRow n number of rows.
        /// If you use a negative number, it will shift rows up.
        /// Code ensures that rows don't wrap around
        ///
        /// Additionally shifts merged regions that are completely defined in these
        /// rows (ie. merged 2 cells on a row to be shifted).
        /// </summary>
        /// <param name="startRow">the row to start shifting</param>
        /// <param name="endRow">the row to end shifting</param>
        /// <param name="n">the number of rows to shift</param>
        /// <param name="copyRowHeight">whether to copy the row height during the shift</param>
        /// <param name="resetOriginalRowHeight">whether to set the original row's height to the default</param>
        void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight);

        /// <summary>
        /// Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizonatal position of split</param>
        /// <param name="rowSplit">Vertical position of split</param>
        /// <param name="leftmostColumn">Top row visible in bottom pane</param>
        /// <param name="topRow">Left column visible in right pane</param>
        void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow);

        /// <summary>
        /// Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizonatal position of split.</param>
        /// <param name="rowSplit">Vertical position of split.</param>
        void CreateFreezePane(int colSplit, int rowSplit);

        /// <summary>
        /// Creates a split pane. Any existing freezepane or split pane is overwritten.
        /// </summary>
        /// <param name="xSplitPos">Horizonatal position of split (in 1/20th of a point)</param>
        /// <param name="ySplitPos">Vertical position of split (in 1/20th of a point)</param>
        /// <param name="leftmostColumn">Left column visible in right pane</param>
        /// <param name="topRow">Top row visible in bottom pane</param>
        /// <param name="activePane">Active pane.  One of: PANE_LOWER_RIGHT, PANE_UPPER_RIGHT, PANE_LOWER_LEFT, PANE_UPPER_LEFT</param>
        /// @see #PANE_LOWER_LEFT
        /// @see #PANE_LOWER_RIGHT
        /// @see #PANE_UPPER_LEFT
        /// @see #PANE_UPPER_RIGHT
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

        /// <summary>
        /// Sets the active cell.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        void SetActiveCell(int row, int column);

        /// <summary>
        /// Sets the active cell range.
        /// </summary>
        /// <param name="firstRow">The firstrow.</param>
        /// <param name="lastRow">The lastrow.</param>
        /// <param name="firstColumn">The firstcolumn.</param>
        /// <param name="lastColumn">The lastcolumn.</param>
        void SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn);

        /// <summary>
        /// Sets the active cell range.
        /// </summary>
        /// <param name="cellranges">The cellranges.</param>
        /// <param name="activeRange">The index of the active range.</param>
        /// <param name="activeRow">The active row in the active range</param>
        /// <param name="activeColumn">The active column in the active range</param>
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
        /// <param name="useMergedCells">whether to use the contents of merged cells when 
        /// calculating the width of the column. Default is to ignore merged cells.</param>
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

        /// <summary>
        /// Sets whether sheet is selected.
        /// </summary>
        /// <param name="value">Whether to select the sheet or deselect the sheet.</param> 
        void SetActive(bool value);

        /// <summary>
        /// Sets array formula to specified region for result.
        /// </summary>
        /// <param name="formula">text representation of the formula</param>
        /// <param name="range">Region of array formula for result</param>
        /// <returns>the <see cref="ICellRange{ICell}"/> of cells affected by this change</returns>
        ICellRange<ICell> SetArrayFormula(String formula, CellRangeAddress range);

        /// <summary>
        /// Remove a Array Formula from this sheet.  All cells contained in the Array Formula range are removed as well
        /// </summary>
        /// <param name="cell">any cell within Array Formula range</param>
        /// <returns>the <see cref="ICellRange{ICell}"/> of cells affected by this change</returns>
        ICellRange<ICell> RemoveArrayFormula(ICell cell);

        /// <summary>
        /// Checks if the provided region is part of the merged regions.
        /// </summary>
        /// <param name="mergedRegion">Region searched in the merged regions</param>
        /// <returns><c>true</c>, when the region is contained in at least one of the merged regions</returns>
        bool IsMergedRegion(CellRangeAddress mergedRegion);

        /// <summary>
        /// Create an instance of a DataValidationHelper.
        /// </summary>
        /// <returns>Instance of a DataValidationHelper</returns>
        IDataValidationHelper GetDataValidationHelper();

        /// <summary>
        /// Returns the list of DataValidation in the sheet.
        /// </summary>
        /// <returns>list of DataValidation in the sheet</returns>
        List<IDataValidation> GetDataValidations();

        /// <summary>
        /// Creates a data validation object
        /// </summary>
        /// <param name="dataValidation">The data validation object settings</param>
        void AddValidationData(IDataValidation dataValidation);

        /// <summary>
        /// Enable filtering for a range of cells
        /// </summary>
        /// <param name="range">the range of cells to filter</param>
        IAutoFilter SetAutoFilter(CellRangeAddress range);

        /// <summary>
        /// The 'Conditional Formatting' facet for this <c>Sheet</c>
        /// </summary>
        /// <returns>conditional formatting rule for this sheet</returns>
        ISheetConditionalFormatting SheetConditionalFormatting { get; }

        /// <summary>
        /// Whether the text is displayed in right-to-left mode in the window
        /// </summary>
        bool IsRightToLeft { get; set; }

        
        /// <summary>
        ///  Get or set the repeating rows used when printing the sheet, as found in File->PageSetup->Sheet.
        /// <p/>
        /// Repeating rows cover a range of contiguous rows, e.g.:
        /// <pre>
        /// Sheet1!$1:$1
        /// Sheet2!$5:$8
        /// </pre>
        /// The {@link CellRangeAddress} returned contains a column part which spans
        /// all columns, and a row part which specifies the contiguous range of 
        /// repeating rows.
        /// <p/>
        /// If the Sheet does not have any repeating rows defined, null is returned.
        /// </summary>
        CellRangeAddress RepeatingRows { get; set; }


        /// <summary>
        ///  Gets or set the repeating columns used when printing the sheet, as found in File->PageSetup->Sheet.
        /// <p/>
        /// Repeating columns cover a range of contiguous columns, e.g.:
        /// <pre>
        /// Sheet1!$A:$A
        /// Sheet2!$C:$F
        /// </pre>
        /// The {@link CellRangeAddress} returned contains a row part which spans all 
        /// rows, and a column part which specifies the contiguous range of 
        /// repeating columns.
        /// <p/>
        /// If the Sheet does not have any repeating columns defined, null is 
        /// returned.
        /// </summary>
        CellRangeAddress RepeatingColumns { get; set; }

        /// <summary>
        /// Copy sheet with a new name
        /// </summary>
        /// <param name="Name">new sheet name</param>
        /// <returns>cloned sheet</returns>
        ISheet CopySheet(String Name);

        /// <summary>
        /// Copy sheet with a new name
        /// </summary>
        /// <param name="Name">new sheet name</param>
        /// <param name="copyStyle">whether to copy styles</param>
        /// <returns>cloned sheet</returns>
        ISheet CopySheet(String Name, Boolean copyStyle);

        /// <summary>
        /// Returns the column outline level. Increased as you
        /// put it into more groups (outlines), reduced as
        /// you take it out of them.
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        int GetColumnOutlineLevel(int columnIndex);
    }

}