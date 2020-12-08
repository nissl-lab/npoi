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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFSheet : ISheet
    {
        // TODO: fields should be private and use public property
        internal XSSFSheet _sh;
        private SXSSFWorkbook _workbook;
        //private TreeMap<Integer, SXSSFRow> _rows = new TreeMap<Integer, SXSSFRow>();
        private IDictionary<int, SXSSFRow> _rows = new Dictionary<int, SXSSFRow>();
        private SheetDataWriter _writer;
        private int _randomAccessWindowSize = SXSSFWorkbook.DEFAULT_WINDOW_SIZE;
        private Lazy<AutoSizeColumnTracker> _autoSizeColumnTracker;
        private int outlineLevelRow = 0;
        private int lastFlushedRowNumber = -1;
        private bool allFlushed = false;

        private int _FirstRowNum = -1;
        private int _LastRowNum = -1;


        public SXSSFSheet(SXSSFWorkbook workbook, XSSFSheet xSheet)
        {
            _workbook = workbook;
            _sh = xSheet;
            _writer = workbook.CreateSheetDataWriter();
            SetRandomAccessWindowSize(_workbook.RandomAccessWindowSize);
            _autoSizeColumnTracker = new Lazy<AutoSizeColumnTracker>(
                () => new AutoSizeColumnTracker(this));
        }

        public void SetRandomAccessWindowSize(int value)
        {
            if (value == 0 || value < -1)
            {
                throw new ArgumentException("RandomAccessWindowSize must be either -1 or a positive integer");
            }
            _randomAccessWindowSize = value;
        }

        public bool Autobreaks
        {
            get
            {
                return _sh.Autobreaks;
            }

            set { _sh.Autobreaks = value; }
        }

        public int[] ColumnBreaks
        {
            get
            {
                return _sh.ColumnBreaks;
            }
            //set { _sh.ColumnBreaks = value; }
        }

        public int DefaultColumnWidth
        {
            get
            {
                return _sh.DefaultColumnWidth;
            }

            set
            {
                _sh.DefaultColumnWidth = value;
            }
        }

        public short DefaultRowHeight
        {
            get { return _sh.DefaultRowHeight; }

            set
            {
                _sh.DefaultRowHeight = value;
            }
        }

        public float DefaultRowHeightInPoints
        {
            get
            {
                return _sh.DefaultRowHeightInPoints;
            }

            set
            {
                _sh.DefaultRowHeightInPoints = value;
            }
        }

        public bool DisplayFormulas
        {
            get
            {
                return _sh.DisplayFormulas;
            }

            set { _sh.DisplayFormulas = value; }
        }

        public bool DisplayGridlines
        {
            get
            {
                return _sh.DisplayGridlines;
            }

            set
            {
                _sh.DisplayGridlines = value;
            }
        }

        public bool DisplayGuts
        {
            get { return _sh.DisplayGuts; }

            set
            {
                _sh.DisplayGuts = value;
            }
        }

        public bool DisplayRowColHeadings
        {
            get
            {
                return _sh.DisplayRowColHeadings;
            }

            set
            {
                _sh.DisplayRowColHeadings = value;
            }
        }

        public bool DisplayZeros
        {
            get
            {
                return _sh.DisplayZeros;
            }

            set
            {
                _sh.DisplayZeros = value;
            }
        }

        public IDrawing DrawingPatriarch
        {
            get
            {
                return _sh.DrawingPatriarch;
            }
        }

        public int FirstRowNum
        {
            get
            {
                if (_writer.NumberOfFlushedRows > 0)
                    return _writer.LowestIndexOfFlushedRows;
                return _rows.Count == 0 ? 0 : _FirstRowNum;
            }
        }

        public bool FitToPage
        {
            get
            {
                return _sh.FitToPage;
            }

            set { _sh.FitToPage = value; }
        }

        public IFooter Footer
        {
            get
            {
                return _sh.Footer;
            }
        }

        public bool ForceFormulaRecalculation
        {
            get
            {
                return _sh.ForceFormulaRecalculation;
            }

            set { _sh.ForceFormulaRecalculation = value; }
        }

        public IHeader Header
        {
            get { return _sh.Header; }
        }

        public bool HorizontallyCenter
        {
            get
            {
                return _sh.HorizontallyCenter;
            }

            set { _sh.HorizontallyCenter = value; }
        }

        public bool IsActive
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool IsPrintGridlines
        {
            get { return _sh.IsPrintGridlines; }

            set
            {
                _sh.IsPrintGridlines = value;
            }
        }
        /**
         * Returns whether row and column headings are printed.
         *
         * @return whether row and column headings are printed
         */
        public bool IsPrintRowAndColumnHeadings
        {
            get
            {
                return _sh.IsPrintRowAndColumnHeadings;
            }
            set
            {
                _sh.IsPrintRowAndColumnHeadings = value;
            }
        }
        public bool IsRightToLeft
        {
            get
            {
                return _sh.IsRightToLeft;
            }

            set
            {
                _sh.IsRightToLeft = value;
            }
        }

        public bool IsSelected
        {
            get
            {
                return _sh.IsSelected;
            }

            set
            {
                _sh.IsSelected = value;
            }
        }

        public int LastRowNum
        {
            get
            {
                if (_rows.Count == 0)
                    return _writer.NumberOfFlushedRows > 0 ? LastFlushedRowNumber : 0;
                return  _LastRowNum;
            }
        }

        public short LeftCol
        {
            get { return _sh.LeftCol; }

            set
            {
                throw new NotImplementedException();
            }
        }

        public int NumMergedRegions
        {
            get { return _sh.NumMergedRegions; }
        }

        /**
         * Returns the list of merged regions. If you want multiple regions, this is
         * faster than calling {@link #getMergedRegion(int)} each time.
         *
         * @return the list of merged regions
         */
        public List<CellRangeAddress> MergedRegions
        {
            get { return _sh.MergedRegions; }
        }
        public PaneInformation PaneInformation
        {
            get { return _sh.PaneInformation; }
        }

        public int PhysicalNumberOfRows
        {
            get
            {
                return _rows.Count + _writer.NumberOfFlushedRows;
            }
        }

        public IPrintSetup PrintSetup
        {
            get { return _sh.PrintSetup; }
        }

        public bool Protect
        {
            get { return _sh.Protect; }
        }

        public CellRangeAddress RepeatingColumns
        {
            get { return _sh.RepeatingColumns; }

            set
            {
                _sh.RepeatingColumns = value;
            }
        }

        public CellRangeAddress RepeatingRows
        {
            get { return _sh.RepeatingRows; }

            set
            {
                _sh.RepeatingRows = value;
            }
        }

        public int[] RowBreaks
        {

            get { return _sh.RowBreaks; }

        }

        public bool RowSumsBelow
        {
            get { return _sh.RowSumsBelow; }

            set
            {
                _sh.RowSumsBelow = value;
            }
        }

        public bool RowSumsRight
        {
            get { return _sh.RowSumsRight; }

            set
            {
                _sh.RowSumsRight = value;
            }
        }

        public bool ScenarioProtect
        {
            get { return _sh.ScenarioProtect; }
        }

        public ISheetConditionalFormatting SheetConditionalFormatting
        {

            get { return _sh.SheetConditionalFormatting; }
        }

        public string SheetName
        {
            get { return _sh.SheetName; }
        }

        public short TabColorIndex
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public short TopRow
        {
            get { return _sh.TopRow; }

            set
            {
                _sh.TopRow = value;
            }
        }

        public bool VerticallyCenter
        {
            get { return _sh.VerticallyCenter; }

            set
            {
                _sh.VerticallyCenter = value;
            }
        }

        public IWorkbook Workbook
        {
            get { return _workbook; }
        }

        public int AddMergedRegion(CellRangeAddress region)
        {
            return _sh.AddMergedRegion(region);
        }

        /// <summary>
        /// Adds a merged region of cells (hence those cells form one).
        /// Skips validation.It is possible to create overlapping merged regions
        /// or create a merged region that intersects a multi-cell array formula
        /// with this formula, which may result in a corrupt workbook.
        /// </summary>
        /// <param name="region">region to merge</param>
        /// <returns>index of this region</returns>
        /// <exception cref="System.ArgumentException">if region contains fewer than 2 cells</exception>
        public int AddMergedRegionUnsafe(CellRangeAddress region)
        {
            return _sh.AddMergedRegionUnsafe(region);
        }

        /**
         * Verify that merged regions do not intersect multi-cell array formulas and
         * no merged regions intersect another merged region in this sheet.
         *
         * @throws InvalidOperationException if region intersects with a multi-cell array formula
         * @throws InvalidOperationException if at least one region intersects with another merged region in this sheet
         */
        public void ValidateMergedRegions() {
            _sh.ValidateMergedRegions();
        }


        public void AddValidationData(IDataValidation dataValidation)
        {
            _sh.AddValidationData(dataValidation);
        }
        /**
         * Adjusts the column width to fit the contents.
         *
         * <p>
         * This process can be relatively slow on large sheets, so this should
         *  normally only be called once per column, at the end of your
         *  processing.
         * </p>
         * You can specify whether the content of merged cells should be considered or ignored.
         *  Default is to ignore merged cells.
         *  
         *  <p>
         *  Special note about SXSSF implementation: You must register the columns you wish to track with
         *  the SXSSFSheet using {@link #trackColumnForAutoSizing(int)} or {@link #trackAllColumnsForAutoSizing()}.
         *  This is needed because the rows needed to compute the column width may have fallen outside the
         *  random access window and been flushed to disk.
         *  Tracking columns is required even if all rows are in the random access window.
         *  </p>
         *  <p><i>New in POI 3.14 beta 1: auto-sizes columns using cells from current and flushed rows.</i></p>
         *
         * @param column the column index to auto-size
         */
        public void AutoSizeColumn(int column)
        {
            AutoSizeColumn(column, false);
        }

        /**
         * Adjusts the column width to fit the contents.
         * <p>
         * This process can be relatively slow on large sheets, so this should
         *  normally only be called once per column, at the end of your
         *  processing.
         * </p>
         * You can specify whether the content of merged cells should be considered or ignored.
         *  Default is to ignore merged cells.
         *  
         *  <p>
         *  Special note about SXSSF implementation: You must register the columns you wish to track with
         *  the SXSSFSheet using {@link #trackColumnForAutoSizing(int)} or {@link #trackAllColumnsForAutoSizing()}.
         *  This is needed because the rows needed to compute the column width may have fallen outside the
         *  random access window and been flushed to disk.
         *  Tracking columns is required even if all rows are in the random access window.
         *  </p>
         *  <p><i>New in POI 3.14 beta 1: auto-sizes columns using cells from current and flushed rows.</i></p>
         *
         * @param column the column index to auto-size
         * @param useMergedCells whether to use the contents of merged cells when calculating the width of the column
         */
        public void AutoSizeColumn(int column, bool useMergedCells)
        {
            // Multiple calls to autoSizeColumn need to look up the best-fit width
            // of rows already flushed to disk plus re-calculate the best-fit width
            // of rows in the current window. It isn't safe to update the column
            // widths before flushing to disk because columns in the random access
            // window rows may change in best-fit width. The best-fit width of a cell
            // is only fixed when it becomes inaccessible for modification.
            // Changes to the shared strings table, styles table, or formulas might
            // be able to invalidate the auto-size width without the opportunity
            // to recalculate the best-fit width for the flushed rows. This is an
            // inherent limitation of SXSSF. If having correct auto-sizing is
            // critical, the flushed rows would need to be re-read by the read-only
            // XSSF eventmodel (SAX) or the memory-heavy XSSF usermodel (DOM). 
            int flushedWidth;
            try
            {
                // get the best fit width of rows already flushed to disk
                flushedWidth = _autoSizeColumnTracker.Value.GetBestFitColumnWidth(column, useMergedCells);
            }
            catch (Exception e)
            {
                throw new InvalidOperationException("Could not auto-size column. Make sure the column was tracked prior to auto-sizing the column.", e);
            }

            // get the best-fit width of rows currently in the random access window
            int activeWidth = (int)(256 * SheetUtil.GetColumnWidth(this, column, useMergedCells));

            // the best-fit width for both flushed rows and random access window rows
            // flushedWidth or activeWidth may be negative if column contains only blank cells
            int bestFitWidth = Math.Max(flushedWidth, activeWidth);

            if (bestFitWidth > 0)
            {
                int maxColumnWidth = 255 * 256; // The maximum column width for an individual cell is 255 characters
                int width = Math.Min(bestFitWidth, maxColumnWidth);
                SetColumnWidth(column, width);
            }
        }

        public IRow CopyRow(int sourceIndex, int targetIndex)
        {
            throw new NotImplementedException();
        }

        public ISheet CopySheet(string Name)
        {
            throw new NotImplementedException();
        }

        public ISheet CopySheet(string Name,string newName, bool copyStyle)
        {
            throw new NotImplementedException();
        }

        public ISheet CopySheet(string Name, bool copyStyle)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get a Hyperlink in this sheet anchored at row, column
        /// </summary>
        /// <param name="row">The index of the row of the hyperlink, zero-based</param>
        /// <param name="column">the index of the column of the hyperlink, zero-based</param>
        /// <returns>return hyperlink if there is a hyperlink anchored at row, column; otherwise returns null</returns>
        public IHyperlink GetHyperlink(int row, int column)
        {
            return _sh.GetHyperlink(row, column);
        }
        /// <summary>
        /// Get a Hyperlink in this sheet located in a cell specified by {code addr}
        /// </summary>
        /// <param name="addr">The address of the cell containing the hyperlink</param>
        /// <returns>return hyperlink if there is a hyperlink anchored at {@code addr}; otherwise returns {@code null}</returns>
        public IHyperlink GetHyperlink(CellAddress addr)
        {
            return _sh.GetHyperlink(addr);
        }
        /**
         * Get a list of Hyperlinks in this sheet
         *
         * @return Hyperlinks for the sheet
         */

        public List<IHyperlink> GetHyperlinkList()
        {
            return _sh.GetHyperlinkList();
        }

        public IDrawing CreateDrawingPatriarch()
        {
            return _sh.CreateDrawingPatriarch();
        }

        public void CreateFreezePane(int colSplit, int rowSplit)
        {
            _sh.CreateFreezePane(colSplit, rowSplit);
        }

        public void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
        {
            _sh.CreateFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
        }

        public IRow CreateRow(int rownum)
        {
            int maxrow = SpreadsheetVersion.EXCEL2007.LastRowIndex;
            if (rownum < 0 || rownum > maxrow)
            {
                throw new ArgumentException("Invalid row number (" + rownum
                        + ") outside allowable range (0.." + maxrow + ")");
            }

            // attempt to overwrite a row that is already flushed to disk
            if (rownum <= _writer.NumberLastFlushedRow)
            {
                throw new ArgumentException(
                        "Attempting to write a row[" + rownum + "] " +
                        "in the range [0," + _writer.NumberLastFlushedRow + "] that is already written to disk.");
            }

            // attempt to overwrite a existing row in the input template
            if (_sh.PhysicalNumberOfRows > 0 && rownum <= _sh.LastRowNum)
            {
                throw new ArgumentException(
                        "Attempting to write a row[" + rownum + "] " +
                                "in the range [0," + _sh.LastRowNum + "] that is already written to disk.");
            }

            SXSSFRow newRow = new SXSSFRow(this);
            _rows[rownum] = newRow;

            UpdateIndexWhenAdd(rownum);

            allFlushed = false;
            if (_randomAccessWindowSize >= 0 && _rows.Count > _randomAccessWindowSize)
            {
                try
                {
                    FlushRows(_randomAccessWindowSize, false);
                }
                catch (IOException ioe)
                {
                    throw new RuntimeException(ioe);
                }
            }
            return newRow;
        }

        private void UpdateIndexWhenAdd(int rownum)
        {
            if (_FirstRowNum == -1 || rownum < _FirstRowNum)
            {
                _FirstRowNum = rownum;
            }

            if (rownum > _LastRowNum)
            {
                _LastRowNum = rownum;
            }
        }

        public void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane)
        {
            _sh.CreateSplitPane(xSplitPos, ySplitPos, leftmostColumn, topRow, activePane);
        }

        /// <summary>
        /// Returns cell comment for the specified row and column
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <returns>cell comment or <code>null</code> if not found</returns>
        [Obsolete("deprecated as of 2015-11-23 (circa POI 3.14beta1). Use {@link #getCellComment(CellAddress)} instead.")]
        public IComment GetCellComment(int row, int column)
        {
            return GetCellComment(new CellAddress(row, column));
        }
        /// <summary>
        /// Returns cell comment for the specified location
        /// </summary>
        /// <param name="ref1">cell location</param>
        /// <returns>return cell comment or null if not found</returns>
        public IComment GetCellComment(CellAddress ref1)
        {
            return _sh.GetCellComment(ref1);
        }

        /// <summary>
        /// Returns all cell comments on this sheet.
        /// </summary>
        /// <returns>return A Dictionary of each Comment in the sheet, keyed on the cell address where the comment is located.</returns>
        public Dictionary<CellAddress, IComment> GetCellComments()
        {
            return _sh.GetCellComments();
        }

        public int GetColumnOutlineLevel(int columnIndex)
        {
            return _sh.GetColumnOutlineLevel(columnIndex);
        }

        public ICellStyle GetColumnStyle(int column)
        {
            return _sh.GetColumnStyle(column);
        }

        public int GetColumnWidth(int columnIndex)
        {
            return _sh.GetColumnWidth(columnIndex);
        }

        public float GetColumnWidthInPixels(int columnIndex)
        {
            return _sh.GetColumnWidthInPixels(columnIndex);
        }


        public IDataValidationHelper GetDataValidationHelper()
        {
            return _sh.GetDataValidationHelper();
        }

        public List<IDataValidation> GetDataValidations()
        {
            return _sh.GetDataValidations();
        }

        public IEnumerator GetEnumerator()
        {
            return (IEnumerator<IRow>)new SortedDictionary<int,SXSSFRow>(_rows).Values.GetEnumerator();
        }

        public double GetMargin(MarginType margin)
        {
            return _sh.GetMargin(margin);
        }

        public CellRangeAddress GetMergedRegion(int index)
        {
            return _sh.GetMergedRegion(index);
        }

        public IRow GetRow(int rownum)
        {
            if (_rows.ContainsKey(rownum))
                return _rows[rownum];
            else
                return null;
        }

        public IEnumerator GetRowEnumerator()
        {
            return GetEnumerator();
        }

        public void GroupColumn(int fromColumn, int toColumn)
        {
            _sh.GroupColumn(fromColumn, toColumn);
        }


        //TODO: test
        public void GroupRow(int fromRow, int toRow)
        {
            var groupRows = _rows.Where(kvp => kvp.Key >= fromRow && kvp.Key <= toRow + 1).Select(r => r.Value);
            foreach (SXSSFRow row in groupRows)
            {
                int level = row.OutlineLevel + 1;
                row.OutlineLevel = level;

                if (level > outlineLevelRow) outlineLevelRow = level;
            }

            SetWorksheetOutlineLevelRow();
        }

        /**
 * Set row groupings (like groupRow) in a stream-friendly manner
 *
 * <p>
 *    groupRows requires all rows in the group to be in the current window.
 *    This is not always practical.  Instead use setRowOutlineLevel to 
 *    explicitly set the group level.  Level 1 is the top level group, 
 *    followed by 2, etc.  It is up to the user to ensure that level 2
 *    groups are correctly nested under level 1, etc.
 * </p>
 *
 * @param rownum    index of row to update (0-based)
 * @param level     outline level (greater than 0)
 */
        public void SetRowOutlineLevel(int rownum, int level)
        {
            SXSSFRow row = _rows[rownum];

            row.OutlineLevel = level;
            if (level > 0 && level > outlineLevelRow)
            {
                outlineLevelRow = level;
                SetWorksheetOutlineLevelRow();
            }
        }

        private void SetWorksheetOutlineLevelRow()
        {
            var ct = _sh.GetCTWorksheet();
            var pr = ct.IsSetSheetFormatPr() ?
                ct.sheetFormatPr :
                ct.AddNewSheetFormatPr();
            if (outlineLevelRow > 0) pr.outlineLevelRow = (byte)outlineLevelRow;
        }

        public bool IsColumnBroken(int column)
        {
            return _sh.IsColumnBroken(column);
        }

        public bool IsColumnHidden(int columnIndex)
        {
            return _sh.IsColumnHidden(columnIndex);
        }

        public bool IsMergedRegion(CellRangeAddress mergedRegion)
        {
            throw new NotImplementedException();
        }

        public bool IsRowBroken(int row)
        {
            return _sh.IsRowBroken(row);
        }

        public void ProtectSheet(string password)
        {
            _sh.ProtectSheet(password);
        }

        public ICellRange<ICell> RemoveArrayFormula(ICell cell)
        {
            return _sh.RemoveArrayFormula(cell);
        }

        public void RemoveColumnBreak(int column)
        {
            _sh.RemoveColumnBreak(column);
        }

        public void RemoveMergedRegion(int index)
        {
            _sh.RemoveMergedRegion(index);
        }
        /**
         * Removes a merged region of cells (hence letting them free)
         *
         * @param indices of the regions to unmerge
         */
        public void RemoveMergedRegions(IList<int> indices)
        {
            _sh.RemoveMergedRegions(indices);
        }
        public void RemoveRow(IRow row)
        {
            if (row == null)
            {
                throw new ArgumentException("Invalid row (null)");
            }
            if (row.Sheet != this)
            {
                throw new ArgumentException("Specified row does not belong to this sheet");
            }
            List<int> toRemove = new List<int>();
            foreach(var kv in _rows)
            {
                if(kv.Value == row)
                {
                    toRemove.Add(kv.Key);
                }
            }

            var invalidatedFirst = false;
            var invalidatedLast = false;
            foreach(var key in toRemove)
            {
                if (key == _FirstRowNum)
                {
                    invalidatedFirst = true;
                }

                if (key >= (_LastRowNum -1))
                {
                    invalidatedLast = true;
                }
                _rows.Remove(key);
            }

            if (invalidatedFirst)
            {
                InvalidateFirstRowNum();
            }

            if (invalidatedLast)
            {
                InvalidateLastRowNum();
            }
            
        }

        private void InvalidateFirstRowNum()
        {
            if (_rows.Count == 0)
            {
                _FirstRowNum = -1;
            }
            else
            {
                _FirstRowNum = _rows.Keys.Min();
            }
        }
        
        private void InvalidateLastRowNum()
        {
            if (_rows.Count == 0)
            {
                _LastRowNum = -1;
            }
            else
            {
                _LastRowNum = _rows.Keys.Max();
            }
        }

        public void RemoveRowBreak(int row)
        {
            _sh.RemoveRowBreak(row);
        }

        public void SetActive(bool value)
        {
            throw new NotImplementedException();
        }

        public void SetActiveCell(int row, int column)
        {
            throw new NotImplementedException();
        }

        public void SetActiveCellRange(List<CellRangeAddress8Bit> cellranges, int activeRange, int activeRow, int activeColumn)
        {
            throw new NotImplementedException();
        }

        public void SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            throw new NotImplementedException();
        }

        public ICellRange<ICell> SetArrayFormula(string formula, CellRangeAddress range)
        {
            return _sh.SetArrayFormula(formula, range);
        }

        public IAutoFilter SetAutoFilter(CellRangeAddress range)
        {
            return _sh.SetAutoFilter(range);
        }

        public void SetColumnBreak(int column)
        {
            _sh.SetColumnBreak(column);
        }

        public void SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            _sh.SetColumnGroupCollapsed(columnNumber, collapsed);
        }

        public void SetColumnHidden(int columnIndex, bool hidden)
        {
            _sh.SetColumnHidden(columnIndex, hidden);
        }

        public void SetColumnWidth(int columnIndex, int width)
        {
            _sh.SetColumnWidth(columnIndex, width);
        }

        public void SetDefaultColumnStyle(int column, ICellStyle style)
        {
            _sh.SetDefaultColumnStyle(column, style);
        }


        /**
         * Track a column in the sheet for auto-sizing.
         * Note this has undefined behavior if a column is tracked after one or more rows are written to the sheet.
         * If <code>column</code> is already tracked, this call does nothing.
         *
         * @param column the column to track for autosizing
         * @since 3.14beta1
         * @see #trackColumnsForAutoSizing(Collection)
         * @see #trackAllColumnsForAutoSizing()
         */
        public void TrackColumnForAutoSizing(int column)
        {
            _autoSizeColumnTracker.Value.TrackColumn(column);
        }

        /**
         * Track several columns in the sheet for auto-sizing.
         * Note this has undefined behavior if columns are tracked after one or more rows are written to the sheet.
         * Any column in <code>columns</code> that are already tracked are ignored by this call.
         *
         * @param columns the columns to track for autosizing
         * @since 3.14beta1
         */
        public void TrackColumnsForAutoSizing(ICollection<int> columns)
        {
            _autoSizeColumnTracker.Value.TrackColumns(columns);
        }

        /**
         * Tracks all columns in the sheet for auto-sizing. If this is called, individual columns do not need to be tracked.
         * Because determining the best-fit width for a cell is expensive, this may affect the performance.
         * @since 3.14beta1
         */
        public void TrackAllColumnsForAutoSizing()
        {
            _autoSizeColumnTracker.Value.TrackAllColumns();
        }

        /**
         * Removes a column that was previously marked for inclusion in auto-size column tracking.
         * When a column is untracked, the best-fit width is forgotten.
         * If <code>column</code> is not tracked, it will be ignored by this call.
         *
         * @param column the index of the column to track for auto-sizing
         * @return true if column was tracked prior to this call, false if no action was taken
         * @since 3.14beta1
         * @see #untrackColumnsForAutoSizing(Collection)
         * @see #untrackAllColumnsForAutoSizing(int)
         */
        public bool UntrackColumnForAutoSizing(int column)
        {
            return _autoSizeColumnTracker.Value.UntrackColumn(column);
        }

        /**
         * Untracks several columns in the sheet for auto-sizing.
         * When a column is untracked, the best-fit width is forgotten.
         * Any column in <code>columns</code> that is not tracked will be ignored by this call.
         *
         * @param columns the indices of the columns to track for auto-sizing
         * @return true if one or more columns were untracked as a result of this call
         *
         * @param columns the columns to track for autosizing
         * @since 3.14beta1
         */
        public bool UntrackColumnsForAutoSizing(ICollection<int> columns)
        {
            return _autoSizeColumnTracker.Value.UntrackColumns(columns);
        }

        /**
         * Untracks all columns in the sheet for auto-sizing. Best-fit column widths are forgotten.
         * If this is called, individual columns do not need to be untracked.
         * @since 3.14beta1
         */
        public void UntrackAllColumnsForAutoSizing()
        {
            _autoSizeColumnTracker.Value.UntrackAllColumns();
        }

        /**
         * Returns true if column is currently tracked for auto-sizing.
         *
         * @param column the index of the column to check
         * @return true if column is tracked
         * @since 3.14beta1
         */
        public bool IsColumnTrackedForAutoSizing(int column)
        {
            return _autoSizeColumnTracker.Value.IsColumnTracked(column);
        }

        /**
         * Get the currently tracked columns for auto-sizing.
         * Note if all columns are tracked, this will only return the columns that have been explicitly or implicitly tracked,
         * which is probably only columns containing 1 or more non-blank values
         *
         * @return a set of the indices of all tracked columns
         * @since 3.14beta1
         */
        public ISet<int> TrackedColumnsForAutoSizing
        {
            get
            {
                return _autoSizeColumnTracker.Value.TrackedColumns;
            }
        }

        public void SetMargin(MarginType margin, double size)
        {
            _sh.SetMargin(margin, size);
        }

        public void SetRowBreak(int row)
        {
            _sh.SetRowBreak(row);
        }

        public void SetRowGroupCollapsed(int row, bool collapse)
        {
            if (collapse)
            {
                collapseRow(row);
            }
            else
            {
                //expandRow(rowIndex);
                throw new RuntimeException("Unable to expand row: Not Implemented");
            }
        }

        private void collapseRow(int rowIndex)
        {
            SXSSFRow row = (SXSSFRow)GetRow(rowIndex);
            if (row == null)
            {
                throw new InvalidOperationException("Invalid row number(" + rowIndex + "). Row does not exist.");
            }
            else
            {
                int startRow = FindStartOfRowOutlineGroup(rowIndex);

                // Hide all the columns until the end of the group
                int lastRow = WriteHidden(row, startRow, true);
                SXSSFRow lastRowObj = (SXSSFRow)GetRow(lastRow);
                if (lastRowObj != null)
                {
                    lastRowObj.Collapsed = true;
                }
                else
                {
                    SXSSFRow newRow = (SXSSFRow)CreateRow(lastRow);
                    newRow.Collapsed = true;
                }
            }
        }

        /**
         * @param rowIndex the zero based row index to find from
         */
        private int FindStartOfRowOutlineGroup(int rowIndex)
        {
            // Find the start of the group.
            IRow row = GetRow(rowIndex);
            int level = ((SXSSFRow)row).OutlineLevel;
            if (level == 0)
            {
                throw new InvalidOperationException("Outline level is zero for the row (" + rowIndex + ").");
            }
            int currentRow = rowIndex;
            while (GetRow(currentRow) != null)
            {
                if (GetRow(currentRow).OutlineLevel < level)
                    return currentRow + 1;
                currentRow--;
            }
            return currentRow + 1;
        }

        private int WriteHidden(SXSSFRow xRow, int rowIndex, bool hidden)
        {
            int level = xRow.OutlineLevel;
            var currRow = (SXSSFRow)GetRow(rowIndex);

            while (currRow != null && currRow.OutlineLevel >= level)
            {
                currRow.Hidden = hidden;
                rowIndex++;
                currRow = (SXSSFRow)GetRow(rowIndex);
            }
            return rowIndex;
        }

        [Obsolete("deprecated 2015-11-23 (circa POI 3.14beta1). Use {@link #setZoom(int)} instead.")]
        public void SetZoom(int numerator, int denominator)
        {
            _sh.SetZoom(numerator, denominator);
        }

        /**
         * Window zoom magnification for current view representing percent values.
         * Valid values range from 10 to 400. Horizontal & Vertical scale together.
         *
         * For example:
         * <pre>
         * 10 - 10%
         * 20 - 20%
         * ...
         * 100 - 100%
         * ...
         * 400 - 400%
         * </pre>
         *
         * Current view can be Normal, Page Layout, or Page Break Preview.
         *
         * @param scale window zoom magnification
         * @throws IllegalArgumentException if scale is invalid
         */

        public void SetZoom(int scale)
        {
            _sh.SetZoom(scale);
        }

        public void ShiftRows(int startRow, int endRow, int n)
        {
            throw new NotImplementedException();
        }

        public void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight)
        {
            throw new NotImplementedException();
        }

        public void ShowInPane(int toprow, int leftcol)
        {
            _sh.ShowInPane(toprow, leftcol);
        }

        public void UngroupColumn(int fromColumn, int toColumn)
        {
            _sh.UngroupColumn(fromColumn, toColumn);

        }

        public void UngroupRow(int fromRow, int toRow)
        {
            _sh.UngroupRow(fromRow, toRow);
        }

        public bool IsDate1904()
        {
            return _workbook.IsDate1904();
        }
        public int GetRowNum(SXSSFRow row)
        {
            foreach (KeyValuePair<int, SXSSFRow> entry in _rows)
            {
                if (entry.Value == row)
                    return entry.Key;
            }
            return -1;
        }
        public void ChangeRowNum(SXSSFRow row, int newRowNum)
        {

            RemoveRow(row);
            _rows.Add(newRowNum, row);
            UpdateIndexWhenAdd(newRowNum);
        }

        public bool Dispose()
        {
            if (!allFlushed) FlushRows();
            return _writer.Dispose();
        }
        /**
 * Specifies how many rows can be accessed at most via getRow().
 * The exceeding rows (if any) are flushed to the disk while rows
 * with lower index values are flushed first.
 */
        private void FlushRows(int remaining, bool flushOnDisk)
        {
            KeyValuePair<int, SXSSFRow>? lastRow = null;
            var flushedRowsCount = 0;
            
            while (_rows.Count > remaining)
            {
                flushedRowsCount++;
                lastRow = flushOneRow();
            }
            
            InvalidateFirstRowNum();
            InvalidateLastRowNum();

            if (remaining == 0) 
                allFlushed = true;

            //TODO: review this.
            if (lastRow != null && flushOnDisk)
                _writer.FlushRows(flushedRowsCount, lastRow.Value.Key, lastRow.Value.Value.LastCellNum);
       }
        /*
         * Are all rows flushed to disk?
         */
        public bool AllRowsFlushed
        {
            get
            {
                return allFlushed;
            }
        }
        /**
         * @return Last row number to be flushed to disk, or -1 if none flushed yet
         */
        public int LastFlushedRowNumber
        {
            get
            {
                return lastFlushedRowNumber;
            }
        }

        public void FlushRows()
        {
            FlushRows(0, true);
        }

        private KeyValuePair<int, SXSSFRow>? flushOneRow()
        {
            if (_rows.Count == 0)
                return null;

            var firstRowNum = _rows.Keys.Min();
            // Update the best fit column widths for auto-sizing just before the rows are flushed
            // _autoSizeColumnTracker.UpdateColumnWidths(row);
            var firstRow = _rows[firstRowNum];
            _writer.WriteRow(firstRowNum, firstRow);
            _rows.Remove(firstRowNum);
            lastFlushedRowNumber = firstRowNum;
            return new KeyValuePair<int, SXSSFRow>(firstRowNum,firstRow);
        }

        /* Gets "<sheetData>" document fragment*/
        public Stream GetWorksheetXMLInputStream()
        {
            // flush all remaining data and close the temp file writer
            FlushRows(0, true);

            _writer.Close();
            return _writer.GetWorksheetXmlInputStream();
        }

        public SheetDataWriter SheetDataWriter
        {
            get
            {
                return _writer;
            }
        }

        public CellAddress ActiveCell
        {
            get
            {
                return _sh.ActiveCell;
            }
            set
            {
                _sh.ActiveCell = value;
            }
        }

        public XSSFColor TabColor
        {
            get
            {
                return _sh.TabColor;
            }
            set
            {
                _sh.TabColor = value;
            }
        }
        public void CopyTo(IWorkbook dest, string name, bool copyStyle, bool keepFormulas)
        {
            throw new NotImplementedException();
        }
    }
}
