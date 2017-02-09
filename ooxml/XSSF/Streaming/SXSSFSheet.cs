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
        /*package*/
        public XSSFSheet _sh;
        public SXSSFWorkbook _workbook;
        //private TreeMap<Integer, SXSSFRow> _rows = new TreeMap<Integer, SXSSFRow>();
        public SortedDictionary<int, SXSSFRow> _rows = new SortedDictionary<int, SXSSFRow>();
        public SheetDataWriter _writer;
        public int _randomAccessWindowSize = SXSSFWorkbook.DEFAULT_WINDOW_SIZE;
        public AutoSizeColumnTracker _autoSizeColumnTracker;
        public int outlineLevelRow = 0;
        public int lastFlushedRowNumber = -1;
        public bool allFlushed = false;


        public SXSSFSheet(SXSSFWorkbook workbook, XSSFSheet xSheet)
        {
            _workbook = workbook;
            _sh = xSheet;
            _writer = workbook.CreateSheetDataWriter();
            SetRandomAccessWindowSize(_workbook.RandomAccessWindowSize);
            _autoSizeColumnTracker = new AutoSizeColumnTracker(this);
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
                return _rows.Count == 0 ? 0 : _rows.Keys.First();
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
                return _rows.Count == 0 ? 0 : _rows.Keys.Last();
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

        public void AddValidationData(IDataValidation dataValidation)
        {
            _sh.AddValidationData(dataValidation);
        }

        public void AutoSizeColumn(int column)
        {
            AutoSizeColumn(column, false);
        }

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
                flushedWidth = _autoSizeColumnTracker.getBestFitColumnWidth(column, useMergedCells);
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
                int width = Math.Max(bestFitWidth, maxColumnWidth);
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

        public ISheet CopySheet(string Name, bool copyStyle)
        {
            throw new NotImplementedException();
        }

        public IDrawing CreateDrawingPatriarch()
        {
            throw new NotImplementedException();
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
                throw new InvalidOperationException("Invalid row number (" + rownum
                        + ") outside allowable range (0.." + maxrow + ")");
            }

            // attempt to overwrite a row that is already flushed to disk
            if (rownum <= _writer.NumberLastFlushedRow)
            {
                throw new InvalidOperationException(
                        "Attempting to write a row[" + rownum + "] " +
                        "in the range [0," + _writer.NumberLastFlushedRow + "] that is already written to disk.");
            }

            // attempt to overwrite a existing row in the input template
            if (_sh.PhysicalNumberOfRows > 0 && rownum <= _sh.LastRowNum)
            {
                throw new InvalidOperationException(
                        "Attempting to write a row[" + rownum + "] " +
                                "in the range [0," + _sh.LastRowNum + "] that is already written to disk.");
            }

            SXSSFRow newRow = new SXSSFRow(this);
            _rows.Add(rownum, newRow);
            allFlushed = false;
            if (_randomAccessWindowSize >= 0 && _rows.Count > _randomAccessWindowSize)
            {
                try
                {
                    flushRows(_randomAccessWindowSize);
                }
                catch (IOException ioe)
                {
                    throw new RuntimeException(ioe);
                }
            }
            return newRow;
        }

        public void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane)
        {
            _sh.CreateSplitPane(xSplitPos, ySplitPos, leftmostColumn, topRow, activePane);
        }

        public IComment GetCellComment(int row, int column)
        {
            return _sh.GetCellComment(row, column);
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
            return _sh.GetEnumerator();
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
            return _rows[rownum];
        }

        public IEnumerator GetRowEnumerator()
        {
            return _sh.GetRowEnumerator();
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

        public void RemoveRow(IRow row)
        {
            throw new NotImplementedException();
            //if (row.Sheet != this)
            //{
            //    throw new InvalidOperationException("Specified row does not belong to this sheet");
            //}


            //for (Iterator<Map.Entry<Integer, SXSSFRow>> iter = _rows.entrySet().iterator(); iter.hasNext();)
            //{
            //    Map.Entry<int, SXSSFRow> entry = iter.next();
            //    if (entry.getValue() == row)
            //    {
            //        iter.remove();
            //        return;
            //    }
            //}
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

        [Obsolete("in poi 3.16")]
        public void SetZoom(int numerator, int denominator)
        {
            _sh.SetZoom(numerator, denominator);
        }

        public void ShiftRows(int startRow, int endRow, int n)
        {
            throw new NotImplementedException();
        }

        public void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight)
        {
            throw new NotImplementedException();
        }

        public void ShowInPane(short toprow, short leftcol)
        {
            _sh.ShowInPane(toprow, leftcol);
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
            throw new NotImplementedException();
        }

        public void changeRowNum(SXSSFRow row, int newRowNum)
        {

            RemoveRow(row);
            _rows.Add(newRowNum, row);
        }

        public bool dispose()
        {
            if (!allFlushed) flushRows();
            return _writer.Dispose();
        }

        /**
 * Specifies how many rows can be accessed at most via getRow().
 * The exeeding rows (if any) are flushed to the disk while rows
 * with lower index values are flushed first.
 */
        private void flushRows(int remaining)
        {
            while (_rows.Count > remaining) flushOneRow();
            if (remaining == 0) allFlushed = true;
        }

        public void flushRows()
        {
            flushRows(0);
        }

        private void flushOneRow()
        {

            int firstRowNum = _rows.FirstOrDefault().Key;
            if (firstRowNum != null)
            {
                int rowIndex = firstRowNum;
                SXSSFRow row = _rows[firstRowNum];
                // Update the best fit column widths for auto-sizing just before the rows are flushed
                // _autoSizeColumnTracker.UpdateColumnWidths(row);
                _writer.WriteRow(rowIndex, row);
                _rows.Remove(firstRowNum);
                lastFlushedRowNumber = rowIndex;
            }
        }

        /* Gets "<sheetData>" document fragment*/
        public Stream getWorksheetXMLInputStream()
        {
            // flush all remaining data and close the temp file writer
            flushRows(0);
            _writer.Close();
            return _writer.GetWorksheetXmlInputStream();
        }

        public SheetDataWriter getSheetDataWriter()
        {
            return _writer;
        }
    }
}
