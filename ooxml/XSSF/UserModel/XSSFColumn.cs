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

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel.Helpers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.XSSF.UserModel
{
    public class XSSFColumn : IColumn, IComparable<XSSFColumn>
    {
        #region Private properties
        private static readonly POILogger _logger =
            POILogFactory.GetLogger(typeof(XSSFColumn));

        /// <summary>
        /// the xml node Containing defInition for this column
        /// </summary>
        private readonly CT_Col _column;

        /// <summary>
        /// the parent sheet
        /// </summary>
        private readonly XSSFSheet _sheet;

        private readonly StylesTable _stylesSource;
        #endregion

        #region Public properties
        /// <summary>
        /// XSSFSheet this column belongs to
        /// </summary>
        public ISheet Sheet
        {
            get
            {
                return _sheet;
            }
        }

        /// <summary>
        /// Get the number of the first cell Contained in this column.
        /// </summary>
        /// <returns>short representing the first logical cell in the column,
        /// or -1 if the column does not contain any cells.</returns>
        public short FirstCellNum
        {
            get
            {
                return (short)(Cells.Count == 0 ? -1 : GetFirstKey());
            }
        }

        /// <summary>
        /// Gets the index of the last cell Contained in this column <b>PLUS 
        /// ONE</b>. The result also happens to be the 1-based row number of 
        /// the last cell. This value can be used as a standard upper bound 
        /// when iterating over cells:
        /// </summary>
        /// <returns>short representing the last logical cell in the column 
        /// <b>PLUS ONE</b>, or -1 if the column does not contain any cells.</returns>
        public short LastCellNum
        {
            get
            {
                return (short)(Cells.Count == 0 ? -1 : GetLastKey() + 1);
            }
        }

        /// <summary>
        /// Get the column's width. 
        /// If the width is not Set, the default worksheet value is returned,
        /// See <see cref="XSSFSheet.DefaultColumnWidth"/>
        /// </summary>
        /// <returns>column width</returns>
        public double Width
        {
            get
            {
                return _column.width == 0
                    ? _sheet.DefaultColumnWidth
                    : _column.width;
            }

            set
            {
                if (value < 0)
                {
                    if (_column.IsSetWidth())
                    {
                        _column.UnsetWidth();
                    }

                    if (_column.IsSetCustomWidth())
                    {
                        _column.UnsetCustomWidth();
                    }
                }
                else
                {
                    _column.width = value;
                    _column.customWidth = true;

                }
            }
        }

        /// <summary>
        /// Gets the number of defined cells (NOT number of cells in the actual
        /// column!). That is to say if only rows 0,4,5 have values then 
        /// there would be 3.
        /// </summary>
        /// <returns>int representing the number of defined cells in the column.</returns>
        public int PhysicalNumberOfCells
        {
            get
            {
                return Cells.Count;
            }
        }

        /// <summary>
        /// Get column number this column represents
        /// </summary>
        /// <returns>the column number (0 based)</returns>
        public int ColumnNum
        {
            get
            {
                return (int)_column.min - 1;
            }

            set
            {
                int maxColumn = SpreadsheetVersion.EXCEL2007.LastColumnIndex;
                if (value < 0 || value > maxColumn)
                {
                    throw new ArgumentException("Invalid column number (" + value
                            + ") outside allowable range (0.." + maxColumn + ")");
                }

                // As current implementation of XSSFColumn is breaking CT_Col 
                // objects that span over multiple columns into individual
                // columns we need to set min and max values to the same value.
                _column.min = (uint)(value + 1);
                _column.max = (uint)(value + 1);
            }
        }

        /// <summary>
        /// Get whether or not to display this column with 0 width
        /// </summary>
        public bool ZeroWidth
        {
            get
            {
                return _column.hidden;
            }

            set
            {
                _column.hidden = value;
            }
        }

        /// <summary>
        /// Is this column formatted? Most aren't, but some columns
        /// do have whole-column styles. For those that do, you
        /// can get the formatting from <see cref="ColumnStyle"/>
        /// </summary>
        public bool IsFormatted
        {
            get
            {
                return _column.IsSetStyle();
            }
        }

        /// <summary>
        /// Is the column width set to best fit the content?
        /// </summary>
        public bool IsBestFit
        {
            get
            {
                return _column.bestFit;
            }

            set
            {
                _column.bestFit = value;
            }
        }

        /// <summary>
        /// Returns the whole-column cell style. Most columns won't
        /// have one of these, so will return null. Call
        /// <see cref="IsFormatted"/> to check first.
        /// </summary>
        public ICellStyle ColumnStyle
        {
            get
            {
                return IsFormatted
                    && _stylesSource != null
                    && _stylesSource.NumCellStyles > 0
                    ? _stylesSource.GetStyleAt((int)_column.style)
                    : (ICellStyle)null;
            }

            set
            {
                if (value == null)
                {
                    if (_column.IsSetStyle())
                    {
                        _column.UnsetStyle();
                    }
                }
                else
                {
                    XSSFCellStyle xStyle = (XSSFCellStyle)value;
                    xStyle.VerifyBelongsToStylesSource(_stylesSource);

                    long idx = _stylesSource.PutStyle(xStyle);
                    _column.style = (uint)idx;
                }

                foreach (ICell cell in Cells)
                {
                    cell.CellStyle = value;
                }
            }
        }

        public List<ICell> Cells
        {
            get
            {
                List<ICell> cells = new List<ICell>();

                foreach (IRow row in _sheet.Cast<IRow>())
                {
                    foreach (ICell cell in row.Where(c => c.ColumnIndex == ColumnNum))
                    {
                        cells.Add(cell);
                        _sheet.OnReadCell((XSSFCell)cell);
                    }
                }

                return cells;
            }
        }

        public int OutlineLevel
        {
            get
            {
                return _column.outlineLevel;
            }

            set
            {
                _column.outlineLevel = (byte)value;
            }
        }

        public bool Hidden
        {
            get
            {
                return _column.hidden;
            }

            set
            {
                _column.hidden = value;
            }
        }

        public bool Collapsed
        {
            get
            {
                return _column.collapsed;
            }

            set
            {
                _column.collapsed = value;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Construct an XSSFColumn.
        /// </summary>
        /// <param name="column">the xml node Containing defInitions for this column.</param>
        /// <param name="sheet">the parent sheet.</param>
        public XSSFColumn(CT_Col column, XSSFSheet sheet)
        {
            _column = column;
            _sheet = sheet;

            if (!column.IsSetNumber())
            {
                // Certain file format writers skip the column number
                // Assume no gaps, and give this the next column number
                int nextColumnNum = sheet.LastColumnNum + 2;
                if (nextColumnNum == 2 && sheet.PhysicalNumberOfColumns == 0)
                {
                    nextColumnNum = 1;
                }

                // As current implementation of XSSFColumn is breaking CT_Col 
                // objects that span over multiple columns into individual
                // columns we need to set min and max values to the same value.
                _column.min = (uint)nextColumnNum + 1;
                _column.max = (uint)nextColumnNum + 1;
            }

            _stylesSource = ((XSSFWorkbook)sheet.Workbook).GetStylesSource();
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Use this to create new cells within the column and return it. The 
        /// cell that is returned is a <see cref="CellType.Blank"/>. The type 
        /// can be Changed either through calling 
        /// <see cref="ICell.SetCellValue"/> or <see cref="ICell.SetCellType"/>.
        /// </summary>
        /// <param name="rowIndex">the row number this cell represents</param>
        /// <returns>a high level representation of the Created cell</returns>
        /// <exception cref="ArgumentOutOfRangeException">if columnIndex is 
        /// less than 0 or greater than 16384, the maximum number of columns 
        /// supported by the SpreadsheetML format(.xlsx)</exception>
        public ICell CreateCell(int rowIndex)
        {
            return CreateCell(rowIndex, CellType.Blank);
        }

        /// <summary>
        /// Use this to create new cells within the column and return it.
        /// </summary>
        /// <param name="rowIndex">the row number this cell represents</param>
        /// <param name="type">the cell's data type</param>
        /// <returns>a high level representation of the Created cell.</returns>
        /// <exception cref="ArgumentOutOfRangeException">if columnIndex is 
        /// less than 0 or greater than 16384, the maximum number of columns 
        /// supported by the SpreadsheetML format(.xlsx)</exception>
        public ICell CreateCell(int rowIndex, CellType type)
        {
            if (rowIndex > SpreadsheetVersion.EXCEL2007.LastRowIndex ||
                rowIndex < 0)
            {
                throw new ArgumentOutOfRangeException(
                    "Row index should be between 0 and " +
                    SpreadsheetVersion.EXCEL2007.LastRowIndex + ". Input was " +
                    rowIndex);
            }

            ValidateCellType(type);

            IRow row = Sheet.GetRow(rowIndex) ?? Sheet.CreateRow(rowIndex);
            ICell newCell = row.CreateCell(ColumnNum, type);

            if (IsFormatted)
            {
                newCell.CellStyle = ColumnStyle;
            }

            return newCell;
        }

        /// <summary>
        /// Returns the cell at the given (0 based) index,
        /// with the <see cref="MissingCellPolicy"/> from the parent Workbook.
        /// </summary>
        /// <param name="cellNum"></param>
        /// <returns>the cell at the given (0 based) index</returns>
        public ICell GetCell(int cellNum)
        {
            return GetCell(cellNum, _sheet.Workbook.MissingCellPolicy);
        }

        /// <summary>
        /// Returns the cell at the given (0 based) index, with the specified 
        /// <see cref="MissingCellPolicy"/>
        /// </summary>
        /// <param name="cellNum"></param>
        /// <param name="policy"></param>
        /// <returns>the cell at the given (0 based) index</returns>
        /// <exception cref="ArgumentException">if cellnum is less than 0 or 
        /// the specified MissingCellPolicy is invalid</exception>
        public ICell GetCell(int cellNum, MissingCellPolicy policy)
        {
            if (cellNum < 0 || cellNum > SpreadsheetVersion.EXCEL2007.LastRowIndex)
            {
                throw new ArgumentOutOfRangeException(
                    "Cell number should be between 0 and " +
                    SpreadsheetVersion.EXCEL2007.LastRowIndex + ". Input was " +
                    cellNum);
            }

            XSSFCell cell = (XSSFCell)RetrieveCell(cellNum);

            switch (policy)
            {
                case MissingCellPolicy.RETURN_NULL_AND_BLANK:
                    return cell;
                case MissingCellPolicy.RETURN_BLANK_AS_NULL:
                    bool isBlank = cell != null && cell.CellType == CellType.Blank;
                    return isBlank ? null : cell;
                case MissingCellPolicy.CREATE_NULL_AS_BLANK:
                    return cell ?? CreateCell(cellNum, CellType.Blank);
                default:
                    throw new ArgumentException("Illegal policy " + policy + " (" + policy + ")");
            }
        }

        /// <summary>
        /// Remove the Cell from this column.
        /// </summary>
        /// <param name="cell">the cell to remove</param>
        /// <exception cref="ArgumentException"></exception>
        public void RemoveCell(ICell cell)
        {
            if (cell == null)
            {
                throw new ArgumentException("Cell can not be null");
            }

            if (cell.ColumnIndex != ColumnNum)
            {
                throw new ArgumentException("Specified cell does not belong " +
                    "to this column");
            }

            IRow row = _sheet.GetRow(cell.RowIndex);

            row.RemoveCell(cell);
        }

        /// <summary>
        /// Copy the cells from srcColumn to this column If this column is not 
        /// a blank column, this will merge the two columns, overwriting the 
        /// cells in this column with the cells in srcColumn If srcColumn is 
        /// null, overwrite cells in destination column with blank values, 
        /// styles, etc per cell copy policy srcColumn may be from a different 
        /// sheet in the same workbook
        /// </summary>
        /// <param name="srcColumn">the column to copy from</param>
        /// <param name="policy">the policy to determine what gets copied</param>
        public void CopyColumnFrom(IColumn srcColumn, CellCopyPolicy policy)
        {
            if (srcColumn == null)
            {
                // srcColumn is blank. Overwrite cells with blank values,
                // blank styles, etc per cell copy policy
                foreach (ICell destCell in this)
                {
                    XSSFCell srcCell = null;

                    ((XSSFCell)destCell).CopyCellFrom(srcCell, policy);
                }

                if (policy.IsCopyMergedRegions)
                {
                    // Remove MergedRegions in dest column
                    int destColumnNum = ColumnNum;
                    int index = 0;
                    HashSet<int> indices = new HashSet<int>();
                    foreach (CellRangeAddress destRegion in Sheet.MergedRegions)
                    {
                        if (destColumnNum == destRegion.FirstColumn
                            && destColumnNum == destRegion.LastColumn)
                        {
                            _ = indices.Add(index);
                        }

                        index++;
                    }

                    (Sheet as XSSFSheet).RemoveMergedRegions(indices.ToList());
                }

                if (policy.IsCopyColumnWidth)
                {
                    // clear row height
                    Width = -1;
                }
            }
            else
            {
                foreach (ICell c in srcColumn)
                {
                    XSSFCell srcCell = (XSSFCell)c;
                    XSSFCell destCell =
                        CreateCell(srcCell.RowIndex, srcCell.CellType) as XSSFCell;
                    destCell.CopyCellFrom(srcCell, policy);
                }

                XSSFColumnShifter columnShifter = new XSSFColumnShifter(_sheet);
                int sheetIndex = _sheet.Workbook.GetSheetIndex(_sheet);
                string sheetName = _sheet.Workbook.GetSheetName(sheetIndex);
                int srcColumnNum = srcColumn.ColumnNum;
                int destColumnNum = ColumnNum;
                int columnDifference = destColumnNum - srcColumnNum;
                FormulaShifter shifter = FormulaShifter.CreateForColumnCopy(
                    sheetIndex,
                    sheetName,
                    srcColumnNum,
                    srcColumnNum,
                    columnDifference,
                    SpreadsheetVersion.EXCEL2007);
                columnShifter.UpdateColumnFormulas(this, shifter);

                if (policy.IsCopyMergedRegions)
                {
                    foreach (CellRangeAddress srcRegion in srcColumn.Sheet.MergedRegions)
                    {
                        if (srcColumnNum == srcRegion.FirstColumn
                            && srcColumnNum == srcRegion.LastColumn)
                        {
                            CellRangeAddress destRegion = srcRegion.Copy();
                            destRegion.FirstColumn = destColumnNum;
                            destRegion.LastColumn = destColumnNum;
                            _ = Sheet.AddMergedRegion(destRegion);
                        }
                    }
                }

                if (policy.IsCopyColumnWidth)
                {
                    Width = srcColumn.Width;
                }
            }
        }

        public void MoveCell(ICell cell, int newColumn)
        {
            throw new NotImplementedException();
        }

        public IColumn CopyColumnTo(int targetIndex)
        {
            return ((XSSFSheet)Sheet).CopyColumn(ColumnNum, targetIndex);
        }

        public ICell CopyCell(int sourceIndex, int targetIndex)
        {
            return CellUtil.CopyCell(this, sourceIndex, targetIndex);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// Returns the underlying CT_Col xml node Containing all cell 
        /// defInitions in this column
        /// </summary>
        /// <returns>the underlying CT_Col xml node</returns>
        public CT_Col GetCTCol()
        {
            return _column;
        }
        #endregion

        #region Internal methods
        /// <summary>
        /// update cell references when Shifting columns
        /// </summary>
        /// <param name="n">n the number of columns to move</param>
        internal void Shift(int n)
        {
            int columnNum = ColumnNum + n;
            CalculationChain calcChain =
                ((XSSFWorkbook)_sheet.Workbook).GetCalculationChain();
            int sheetId = (int)_sheet.sheet.sheetId;
            string msg = "Column[columnNum=" + ColumnNum + "] contains cell(s) " +
                "included in a multi-cell array formula. You cannot change " +
                "part of an array.";

            foreach (ICell c in this)
            {
                XSSFCell cell = (XSSFCell)c;

                if (cell.IsPartOfArrayFormulaGroup)
                {
                    cell.NotifyArrayFormulaChanging(msg);
                }

                //remove the reference in the calculation chain
                calcChain?.RemoveItem(sheetId, cell.GetReference());

                CT_Cell ctCell = cell.GetCTCell();
                string r = new CellReference(cell.RowIndex, columnNum)
                    .FormatAsString();
                ctCell.r = r;
                cell.ColumnIndex = columnNum;
            }

            ColumnNum = columnNum;
        }

        /// <summary>
        /// Fired when the document is written to an output stream.
        /// See <see cref="XSSFSheet.Write"/>
        /// </summary>
        internal void OnDocumentWrite()
        {

        }
        #endregion

        #region IEnumerable and IComparable members
        /// <summary>
        /// Cell iterator over the physically defined cell
        /// </summary>
        /// <returns>an iterator over cells in this column.</returns>
        public List<ICell>.Enumerator CellIterator()
        {
            return Cells.GetEnumerator();
        }

        /// <summary>
        /// Alias for <see cref="CellIterator"/> to allow  foreach loops
        /// </summary>
        /// <returns>an iterator over cells in this row.</returns>
        public IEnumerator<ICell> GetEnumerator()
        {
            return CellIterator();
        }

        /// <summary>
        /// Compares two <see cref="XSSFColumn"/> objects. Two columns are 
        /// equal if they belong to the same worksheet and their column indexes
        /// are equal.
        /// </summary>
        /// <param name="other">the <see cref="XSSFColumn"/> to be compared.</param>
        /// <returns>
        /// the value 0 if the column number of this <see cref="XSSFColumn"/> 
        /// is equal to the column number of the argument <see cref="XSSFColumn"/> 
        /// a value less than 0 if the column number of this this 
        /// <see cref="XSSFColumn"/> is numerically less than the column number
        /// of the argument <see cref="XSSFColumn"/> a value greater than 0 if
        /// the column number of this this <see cref="XSSFColumn"/> is 
        /// numerically greater than the column number of the 
        /// argument <see cref="XSSFColumn"/>
        /// </returns>
        /// <exception cref="ArgumentException">if the argument column belongs 
        /// to a different worksheet</exception>
        public int CompareTo(XSSFColumn other)
        {
            return Sheet != other.Sheet
                ? throw new ArgumentException(
                    "The compared columns must belong to the same sheet")
                : ColumnNum.CompareTo(other.ColumnNum);
        }

        public override bool Equals(object obj)
        {
            return obj is XSSFColumn other
                && ColumnNum == other.ColumnNum
                && Sheet == other.Sheet;
        }

        public override int GetHashCode()
        {
            return _column.GetHashCode();
        }
        #endregion

        #region Private methods
        /// <summary>
        /// formatted xml representation of this column
        /// </summary>
        /// <returns>formatted xml representation of this column</returns>
        public override string ToString()
        {
            return _column.ToString();
        }

        private void ValidateCellType(CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Blank:
                case CellType.Numeric:
                case CellType.String:
                case CellType.Formula:
                case CellType.Boolean:
                case CellType.Error:
                    break;
                default:
                    throw new ArgumentException("Illegal cell type: " + cellType);
            }
        }

        /// <summary>
        /// Get the xssfcell representing a given column (logical cell)
        /// 0-based. If you ask for a cell that is not defined, then
        /// you Get a null.
        /// This is the basic call, with no policies applied
        /// </summary>
        /// <param name="cellnum">0 based row number</param>
        /// <returns>Cell representing that row or null if Undefined.</returns>
        private ICell RetrieveCell(int cellnum)
        {
            IRow row = Sheet.GetRow(cellnum);

            return row?.GetCell(ColumnNum);
        }

        private int GetFirstKey()
        {
            return Cells.Min(c => c.RowIndex);
        }

        private int GetLastKey()
        {
            return Cells.Max(c => c.RowIndex);
        }
        #endregion
    }
}
