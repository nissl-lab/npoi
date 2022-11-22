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

    /// <summary>
    /// High level representation of a row of a spreadsheet.
    /// </summary>
    public class XSSFRow : IRow, IComparable<XSSFRow>
    {
        #region Private properties
        private static readonly POILogger _logger = POILogFactory.GetLogger(typeof(XSSFRow));

        /// <summary>
        /// the xml node Containing all cell defInitions for this row
        /// </summary>
        private readonly CT_Row _row;

        /// <summary>
        /// Cells of this row keyed by their column indexes.
        /// The SortedDictionary ensures that the cells are ordered by columnIndex in the ascending order.
        /// </summary>
        private readonly SortedDictionary<int, ICell> _cells;

        /// <summary>
        /// the parent sheet
        /// </summary>
        private readonly XSSFSheet _sheet;

        private readonly StylesTable _stylesSource;
        #endregion

        #region Public properties
        /// <summary>
        /// XSSFSheet this row belongs to
        /// </summary>
        public ISheet Sheet
        {
            get
            {
                return _sheet;
            }
        }

        /// <summary>
        /// Get the number of the first cell Contained in this row.
        /// </summary>
        /// <returns>short representing the first logical cell in the row,
        /// or -1 if the row does not contain any cells.</returns>
        public short FirstCellNum
        {
            get
            {
                return (short)(_cells.Count == 0 ? -1 : GetFirstKey());
            }
        }

        /// <summary>
        /// Gets the index of the last cell Contained in this row <b>PLUS ONE</b>. The result also
        /// happens to be the 1-based column number of the last cell. This value can be used as a
        /// standard upper bound when iterating over cells:
        /// </summary>
        /// <returns>short representing the last logical cell in the row <b>PLUS ONE</b>,
        /// or -1 if the row does not contain any cells.</returns>
        public short LastCellNum
        {
            get
            {
                return (short)(_cells.Count == 0 ? -1 : (GetLastKey() + 1));
            }
        }

        /// <summary>
        /// Get the row's height measured in twips (1/20th of a point). 
        /// If the height is not Set, the default worksheet value is returned,
        /// See <see cref="XSSFSheet.DefaultRowHeight"/>
        /// </summary>
        /// <returns>row height measured in twips (1/20th of a point)</returns>
        public short Height
        {
            get
            {
                return (short)(HeightInPoints * 20);
            }

            set
            {
                if (value < 0)
                {
                    if (_row.IsSetHt())
                    {
                        _row.UnsetHt();
                    }

                    if (_row.IsSetCustomHeight())
                    {
                        _row.UnsetCustomHeight();
                    }
                }
                else
                {
                    _row.ht = (double)value / 20;
                    _row.customHeight = true;

                }
            }
        }

        /// <summary>
        /// Returns row height measured in point size. If the height is not Set, 
        /// the default worksheet value is returned,See <see cref="XSSFSheet.DefaultRowHeightInPoints"/>
        /// </summary>
        public float HeightInPoints
        {
            get
            {
                if (_row.IsSetHt())
                {
                    return (float)_row.ht;
                }

                return _sheet.DefaultRowHeightInPoints;
            }
            set
            {
                Height = (short)(value == -1 ? -1 : (value * 20));
            }
        }

        /// <summary>
        /// Gets the number of defined cells (NOT number of cells in the actual row!).
        /// That is to say if only columns 0,4,5 have values then there would be 3.
        /// </summary>
        /// <returns>int representing the number of defined cells in the row.</returns>
        public int PhysicalNumberOfCells
        {
            get
            {
                return _cells.Count;
            }
        }

        /// <summary>
        /// Get row number this row represents
        /// </summary>
        /// <returns>the row number (0 based)</returns>
        public int RowNum
        {
            get
            {
                return (int)_row.r - 1;
            }

            set
            {
                int maxrow = SpreadsheetVersion.EXCEL2007.LastRowIndex;
                if (value < 0 || value > maxrow)
                {
                    throw new ArgumentException("Invalid row number (" + value
                            + ") outside allowable range (0.." + maxrow + ")");
                }

                _row.r = (uint)(value + 1);
            }
        }

        /// <summary>
        /// Get whether or not to display this row with 0 height
        /// </summary>
        public bool ZeroHeight
        {
            get
            {
                return _row.hidden;
            }

            set
            {
                _row.hidden = value;
            }
        }

        /// <summary>
        /// Is this row formatted? Most aren't, but some rows
        /// do have whole-row styles. For those that do, you
        /// can get the formatting from <see cref="RowStyle"/>
        /// </summary>
        public bool IsFormatted
        {
            get
            {
                return _row.IsSetS();
            }
        }

        /// <summary>
        /// Returns the whole-row cell style. Most rows won't
        /// have one of these, so will return null. Call
        /// <see cref="IsFormatted"/> to check first.
        /// </summary>
        public ICellStyle RowStyle
        {
            get
            {
                if (IsFormatted && _stylesSource != null
                    && _stylesSource.NumCellStyles > 0)
                {
                    return _stylesSource.GetStyleAt((int)_row.s);
                }

                    return null;
                }

            set
            {
                if (value == null)
                {
                    if (_row.IsSetS())
                    {
                        _row.UnsetS();
                        _row.UnsetCustomFormat();
                    }
                }
                else
                {
                    XSSFCellStyle xStyle = (XSSFCellStyle)value;
                    xStyle.VerifyBelongsToStylesSource(_stylesSource);

                    long idx = _stylesSource.PutStyle(xStyle);
                    _row.s = (uint)idx;
                    _row.customFormat = true;
                }

                foreach (var cell in _cells.Values)
                {
                    cell.CellStyle = value;
                }
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Construct an XSSFRow.
        /// </summary>
        /// <param name="row">the xml node Containing all cell defInitions for this row.</param>
        /// <param name="sheet">the parent sheet.</param>
        public XSSFRow(CT_Row row, XSSFSheet sheet)
        {
            _row = row;
            _sheet = sheet;
            _cells = new SortedDictionary<int, ICell>();
            if (0 < row.SizeOfCArray())
            {
                foreach (CT_Cell c in row.c)
                {
                    XSSFCell cell = new XSSFCell(this, c);
                    _cells.Add(cell.ColumnIndex, cell);
                    sheet.OnReadCell(cell);
                }
            }

            if (!row.IsSetR())
            {
                // Certain file format writers skip the row number
                // Assume no gaps, and give this the next row number
                int nextRowNum = sheet.LastRowNum + 2;
                if (nextRowNum == 2 && sheet.PhysicalNumberOfRows == 0)
                {
                    nextRowNum = 1;
                }

                row.r = (uint)nextRowNum;
            }

            _stylesSource = ((XSSFWorkbook)sheet.Workbook).GetStylesSource();
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Use this to create new cells within the row and return it.
        /// The cell that is returned is a <see cref="CellType.Blank"/>. The type can be Changed
        /// either through calling <see cref="ICell.SetCellValue"/> or <see cref="ICell.SetCellType"/>.
        /// </summary>
        /// <param name="columnIndex">the column number this cell represents</param>
        /// <returns>a high level representation of the Created cell</returns>
        /// <exception cref="ArgumentException">if columnIndex is less than 0 or greater than 16384, 
        /// the maximum number of columns supported by the SpreadsheetML format(.xlsx)</exception>
        public ICell CreateCell(int columnIndex)
        {
            return CreateCell(columnIndex, CellType.Blank);
        }

        /// <summary>
        /// Use this to create new cells within the row and return it.
        /// </summary>
        /// <param name="columnIndex">the column number this cell represents</param>
        /// <param name="type">the cell's data type</param>
        /// <returns>a high level representation of the Created cell.</returns>
        /// <exception cref="ArgumentException">if columnIndex is less than 0 or greater than 16384, 
        /// the maximum number of columns supported by the SpreadsheetML format(.xlsx)</exception>
        public ICell CreateCell(int columnIndex, CellType type)
        {
            CT_Cell ctCell;
            XSSFCell prev = _cells.ContainsKey(columnIndex) ? (XSSFCell)_cells[columnIndex] : null;
            if (prev != null)
            {
                ctCell = prev.GetCTCell();
                ctCell.Set(new CT_Cell());
            }
            else
            {
                ctCell = _row.AddNewC();
            }

            XSSFCell xcell = new XSSFCell(this, ctCell);
            xcell.SetCellNum(columnIndex);
            if (type != CellType.Blank)
            {
                xcell.SetCellType(type);
            }

            if (IsFormatted)
            {
                xcell.CellStyle = RowStyle;
            }

            _cells[columnIndex] = xcell;
            return xcell;
        }

        /// <summary>
        /// Returns the cell at the given (0 based) index,
        /// with the <see cref="MissingCellPolicy"/> from the parent Workbook.
        /// </summary>
        /// <param name="cellnum"></param>
        /// <returns>the cell at the given (0 based) index</returns>
        public ICell GetCell(int cellnum)
        {
            return GetCell(cellnum, _sheet.Workbook.MissingCellPolicy);
        }

        /// <summary>
        /// Returns the cell at the given (0 based) index, with the specified <see cref="MissingCellPolicy"/>
        /// </summary>
        /// <param name="cellnum"></param>
        /// <param name="policy"></param>
        /// <returns>the cell at the given (0 based) index</returns>
        /// <exception cref="ArgumentException">if cellnum is less than 0 or the specified MissingCellPolicy is invalid</exception>
        public ICell GetCell(int cellnum, MissingCellPolicy policy)
        {
            if (cellnum < 0)
            {
                throw new ArgumentException("Cell index must be >= 0");
            }

            XSSFCell cell = (XSSFCell)RetrieveCell(cellnum);
            switch (policy)
            {
                case MissingCellPolicy.RETURN_NULL_AND_BLANK:
                    return cell;
                case MissingCellPolicy.RETURN_BLANK_AS_NULL:
                    bool isBlank = cell != null && cell.CellType == CellType.Blank;
                    return isBlank ? null : cell;
                case MissingCellPolicy.CREATE_NULL_AS_BLANK:
                    return cell ?? CreateCell(cellnum, CellType.Blank);
                default:
                    throw new ArgumentException("Illegal policy " + policy + " (" + policy + ")");
            }
        }

        /// <summary>
        /// Remove the Cell from this row.
        /// </summary>
        /// <param name="cell">the cell to remove</param>
        /// <exception cref="ArgumentException"></exception>
        public void RemoveCell(ICell cell)
        {
            if (cell.Row != this)
            {
                throw new ArgumentException("Specified cell does not belong to this row");
            }

            XSSFCell xcell = (XSSFCell)cell;
            if (xcell.IsPartOfArrayFormulaGroup)
            {
                xcell.NotifyArrayFormulaChanging();
            }

            if (cell.CellType == CellType.Formula)
            {
                ((XSSFWorkbook)_sheet.Workbook).OnDeleteFormula(xcell);
            }

            _cells.Remove(cell.ColumnIndex);
        }

        /// <summary>
        /// Copy the cells from srcRow to this row
        /// If this row is not a blank row, this will merge the two rows, overwriting
        /// the cells in this row with the cells in srcRow
        /// If srcRow is null, overwrite cells in destination row with blank values, styles, etc per cell copy policy
        /// srcRow may be from a different sheet in the same workbook
        /// </summary>
        /// <param name="srcRow">the rows to copy from</param>
        /// <param name="policy">policy the policy to determine what gets copied</param>
        public void CopyRowFrom(IRow srcRow, CellCopyPolicy policy)
        {
            if (srcRow == null)
            {
                // srcRow is blank. Overwrite cells with blank values, blank styles, etc per cell copy policy
                foreach (ICell destCell in this)
                {
                    XSSFCell srcCell = null;
                    // FIXME: remove type casting when copyCellFrom(Cell, CellCopyPolicy) is added to Cell interface
                    ((XSSFCell)destCell).CopyCellFrom(srcCell, policy);
                }

                if (policy.IsCopyMergedRegions)
                {
                    // Remove MergedRegions in dest row
                    int destRowNum = RowNum;
                    int index = 0;
                    HashSet<int> indices = new HashSet<int>();
                    foreach (CellRangeAddress destRegion in Sheet.MergedRegions)
                    {
                        if (destRowNum == destRegion.FirstRow && destRowNum == destRegion.LastRow)
                        {
                            indices.Add(index);
                        }

                        index++;
                    }

                    (Sheet as XSSFSheet).RemoveMergedRegions(indices.ToList());
                }

                if (policy.IsCopyRowHeight)
                {
                    // clear row height
                    Height = -1;
                }
            }
            else
            {
                foreach (ICell c in srcRow)
                {
                    XSSFCell srcCell = (XSSFCell)c;
                    XSSFCell destCell = CreateCell(srcCell.ColumnIndex, srcCell.CellType) as XSSFCell;
                    destCell.CopyCellFrom(srcCell, policy);
                }

                XSSFRowShifter rowShifter = new XSSFRowShifter(_sheet);
                int sheetIndex = _sheet.Workbook.GetSheetIndex(_sheet);
                string sheetName = _sheet.Workbook.GetSheetName(sheetIndex);
                int srcRowNum = srcRow.RowNum;
                int destRowNum = RowNum;
                int rowDifference = destRowNum - srcRowNum;
                FormulaShifter shifter = FormulaShifter.CreateForRowCopy(sheetIndex, sheetName, srcRowNum, srcRowNum, rowDifference, SpreadsheetVersion.EXCEL2007);
                rowShifter.UpdateRowFormulas(this, shifter);
                // Copy merged regions that are fully contained on the row
                // FIXME: is this something that rowShifter could be doing?
                if (policy.IsCopyMergedRegions)
                {
                    foreach (CellRangeAddress srcRegion in srcRow.Sheet.MergedRegions)
                    {
                        if (srcRowNum == srcRegion.FirstRow && srcRowNum == srcRegion.LastRow)
                        {
                            CellRangeAddress destRegion = srcRegion.Copy();
                            destRegion.FirstRow = destRowNum;
                            destRegion.LastRow = destRowNum;
                            Sheet.AddMergedRegion(destRegion);
                        }
                    }
                }

                if (policy.IsCopyRowHeight)
                {
                    Height = srcRow.Height;
                }
            }
        }

        /// <summary>
        /// Applies a whole-row cell styling to the row.
        /// If the value is null then the style information is Removed,
        /// causing the cell to used the default workbook style.
        /// </summary>
        /// <param name="style"></param>
        public void SetRowStyle(ICellStyle style)
        {

        }

        /// <summary>
        /// Returns the underlying CT_Row xml node Containing all cell defInitions in this row
        /// </summary>
        /// <returns>the underlying CT_Row xml node</returns>
        public CT_Row GetCTRow()
        {
            return _row;
        }

        /// <summary>
        /// Fired when the document is written to an output stream.
        /// See <see cref="XSSFSheet.Write"/>
        /// </summary>
        internal void OnDocumentWrite()
        {
            // check if cells in the CT_Row are ordered
            bool isOrdered = true;
            if (_row.SizeOfCArray() != _cells.Count)
            {
                isOrdered = false;
            }
            else
            {
                int i = 0;
                foreach (XSSFCell cell in _cells.Values.Cast<XSSFCell>())
                {
                    CT_Cell c1 = cell.GetCTCell();
                    CT_Cell c2 = _row.GetCArray(i++);

                    string r1 = c1.r;
                    string r2 = c2.r;
                    if (!(r1 == null ? r2 == null : r1.Equals(r2)))
                    {
                        isOrdered = false;
                        break;
                    }
                }
            }

            if (!isOrdered)
            {
                CT_Cell[] cArray = new CT_Cell[_cells.Count];
                int i = 0;
                foreach (XSSFCell c in _cells.Values.Cast<XSSFCell>())
                {
                    cArray[i++] = c.GetCTCell();
                }

                _row.SetCArray(cArray);
            }
        }
        #endregion

        #region Internal methods
        /// <summary>
        /// update cell references when Shifting rows
        /// </summary>
        /// <param name="n">n the number of rows to move</param>
        internal void Shift(int n)
        {
            int rownum = RowNum + n;
            Model.CalculationChain calcChain = ((XSSFWorkbook)_sheet.Workbook).GetCalculationChain();
            int sheetId = (int)_sheet.sheet.sheetId;
            string msg = "Row[rownum=" + RowNum + "] contains cell(s) included in a multi-cell array formula. " +
                    "You cannot change part of an array.";
            foreach (ICell c in this)
            {
                XSSFCell cell = (XSSFCell)c;
                if (cell.IsPartOfArrayFormulaGroup)
                {
                    cell.NotifyArrayFormulaChanging(msg);
                }

                //remove the reference in the calculation chain
                if (calcChain != null)
                {
                    calcChain.RemoveItem(sheetId, cell.GetReference());
                }

                CT_Cell CT_Cell = cell.GetCTCell();
                string r = new CellReference(rownum, cell.ColumnIndex).FormatAsString();
                CT_Cell.r = r;
            }

            RowNum = rownum;
        }

        internal void RebuildCells()
        {
            Dictionary<int, XSSFCell> map = new Dictionary<int, XSSFCell>();
            foreach (XSSFCell c in _cells.Values)
            {
                map.Add(c.ColumnIndex, c);
            }

            _cells.Clear();

            foreach (KeyValuePair<int, XSSFCell> kv in map)
            {
                _cells.Add(kv.Key, kv.Value);
            }

            // Sort CT_Cols by index asc.
            _row.c.Sort((col1, col2) => col1.r.CompareTo(col2.r));
        }
        #endregion

        #region IEnumerable and IComparable members
        /// <summary>
        /// Cell iterator over the physically defined cell
        /// </summary>
        /// <returns>an iterator over cells in this row.</returns>
        public SortedDictionary<int, ICell>.ValueCollection.Enumerator CellIterator()
        {
            return _cells.Values.GetEnumerator();
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
        /// Compares two <see cref="XSSFRow"/> objects. Two rows are equal if they belong to the 
        /// same worksheet and their row indexes are equal.
        /// </summary>
        /// <param name="other">the <see cref="XSSFRow"/> to be compared.</param>
        /// <returns>
        /// the value 0 if the row number of this <see cref="XSSFRow"/> is
        /// equal to the row number of the argument <see cref="XSSFRow"/>
        /// a value less than 0 if the row number of this this <see cref="XSSFRow"/> is
        /// numerically less than the row number of the argument <see cref="XSSFRow"/>
        /// a value greater than 0 if the row number of this this <see cref="XSSFRow"/> is
        /// numerically greater than the row number of the argument <see cref="XSSFRow"/>
        /// </returns>
        /// <exception cref="ArgumentException">if the argument row belongs to a different worksheet</exception>
        public int CompareTo(XSSFRow other)
        {
            if (Sheet != other.Sheet)
            {
                throw new ArgumentException("The compared rows must belong to the same sheet");
            }

            return RowNum.CompareTo(other.RowNum);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is XSSFRow))
            {
                return false;
            }

            XSSFRow other = (XSSFRow)obj;

            return (RowNum == other.RowNum) &&
                   (Sheet == other.Sheet);
        }

        public override int GetHashCode()
        {
            return _row.GetHashCode();
        }
        #endregion

        #region IRow Members
        public List<ICell> Cells
        {
            get
            {
                List<ICell> cells = new List<ICell>();
                foreach (ICell cell in _cells.Values)
                {
                    cells.Add(cell);
                }

                return cells;
            }
        }

        public void MoveCell(ICell cell, int newColumn)
        {
            throw new NotImplementedException();
        }

        public IRow CopyRowTo(int targetIndex)
        {
            return Sheet.CopyRow(RowNum, targetIndex);
        }

        public ICell CopyCell(int sourceIndex, int targetIndex)
        {
            return CellUtil.CopyCell(this, sourceIndex, targetIndex);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool HasCustomHeight()
        {
            throw new NotImplementedException();
        }

        public int OutlineLevel
        {
            get
            {
                return _row.outlineLevel;
            }

            set
            {
                _row.outlineLevel = (byte)value;
            }
        }

        public bool? Hidden
        {
            get
            {
                return _row.hidden;
            }

            set
            {
                _row.hidden = value ?? false;
            }
        }

        public bool? Collapsed
        {
            get
            {
                return _row.collapsed;
            }

            set
            {
                _row.collapsed = value ?? false;
            }
        }
        #endregion

        #region Private methods
        /// <summary>
        /// formatted xml representation of this row
        /// </summary>
        /// <returns>formatted xml representation of this row</returns>
        public override string ToString()
        {
            return _row.ToString();
        }

        /// <summary>
        /// Get the hssfcell representing a given column (logical cell)
        /// 0-based. If you ask for a cell that is not defined, then
        /// you Get a null.
        /// This is the basic call, with no policies applied
        /// </summary>
        /// <param name="cellnum">0 based column number</param>
        /// <returns>Cell representing that column or null if Undefined.</returns>
        private ICell RetrieveCell(int cellnum)
        {
            if (!_cells.ContainsKey(cellnum))
            {
                return null;
            }
            //if (cellnum < 0 || cellnum >= cells.Count) return null;
            return _cells[cellnum];
        }

        private int GetFirstKey()
        {
            return _cells.Keys.Min();
        }

        private int GetLastKey()
        {
            return _cells.Keys.Max();
        }
        #endregion
    }
}