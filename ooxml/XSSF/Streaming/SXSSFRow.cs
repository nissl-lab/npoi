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
using System.Linq;
using NPOI.SS;
using NPOI.SS.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFRow : IRow, IComparable<SXSSFRow>
    {
        private SXSSFSheet _sheet; // parent sheet
        private SortedDictionary<int, SXSSFCell> _cells = new SortedDictionary<int, SXSSFCell>();
        private short _style = -1; // index of cell style in style table
        private bool _zHeight; // row zero-height (this is somehow different than being hidden)
        private float _height = -1;


        // use Boolean to have a tri-state for on/off/undefined 
        public bool? Hidden { get; set; }
        public bool? Collapsed { get; set; }

        public SXSSFRow(SXSSFSheet sheet)
        {
            _sheet = sheet;
        }

        public IEnumerator<ICell> allCellsIterator()
        {
            return new CellIterator(LastCellNum, null);
        }
        public bool HasCustomHeight()
        {
            return Height != -1;
        }

        public List<ICell> Cells
        {
            get { return _cells.Values.Select(cell => (ICell) cell).ToList(); }
        }

        public short FirstCellNum
        {
            get
            {
                try
                {
                    return (short) _cells.Keys.First();
                }
                catch
                {
                    return -1;

                }
            }
        }

        public short Height
        {
            get
            {
                return (short)(_height == -1 ? Sheet.DefaultRowHeightInPoints * 20 : _height);
            }
            set { _height = value; }
        }

        public float HeightInPoints
        {
            get
            {
                return (float)(_height == -1 ? Sheet.DefaultRowHeightInPoints : _height / 20.0);
            }

            set
            {
                _height = (value == -1) ? -1 : (value*20);
            }
        }

        public bool IsFormatted
        {
            get
            {
                return _style > -1;
            }
        }

        public short LastCellNum
        {
            get
            {
                return _cells.Count == 0 ? (short)-1 : Convert.ToInt16(_cells.Last().Key + 1);

            }
        }

        public int OutlineLevel { get; set; }

        public int PhysicalNumberOfCells
        {
            get { return Cells.Count; }
        }

        public int RowNum
        {
            get { throw new NotImplementedException();/*return Sheet.GetRow(this);*/ }

            set
            {
                throw new NotImplementedException();

            }
        }

        public ICellStyle RowStyle
        {
            get
            {
                if (!IsFormatted) return null;

                return Sheet.Workbook.GetCellStyleAt(_style);
            }

            set
            {
                if (value == null)
                {
                    _style = -1;
                }
                else
                {
                    _style = value.Index;
                }
            }
        }

        public ISheet Sheet
        {
            get { return _sheet; }
        }

        public bool ZeroHeight
        {
            get { return _zHeight; }

            set { _zHeight = value; }
        }

        public int CompareTo(SXSSFRow other)
        {
            if (this.Sheet != other.Sheet)
            {
                throw new InvalidOperationException("The compared rows must belong to the same sheet");
            }

            var thisRow = this.RowNum;
            var otherRow = other.RowNum;
            return thisRow.CompareTo(otherRow);
        }

        public ICell CopyCell(int sourceIndex, int targetIndex)
        {
            throw new NotImplementedException();
        }

        public IRow CopyRowTo(int targetIndex)
        {
            throw new NotImplementedException();
        }

        public ICell CreateCell(int column)
        {
            return CreateCell(column, CellType.Blank);
        }

        public ICell CreateCell(int column, CellType type)
        {
            CheckBounds(column);
            SXSSFCell cell = new SXSSFCell(this, type);
            _cells.Add(column, cell);
            return cell;
        }


        /// <summary>
        /// throws RuntimeException if the bounds are exceeded.
        /// </summary>
        /// <param name="cellIndex"></param>
        private static void CheckBounds(int cellIndex)
        {
            SpreadsheetVersion v = SpreadsheetVersion.EXCEL2007;
            int maxcol = SpreadsheetVersion.EXCEL2007.LastColumnIndex;
            if (cellIndex < 0 || cellIndex > maxcol)
            {
                throw new InvalidOperationException("Invalid column index (" + cellIndex
                        + ").  Allowable column range for " + v.DefaultExtension + " is (0.."
                        + maxcol + ") or ('A'..'" + v.LastColumnName + "')");
            }
        }

        public ICell GetCell(int cellnum)
        {
            var policy = _sheet.Workbook.MissingCellPolicy;
            return GetCell(cellnum, policy);
        }

        public ICell GetCell(int cellnum, MissingCellPolicy policy)
        {
            CheckBounds(cellnum);

            var cell = _cells[cellnum];

            switch (policy._policy)
            {
                case MissingCellPolicy.Policy.RETURN_NULL_AND_BLANK:
                    return cell;
                case MissingCellPolicy.Policy.RETURN_BLANK_AS_NULL:
                    bool isBlank = (cell != null && cell.CellType == CellType.Blank);
                    return (isBlank) ? null : cell;
                case MissingCellPolicy.Policy.CREATE_NULL_AS_BLANK:
                    return (cell == null) ? CreateCell(cellnum, CellType.Blank) : cell;
                default:
                    throw new InvalidOperationException("Illegal policy " + policy + " (" + policy.id + ")");
            }
        }

        public IEnumerator<ICell> GetEnumerator()
        {
            return new FilledCellIterator(_cells);
        }

        public void MoveCell(ICell cell, int newColumn)
        {
            throw new NotImplementedException();
        }

        public void RemoveCell(ICell cell)
        {
            int index = getCellIndex((SXSSFCell)cell);
            _cells.Remove(index);
        }
        /**
 * Return the column number of a cell if it is in this row
 * Otherwise return -1
 *
 * @param cell the cell to get the index of
 * @return cell column index if it is in this row, -1 otherwise
 */
        /*package*/
        public int getCellIndex(SXSSFCell cell)
        {
            foreach (var entry in _cells)
            {
                if (entry.Value == cell)
                {
                    return entry.Key;
                }
            }
            return -1;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }

        /**
* Create an iterator over the cells from [0, getLastCellNum()).
* Includes blank cells, excludes empty cells
* 
* Returns an iterator over all filled cells (created via Row.createCell())
* Throws ConcurrentModificationException if cells are added, moved, or
* removed after the iterator is created.
*/
        public class FilledCellIterator : IEnumerator<ICell>
        {
            private SortedDictionary<int, SXSSFCell> _cells;
            private int pos = -1;
            public FilledCellIterator(SortedDictionary<int, SXSSFCell> cells)
            {
                _cells = cells;
            }

            public ICell Current
            {
                get { return _cells[pos]; }
            }

            object IEnumerator.Current
            {
                get
                {
                    throw new NotImplementedException();
                }
            }

            public void Dispose()
            {
                throw new NotImplementedException();
            }

            public IEnumerator<ICell> GetEnumerator()
            {
                return _cells.Values.GetEnumerator();
            }

            public bool MoveNext()
            {
                pos += 1;
                return _cells.ContainsKey(pos);
            }

            public void Reset()
            {
                throw new NotImplementedException();
            }
        }

        public class CellIterator : IEnumerator<ICell>
        {
            private Dictionary<int, SXSSFCell> _cells;
            private int maxColumn;
            private int pos;
            public CellIterator(int lastCellNum, Dictionary<int, SXSSFCell> cells)
            {
                maxColumn = lastCellNum; //last column PLUS ONE, SHOULD BE DERIVED from cells enum.
                pos = 0;
                _cells = cells;
            }


            public ICell Current
            {
                get
                {
                    throw new NotImplementedException();
                }
            }

            object IEnumerator.Current
            {
                get
                {
                    throw new NotImplementedException();
                }
            }

            public void Dispose()
            {
                throw new NotImplementedException();
            }

            public IEnumerator<ICell> GetEnumerator()
            {
                throw new NotImplementedException();
            }

            public bool hasNext()
            {
                return pos < maxColumn;
            }

            public bool MoveNext()
            {
                throw new NotImplementedException();
            }

            public ICell next()
            {
                if (hasNext())
                    return _cells[pos++];
                else
                    throw new NullReferenceException();
            }

            public void remove()
            {
                throw new InvalidOperationException();
            }

            public void Reset()
            {
                throw new NotImplementedException();
            }


        }
    }



}


