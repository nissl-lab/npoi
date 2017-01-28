using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.POIFS.FileSystem;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaminging;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFRow : IRow, IComparable<SXSSFRow>
    {
        private static bool? UNDEFINED = null;

        public SXSSFSheet _sheet; // parent sheet
        //TODO: replacing with dict to get compling may need to alter for performance.
        public SortedDictionary<int,SXSSFCell> _cells = new SortedDictionary<int, SXSSFCell>();
        //public List<Tuple<int,SXSSFCell>> _cells = new List<Tuple<int, SXSSFCell>>();
        //private SortedMap<Integer, SXSSFCell> _cells = new TreeMap<Integer, SXSSFCell>();
        public short _style = -1; // index of cell style in style table
        public short _height = -1; // row height in twips (1/20 point)
        public bool _zHeight = false; // row zero-height (this is somehow different than being hidden)
        public int _outlineLevel = 0;   // Outlining level of the row, when outlining is on
                                         // use Boolean to have a tri-state for on/off/undefined 
        public bool? _hidden = UNDEFINED;
        public bool? _collapsed = UNDEFINED;

        public SXSSFRow(SXSSFSheet sheet)
        {
            _sheet = sheet;
        }

        public IEnumerator<ICell> allCellsIterator()
        {
            return new CellIterator(LastCellNum, null);
        }
        public bool hasCustomHeight()
        {
            return _height != -1;
        }
        public List<ICell> Cells
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public short FirstCellNum
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public short Height
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

        public float HeightInPoints
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
                //TODO:should make this work with dictionary
                return _cells.Count == 0 ? (short) -1 : Convert.ToInt16(_cells.Last().Key + 1);
                
            }
        }

        public int OutlineLevel
        {
            get { return _outlineLevel; }
        }

        public int PhysicalNumberOfCells
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public int RowNum
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

        public ICellStyle RowStyle
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

        public ISheet Sheet
        {
            get { return _sheet; }
        }

        public bool ZeroHeight
        {
            get { return _zHeight; }

            set
            {
                throw new NotImplementedException();
            }
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
                //TODO: v shoulod be v.name(); as in the name of the enum
                throw new InvalidOperationException("Invalid column index (" + cellIndex
                        + ").  Allowable column range for " + v + " is (0.."
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

            //TODO come bakc and do comparision with if statement, may need to deep 
            switch (policy._policy)
            {
                //case MissingCellPolicy.RETURN_NULL_AND_BLANK:
                case MissingCellPolicy.Policy.RETURN_NULL_AND_BLANK:
                    return cell;
                //case MissingCellPolicy.RETURN_BLANK_AS_NULL:
                case MissingCellPolicy.Policy.RETURN_BLANK_AS_NULL:
                    bool isBlank = (cell != null && cell.CellType == CellType.Blank);
                    return (isBlank) ? null : cell;
                // case MissingCellPolicy.CREATE_NULL_AS_BLANK:
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
                //TODO: this seems slopy
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
                //TODO: in my case, i can assume no skipped cells
                pos += 1;
                return _cells.ContainsKey(pos);
            }

            public void Reset()
            {
                throw new NotImplementedException();
            }
        }

        //TODO: enumerator or enumberabl?
        public class CellIterator : IEnumerator<ICell>
        {
            private Dictionary<int, SXSSFCell> _cells;
            private int maxColumn;
            private int pos;
            public CellIterator(int lastCellNum, Dictionary<int, SXSSFCell> cells)
            {
                //TODO: Should be static so I don't know if i can just pass this to here
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

            //TODO THese do need to get implemented and translated 
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


