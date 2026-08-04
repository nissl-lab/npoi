using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.UserModel
{
    public class NCellRange: IEnumerable<ICell>
    {
        private ISheet _sheet;
        private CellRangeAddressList _ranges;

        public CellRangeAddressList Ranges => _ranges;

        public int Width => BoundingBox.LastColumn - BoundingBox.FirstColumn + 1;

        public int Height => BoundingBox.LastRow - BoundingBox.FirstRow + 1;

        public int Size => _ranges.NumberOfCells();

        public string Address => string.Join(",", _ranges.CellRangeAddresses.Select(r => r.FormatAsString()));

        public ICell TopLeftCell => GetTopLeftCell();

        public ISheet Sheet => _sheet;

        private CellRangeAddress BoundingBox
        {
            get
            {
                var ranges = _ranges.CellRangeAddresses;
                int firstRow = ranges.Min(r => r.FirstRow);
                int lastRow = ranges.Max(r => r.LastRow);
                int firstColumn = ranges.Min(r => r.FirstColumn);
                int lastColumn = ranges.Max(r => r.LastColumn);
                return new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
            }
        }

        public double Sum(Func<ICell, double> selector)
        {
            return Cells.Sum(selector);
        }

        public double Min(Func<ICell, double> selector)
        {
            return Cells.Min(selector);
        }

        public double Max(Func<ICell, double> selector)
        {
            return Cells.Max(selector);
        }

        public double Average(Func<ICell, double> selector)
        {
            return Cells.Average(selector);
        }

        public T Max<T>(Func<ICell, T> selector) where T : IComparable<T>
        {
            return Cells.Max(selector);
        }

        public T Min<T>(Func<ICell, T> selector) where T : IComparable<T>
        {
            return Cells.Min(selector);
        }

        /// <summary>
        /// Set this cell range as active
        /// </summary>
        /// <returns></returns>
        public NCellRange SetActive()
        {
            var address = _ranges.GetCellRangeAddress(0);
            _sheet.SetActiveCellRange(address.FirstRow, address.LastRow, address.FirstColumn, address.LastColumn);
            return this;
        }

        public string Formula { get => throw new NotImplementedException(); set => this.SetCellFormula(value); }

        public NCellRange SetCellStyle(ICellStyle style, bool createMissingRowAndCol)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.CellStyle = style);
            return this;
        }

        public void SetCellComment(IComment comment) {
            if(comment==null)
            {
                RemoveCellComment();
                return;
            }
            foreach(var address in _ranges.CellRangeAddresses)
            {
                for(int i = address.FirstRow; i<=address.LastRow; i++)
                {
                    for(int j = address.FirstColumn; j<=address.LastColumn; j++)
                    {
                        comment.SetAddress(i, j);
                    }
                }
            }
        }
        public NCellRange SetHyperlink(IHyperlink hyperlink, bool createMissingRowAndCol=true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.Hyperlink = hyperlink);
            return this;
        }

        public NCellRange(ISheet sheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            _sheet = sheet;
            validateRowCol(fromRow, fromCol);
            validateRowCol(toRow, toCol);
            _ranges = new CellRangeAddressList(fromRow, toRow, fromCol, toCol);
        }
        private void validateRowCol(int row, int col)
        {
            if(row<0||row>_sheet.Workbook.SpreadsheetVersion.MaxRows)
                throw new ArgumentException($"row index {row} is out of range");
            if(col<0||col>_sheet.Workbook.SpreadsheetVersion.MaxColumns)
                throw new ArgumentException($"column index {col} is out of range");
        }

        public ICell GetCell(int rowInRange, int colInRange)
        {
            if(_ranges.CountRanges()>1)
                throw new NotSupportedException("GetCell is not supported for a range with multiple areas");
            var address = _ranges.GetCellRangeAddress(0);
            var row = _sheet.GetRow(address.FirstRow+rowInRange);
            if(row==null)
                row = _sheet.CreateRow(address.FirstRow+rowInRange);
            return row.GetCell(address.FirstColumn+colInRange, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }

        private ICell GetTopLeftCell()
        {
            var address = _ranges.GetCellRangeAddress(0);
            var row = _sheet.GetRow(address.FirstRow);
            if(row==null)
                row = _sheet.CreateRow(address.FirstRow);
            return row.GetCell(address.FirstColumn, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }

        private void ForEachCell(bool createMissingRowAndCol, MissingCellPolicy missingPolicy, Action<ICell> action)
        {
            foreach(var address in _ranges.CellRangeAddresses)
            {
                for(int i = address.FirstRow; i<=address.LastRow; i++)
                {
                    var row = _sheet.GetRow(i);
                    if(row==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            row = _sheet.CreateRow(i);
                    }
                    for(int j = address.FirstColumn; j<=address.LastColumn; j++)
                    {
                        var cell = row.GetCell(j, missingPolicy);
                        if(cell==null)
                        {
                            if(!createMissingRowAndCol)
                                continue;
                            else
                                cell = row.CreateCell(j);
                        }
                        action(cell);
                    }
                }
            }
        }

        public List<ICell> Cells
        {
            get {
                List<ICell> cells = new List<ICell>();
                var seen = new HashSet<(int row, int column)>();
                foreach(var address in _ranges.CellRangeAddresses)
                {
                    for(int i = address.FirstRow; i<=((address.LastRow<_sheet.LastRowNum)?address.LastRow: _sheet.LastRowNum); i++)
                    {
                        var row = _sheet.GetRow(i);
                        if(row==null)
                            continue;
                        for(int j = address.FirstColumn; j<=((address.LastColumn<row.LastCellNum)?address.LastColumn:row.LastCellNum); j++)
                        {
                            if(!seen.Add((i, j)))
                                continue;
                            var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                            if(cell!=null)
                                cells.Add(cell);
                        }
                    }
                }
                return cells;
            }
        }

        public IEnumerator<ICell> GetEnumerator()
        {
            return Cells.GetEnumerator();
        }

        public NCellRange SetCellType(CellType cellType, bool createMissingRowAndCol= true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetCellType(cellType));
            return this;
        }

        public NCellRange SetBlank(bool createMissingRowAndCol = false)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetBlank());
            return this;
        }

        public NCellRange SetCellValue(double value, bool createMissingRowAndCol = true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetCellValue(value));
            return this;
        }

        public NCellRange SetCellErrorValue(byte value, bool createMissingRowAndCol = true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetCellErrorValue(value));
            return this;
        }

        public NCellRange SetCellValue(DateTime value, bool createMissingRowAndCol = true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetCellValue(value));
            return this;
        }

        public NCellRange SetCellValue(IRichTextString value, bool createMissingRowAndCol = true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetCellValue(value));
            return this;
        }

        public NCellRange SetCellValue(string value, bool createMissingRowAndCol = true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetCellValue(value));
            return this;
        }

        public NCellRange RemoveFormula()
        {
            ForEachCell(false, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell =>
            {
                if(cell.CellFormula!=null)
                {
                    cell.RemoveFormula();
                }
            });
            return this;
        }

        public NCellRange SetCellFormula(string formula, bool createMissingRowAndCol = true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetCellFormula(formula));
            return this;
        }

        public NCellRange SetCellValue(bool value, bool createMissingRowAndCol = true)
        {
            ForEachCell(createMissingRowAndCol, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell => cell.SetCellValue(value));
            return this;
        }
        public string[][] Texts
        {
            get
            {
                var box = BoundingBox;
                string[][] texts= new string[Height][];
                for(int i = 0; i<Height; i++)
                {
                    texts[i]=new string[Width];
                    var row=_sheet.GetRow(box.FirstRow+i);
                    if(row==null)
                        continue;
                    for(int j = 0; j<Width; j++)
                    {
                        var c=row.GetCell(box.FirstColumn+j);
                        texts[i][j]=c?.ToString();
                    }
                }
                return texts;
            }
        }

        public object Value { 
            set {
                if(value is double || value is Double)
                {
                    SetCellValue((double) value);
                    return;
                }
                else if(value is string || value is String)
                {
                    SetCellValue((string) value);
                    return;
                }
                else if(value is bool || value is Boolean)
                {
                    SetCellValue((bool) value);
                    return;
                }
                else if(value is DateTime)
                {
                    SetCellValue((DateTime) value);
                    return;
                }
                else if(value is IRichTextString)
                {
                    SetCellValue((IRichTextString) value);
                    return;
                }
                throw new InvalidOperationException("invalid value type for cell value");
            } 
        }

        public double Sum()
        {
            double sum = 0;
            foreach (var cell in Cells)
            {
                if (cell.CellType == CellType.Numeric)
                    sum += cell.NumericCellValue;
            }
            return sum;
        }

        public double Min()
        {
            double? min = null;
            foreach (var cell in Cells)
            {
                if (cell.CellType == CellType.Numeric)
                {
                    double val = cell.NumericCellValue;
                    if (!min.HasValue || val < min.Value)
                        min = val;
                }
            }
            return min ?? double.NaN;
        }

        public double Max()
        {
            double? max = null;
            foreach (var cell in Cells)
            {
                if (cell.CellType == CellType.Numeric)
                {
                    double val = cell.NumericCellValue;
                    if (!max.HasValue || val > max.Value)
                        max = val;
                }
            }
            return max ?? double.NaN;
        }

        public double Avg()
        {
            double sum = 0;
            int count = 0;
            foreach (var cell in Cells)
            {
                if (cell.CellType == CellType.Numeric)
                {
                    sum += cell.NumericCellValue;
                    count++;
                }
            }
            return count > 0 ? sum / count : double.NaN;
        }

        public NCellRange RemoveCellComment()
        {
            ForEachCell(false, MissingCellPolicy.CREATE_NULL_AS_BLANK, cell =>
            {
                if(cell.CellComment!=null)
                {
                    cell.RemoveCellComment();
                }
            });
            return this;
        }

        public NCellRange RemoveHyperlink()
        {
            ForEachCell(false, MissingCellPolicy.RETURN_NULL_AND_BLANK, cell =>
            {
                if(cell.Hyperlink!=null)
                {
                    cell.RemoveHyperlink();
                }
            });
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public NCellRange this[string address] { 
            get {
                _ranges = CellRangeAddressList.Parse(address);
                if(_ranges.CountRanges()==0)
                    throw new ArgumentException($"cell range '{address}' is invalid");
                return this;
            } 
        }

        public NCellRange this[int row, int col]
        {
            get
            {
                validateRowCol(row, col);
                _ranges = new CellRangeAddressList(row, row, col, col);
                return this;
            }
        }
    }
}
