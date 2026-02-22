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
        private CellRangeAddress _address;

        public int Width => _address.LastColumn - _address.FirstColumn+1;

        public int Height => _address.LastRow - _address.FirstRow+1;

        public int Size => _address.NumberOfCells;

        public string Address => _address.FormatAsString();
        public ICell TopLeftCell => GetCell(0,0);
        public ISheet Sheet => _sheet;

        public string CellFormula { get => throw new NotImplementedException(); set => this.SetCellFormula(value); }

        public ICellStyle CellStyle
        {
            set 
            {
                for(int i = _address.FirstRow; i<=_address.LastRow; i++)
                {
                    var row = _sheet.GetRow(i);
                    if(row==null)
                        continue;
                    for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                    {
                        var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if(cell!=null)
                            cell.CellStyle=value;
                    }
                }
            }
        }

        public void SetCellComment(IComment comment) {
            if(comment==null)
            {
                //remove comment
                return;
            }
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    comment.SetAddress(i, j);
                }
            }
        }
        public IHyperlink Hyperlink { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public NCellRange(ISheet sheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            _sheet = sheet;
            validateRowCol(fromRow, fromCol);
            validateRowCol(toRow, toCol);
            _address = new CellRangeAddress(fromRow, toRow, fromCol, toCol);
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
            return _sheet.GetRow(_address.FirstRow+rowInRange)?.GetCell(_address.FirstColumn+colInRange);
        }

        public IEnumerator<ICell> GetEnumerator()
        {
            List<ICell> cells = new List<ICell>(); 
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                    continue;
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if(cell!=null)
                        cells.Add(cell);
                }
            }
            return cells.GetEnumerator();
        }

        public void SetCellType(CellType cellType, bool createMissingRowAndCol= true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellType(cellType);
                }
            }
        }

        public void SetBlank(bool createMissingRowAndCol = false)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetBlank();
                }
            }
        }

        public void SetCellValue(double value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellValue(value);
                }
            }
        }

        public void SetCellErrorValue(byte value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellErrorValue(value);
                }
            }
        }

        public void SetCellValue(DateTime value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellValue(value);
                }
            }
        }

        public void SetCellValue(IRichTextString value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellValue(value);
                }
            }
        }

        public void SetCellValue(string value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellValue(value);
                }
            }
        }

        public void RemoveFormula()
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.RemoveFormula();
                }
            }
        }

        public void SetCellFormula(string formula, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellFormula(formula);
                }
            }
        }

        public void SetCellValue(bool value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null&&!createMissingRowAndCol)
                {
                    continue;
                }
                row = _sheet.CreateRow(i);
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellValue(value);
                }
            }
        }

        public void RemoveCellComment()
        {
            throw new NotImplementedException();
        }

        public void RemoveHyperlink()
        {
            throw new NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public NCellRange this[string address] { 
            get {
                _address = CellRangeAddress.ValueOf(address);
                return this;
            } 
        }

        public NCellRange this[int row, int col]
        {
            get
            {
                validateRowCol(row, col);
                _address.FirstRow = row;
                _address.LastRow = row;
                _address.FirstColumn = col;
                _address.LastColumn = col;
                return this;
            }
        }
    }
}
