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

        public string Formula { get => throw new NotImplementedException(); set => this.SetCellFormula(value); }

        public NCellRange SetCellStyle(ICellStyle style, bool createMissingRowAndCol)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.CellStyle=style;
                }
            }
            return this;
        }

        public void SetCellComment(IComment comment) {
            if(comment==null)
            {
                RemoveCellComment();
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
        public NCellRange SetHyperlink(IHyperlink hyperlink, bool createMissingRowAndCol=true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.Hyperlink = hyperlink;
                }
            }
            return this;
        }

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
            var row = _sheet.GetRow(_address.FirstRow+rowInRange);
            if(row==null)
                row = _sheet.CreateRow(_address.FirstRow+rowInRange);
            return row.GetCell(_address.FirstColumn+colInRange, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }

        public List<ICell> Cells
        {
            get {
                List<ICell> cells = new List<ICell>();
                for(int i = _address.FirstRow; i<=((_address.LastRow<_sheet.LastRowNum)?_address.LastRow: _sheet.LastRowNum); i++)
                {
                    var row = _sheet.GetRow(i);
                    if(row==null)
                        continue;
                    for(int j = _address.FirstColumn; j<=((_address.LastColumn<row.LastCellNum)?_address.LastColumn:row.LastCellNum); j++)
                    {
                        var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if(cell!=null)
                            cells.Add(cell);
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
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.SetCellType(cellType);
                }
            }
            return this;
        }

        public NCellRange SetBlank(bool createMissingRowAndCol = false)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.SetBlank();
                }
            }
            return this;
        }

        public NCellRange SetCellValue(double value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.SetCellValue(value);
                }
            }
            return this;
        }

        public NCellRange SetCellErrorValue(byte value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }

                    cell = row.CreateCell(j);
                    cell.SetCellErrorValue(value);
                }
            }
            return this;
        }

        public NCellRange SetCellValue(DateTime value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.SetCellValue(value);
                }
            }
            return this;
        }

        public NCellRange SetCellValue(IRichTextString value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.SetCellValue(value);
                }
            }
            return this;
        }

        public NCellRange SetCellValue(string value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.SetCellValue(value);
                }
            }
            return this;
        }

        public NCellRange RemoveFormula()
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    continue;
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j);
                    if(cell!=null&&cell.CellFormula!=null)
                    {
                        cell.RemoveFormula();
                    }
                }
            }
            return this;
        }

        public NCellRange SetCellFormula(string formula, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if(cell==null)
                    {
                        if(!createMissingRowAndCol)
                            continue;
                        else
                            cell = row.CreateCell(j);
                    }
                    cell.SetCellFormula(formula);
                }
            }
            return this;
        }

        public NCellRange SetCellValue(bool value, bool createMissingRowAndCol = true)
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    if(!createMissingRowAndCol)
                        continue;
                    else
                        row = _sheet.CreateRow(i);
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell==null&&!createMissingRowAndCol)
                        continue;

                    cell = row.CreateCell(j);
                    cell.SetCellValue(value);
                }
            }
            return this;
        }
        public string[][] Texts
        {
            get
            {
                string[][] texts= new string[Height][];
                for(int i = _address.FirstRow; i<=_address.LastRow; i++)
                {
                    bool emptyRow = false;
                    var row=_sheet.GetRow(i);
                    if(row==null)
                    {
                        emptyRow=true;
                    }
                    texts[i-_address.FirstRow]=new string[Width];
                    for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                    {
                        if(emptyRow)
                        {
                            texts[i-_address.FirstRow][j-_address.FirstColumn]=null;
                            break;
                        }
                        var c=row.GetCell(j);
                        if(c==null)
                        {
                            texts[i-_address.FirstRow][j-_address.FirstColumn]=null;
                        }
                        else
                        {
                            texts[i-_address.FirstRow][j-_address.FirstColumn]=c.ToString();
                        }
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
        public NCellRange RemoveCellComment()
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    continue;
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j,MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cell.CellComment!=null)
                    {
                        cell.RemoveCellComment();
                    }
                }
            }
            return this;
        }

        public NCellRange RemoveHyperlink()
        {
            for(int i = _address.FirstRow; i<=_address.LastRow; i++)
            {
                var row = _sheet.GetRow(i);
                if(row==null)
                {
                    continue;
                }
                for(int j = _address.FirstColumn; j<=_address.LastColumn; j++)
                {
                    var cell = row.GetCell(j);
                    if(cell==null)
                        continue;
                    if(cell.Hyperlink!=null)
                    {
                        cell.RemoveHyperlink();
                    }
                }
            }
            return this;
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
