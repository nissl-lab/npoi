using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.UserModel
{
    public class NRowRange:IEnumerable<IRow>
    {
        private int _fromRow;
        private int _toRow;
        private ISheet _sheet;

        public NRowRange(ISheet sheet, int fromRow, int toRow)
        {
            _fromRow=fromRow;
            _toRow=toRow;
            _sheet=sheet;
        }

        private void validateRow(int col)
        {
            if(col<0||col>_sheet.Workbook.SpreadsheetVersion.MaxColumns)
                throw new ArgumentException($"column index {col} is out of range");
        }

        public IEnumerator<IRow> GetEnumerator()
        {
            return Rows.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int RowCount { get => _toRow-_fromRow+1; }

        public short Height
        {
            get=> throw new NotImplementedException();
            set {
                for(int i = _fromRow; i<=_toRow; i++)
                {
                    var row = _sheet.GetRow(i);
                    if(row==null)
                        row=_sheet.CreateRow(i);
                    row.Height=value;
                }
            }
        }
        public float HeightInPoints
        {
            get => throw new NotImplementedException();
            set
            {
                for(int i = _fromRow; i<=_toRow; i++)
                {
                    var row = _sheet.GetRow(i);
                    if(row==null)
                        row=_sheet.CreateRow(i);
                    row.HeightInPoints=value;
                }
            }
        }
        public bool ZeroHeight
        {
            get => throw new NotImplementedException();
            set
            {
                for(int i = _fromRow; i<=_toRow; i++)
                {
                    var row = _sheet.GetRow(i);
                    if(row==null)
                        row=_sheet.CreateRow(i);
                    row.ZeroHeight=value;
                }
            }
        }
        public ICellStyle RowStyle
        {
            get => throw new NotImplementedException();
            set
            {
                if(RowStyle==null)
                    return;
                for(int i = _fromRow; i<=_toRow; i++)
                {
                    var row = _sheet.GetRow(i);
                    if(row==null)
                        row=_sheet.CreateRow(i);
                    row.RowStyle=value;
                }
            }
        }
        public NRowRange Group()
        {
            _sheet.GroupRow(_fromRow, _toRow);
            return this;
        }
        public List<IRow> Rows
        {
            get
            { 
                var rows= new List<IRow>();
                for(int i = _fromRow; i<=_toRow; i++)
                {
                    var row = _sheet.GetRow(i);
                    if(row==null)
                        continue;
                    rows.Add(row);
                }
                return rows;
            }
        }
        public NRowRange this[int fromRow, int toRow]
        {
            get
            {
                validateRow(fromRow);
                validateRow(toRow);
                _fromRow = fromRow;
                _toRow = toRow;
                return this;
            }
        }
    }
}
