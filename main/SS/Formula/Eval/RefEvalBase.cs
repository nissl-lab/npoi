
using System;
namespace NPOI.SS.Formula.Eval
{
    public abstract class RefEvalBase : RefEval
    {
        private int _firstSheetIndex;
        private int _lastSheetIndex;
        private int _rowIndex;
        private int _columnIndex;

        protected RefEvalBase(ISheetRange sheetRange, int rowIndex, int columnIndex)
        {
            if (sheetRange == null)
            {
                throw new ArgumentException("sheetRange must not be null");
            }
            _firstSheetIndex = sheetRange.FirstSheetIndex;
            _lastSheetIndex = sheetRange.LastSheetIndex;
            _rowIndex = rowIndex;
            _columnIndex = columnIndex;
        }
        protected RefEvalBase(int firstSheetIndex, int lastSheetIndex, int rowIndex, int columnIndex)
        {
            _firstSheetIndex = firstSheetIndex;
            _lastSheetIndex = lastSheetIndex;
            _rowIndex = rowIndex;
            _columnIndex = columnIndex;
        }
        protected RefEvalBase(int onlySheetIndex, int rowIndex, int columnIndex)
            : this(onlySheetIndex, onlySheetIndex, rowIndex, columnIndex)
        {
            ;
        }

        public int NumberOfSheets
        {
            get
            {
                return _lastSheetIndex - _firstSheetIndex + 1;
            }
        }
        public int FirstSheetIndex
        {
            get
            {
                return _firstSheetIndex;
            }
        }
        public int LastSheetIndex
        {
            get
            {
                return _lastSheetIndex;
            }
        }
        public int Row
        {
            get
            {
                return _rowIndex;
            }
        }
        public int Column
        {
            get
            {
                return _columnIndex;
            }
        }
        public abstract ValueEval GetInnerValueEval(int sheetIndex);

        public abstract AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx);
    }
}