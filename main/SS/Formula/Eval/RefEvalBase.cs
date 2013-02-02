
namespace NPOI.SS.Formula.Eval
{
    public abstract class RefEvalBase : RefEval
    {

        private int _rowIndex;
        private int _columnIndex;

        protected RefEvalBase(int rowIndex, int columnIndex)
        {
            _rowIndex = rowIndex;
            _columnIndex = columnIndex;
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
        public abstract ValueEval InnerValueEval { get; }

        public abstract AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx);
    }
}