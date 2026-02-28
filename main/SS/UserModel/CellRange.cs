using System.Collections.Generic;

namespace NPOI.SS.UserModel
{
    public interface ICellRange<T> : IEnumerable<T> where T : ICell
    {
        int Width { get; }
        int Height { get; }
        /**
         * Gets the number of cells in this range.
         * @return <tt>height * width </tt>
         */
        int Size { get; }
        /**
         * @return the text format of this range.  Single cell ranges are formatted
         *         like single cell references (e.g. 'A1' instead of 'A1:A1').
         */
        string ReferenceText { get; }
        /**
         * @return the cell at relative coordinates (0,0).  Never <code>null</code>.
         */
        T TopLeftCell { get; }
        /**
         * @param relativeRowIndex must be between <tt>0</tt> and <tt>height-1</tt>
         * @param relativeColumnIndex must be between <tt>0</tt> and <tt>width-1</tt>
         * @return the cell at the specified coordinates.  Never <code>null</code>.
         */
        T GetCell(int relativeRowIndex, int relativeColumnIndex);
        /**
         * @return a flattened array of all the cells in this {@link CellRange}
         */
        T[] FlattenedCells{get;}
        /**
         * @return a 2-D array of all the cells in this {@link CellRange}.  The first
         * array dimension is the row index (values <tt>0...height-1</tt>)
         * and the second dimension is the column index (values <tt>0...width-1</tt>)
         */
        T[][] Cells { get; }
    }
}
