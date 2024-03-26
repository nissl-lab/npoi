using System.Collections.Generic;

namespace NPOI.SS.UserModel
{
    public interface ICellRange<T> : IEnumerable<T> where T : ICell
    {
        int Width { get; }
        int Height { get; }
        /// <summary>
        /// Gets the number of cells in this range.
        /// </summary>
        /// <return><tt>height * width </tt></return>
        int Size { get; }
        /// <summary>
        /// the text format of this range.  Single cell ranges are formatted
        /// like single cell references (e.g. 'A1' instead of 'A1:A1').
        /// </summary>
        string ReferenceText { get; }
        /// <summary>
        /// the cell at relative coordinates (0,0).  Never <c>null</c>.
        /// </summary>
        T TopLeftCell { get; }
        /// <summary>
        /// </summary>
        /// <param name="relativeRowIndex">must be between <tt>0</tt> and <tt>height-1</tt></param>
        /// <param name="relativeColumnIndex">must be between <tt>0</tt> and <tt>width-1</tt></param>
        /// <return>the cell at the specified coordinates.  Never <c>null</c>.</return>
        T GetCell(int relativeRowIndex, int relativeColumnIndex);
        /// <summary>
        /// a flattened array of all the cells in this <see cref="ICellStyle"/>
        /// </summary>
        T[] FlattenedCells { get; }
        /// <summary>
        /// a 2-D array of all the cells in this <see cref="ICellStyle"/>.  The first
        /// array dimension is the row index (values <tt>0...height-1</tt>)
        /// and the second dimension is the column index (values <tt>0...width-1</tt>)
        /// </summary>
        T[][] Cells { get; }
    }
}
