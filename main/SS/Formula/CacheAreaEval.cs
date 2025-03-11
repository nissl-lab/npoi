using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula
{
    public class CacheAreaEval:AreaEvalBase
    {
        readonly ValueEval[] _values;
        public CacheAreaEval(AreaI ptg, ValueEval[] values):base(ptg)
        {
            _values = values;
        }
        public CacheAreaEval(int firstRow, int firstColumn, int lastRow, int lastColumn, ValueEval[] values):
            base(firstRow, firstColumn, lastRow, lastColumn)
        {
            _values = values;
        }
        public override ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
        {
            return GetRelativeValue(-1, relativeRowIndex, relativeColumnIndex);
        }

        public override ValueEval GetRelativeValue(int sheetIndex, int relativeRowIndex, int relativeColumnIndex)
        {
            int oneDimensionalIndex = relativeRowIndex * Width + relativeColumnIndex;
            return _values[oneDimensionalIndex];
        }
        public override AreaEval Offset(int relFirstRowIx, int relLastRowIx,
        int relFirstColIx, int relLastColIx)
        {

            AreaI area = new OffsetArea(FirstRow, FirstColumn,
                    relFirstRowIx, relLastRowIx, relFirstColIx, relLastColIx);

            int height = area.LastRow - area.FirstRow + 1;
            int width = area.LastColumn - area.FirstColumn + 1;

            ValueEval[] newVals = new ValueEval[height * width];

            int startRow = area.FirstRow - FirstRow;
            int startCol = area.FirstColumn - FirstColumn;

            for (int j = 0; j < height; j++)
            {
                for (int i = 0; i < width; i++)
                {
                    ValueEval temp;

                    /* CacheAreaEval is only temporary value representation, does not equal sheet selection
                     * so any attempts going beyond the selection results in BlankEval
                     */
                    if (startRow + j > LastRow || startCol + i > LastColumn)
                    {
                        temp = BlankEval.instance;
                    }
                    else
                    {
                        temp = _values[(startRow + j) * Width + (startCol + i)];
                    }
                    newVals[j * width + i] = temp;
                }
            }

            return new CacheAreaEval(area, newVals);
        }
        public override TwoDEval GetRow(int rowIndex)
        {
            if (rowIndex >= Height)
            {
                throw new ArgumentException("Invalid rowIndex " + rowIndex
                        + ".  Allowable range is (0.." + Height + ").");
            }
            int absRowIndex = FirstRow + rowIndex;
            ValueEval[] values = new ValueEval[Width];

            for (int i = 0; i < values.Length; i++)
            {
                values[i] = GetRelativeValue(rowIndex, i);
            }
            return new CacheAreaEval(absRowIndex, FirstColumn, absRowIndex, LastColumn, values);
        }

        public override TwoDEval GetColumn(int columnIndex)
        {
            if (columnIndex >= Width)
            {
                throw new ArgumentException("Invalid columnIndex " + columnIndex
                        + ".  Allowable range is (0.." + Width + ").");
            }
            int absColIndex = FirstColumn+ columnIndex;
            ValueEval[] values = new ValueEval[Height];

            for (int i = 0; i < values.Length; i++)
            {
                values[i] = GetRelativeValue(i, columnIndex);
            }

            return new CacheAreaEval(FirstRow, absColIndex, LastRow, absColIndex, values);
        }

        public override string ToString()
        {
            CellReference crA = new CellReference(FirstRow, FirstColumn);
            CellReference crB = new CellReference(LastRow, LastColumn);
            return GetType().Name + "[" +
                    crA.FormatAsString() +
                    ':' +
                    crB.FormatAsString() +
                    "]";
        }
    }
}
