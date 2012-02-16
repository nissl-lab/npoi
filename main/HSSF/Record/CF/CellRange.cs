/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.Record.CF
{
    using System;
    using System.Collections;
    using NPOI.SS.Util;

    /**
     * 
     * @author Dmitriy Kumshayev
     */
    [Obsolete]
    public class CellRange
    {
        /** max 65536 rows in BIFF8 */
        private static int LAST_ROW_INDEX = 0x00FFFF;
        /** max 256 columns in BIFF8 */
        private static int LAST_COLUMN_INDEX = 0x00FF;

        private static Region[] EMPTY_REGION_ARRAY = { };

        private int _firstRow;
        private int _lastRow;
        private int _firstColumn;
        private int _lastColumn;

        /**
         * 
         * @param firstRow
         * @param lastRow pass <c>-1</c> for full column ranges
         * @param firstColumn
         * @param lastColumn  pass <c>-1</c> for full row ranges
         */
        public CellRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (!IsValid(firstRow, lastRow, firstColumn, lastColumn))
            {
                throw new ArgumentException("invalid cell range (" + firstRow + ", " + lastRow
                        + ", " + firstColumn + ", " + lastColumn + ")");
            }
            _firstRow = firstRow;
            _lastRow = ConvertM1ToMax(lastRow, LAST_ROW_INDEX);
            _firstColumn = firstColumn;
            _lastColumn = ConvertM1ToMax(lastColumn, LAST_COLUMN_INDEX);
        }

        /** 
         * Range arithmetic Is easier when using a large positive number for 'max row or column' 
         * instead of <c>-1</c>. 
         */
        private static int ConvertM1ToMax(int lastIx, int maxIndex)
        {
            if (lastIx < 0)
            {
                return maxIndex;
            }
            return lastIx;
        }

        public bool IsFullColumnRange()
        {
            return _firstRow == 0 && _lastRow == LAST_ROW_INDEX;
        }
        public bool IsFullRowRange()
        {
            return _firstColumn == 0 && _lastColumn == LAST_COLUMN_INDEX;
        }

        private static CellRange CreateFromRegion(Region r)
        {
            return new CellRange(r.RowFrom, r.RowTo, r.ColumnFrom, r.ColumnTo);
        }

        private static bool IsValid(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (lastRow < 0 || lastRow > LAST_ROW_INDEX)
            {
                return false;
            }
            if (firstRow < 0 || firstRow > LAST_ROW_INDEX)
            {
                return false;
            }

            if (lastColumn < 0 || lastColumn > LAST_COLUMN_INDEX)
            {
                return false;
            }
            if (firstColumn < 0 || firstColumn > LAST_COLUMN_INDEX)
            {
                return false;
            }
            return true;
        }

        public int FirstRow
        {
            get{return _firstRow;}
        }
        public int LastRow
        {
            get{return _lastRow;}
        }
        public int FirstColumn
        {
            get{return _firstColumn;}
        }
        public int LastColumn
        {
            get { return _lastColumn; }
        }

        public const int NO_INTERSECTION = 1;
        public const int OVERLAP = 2;
        /** first range Is within the second range */
        public const int INSIDE = 3;
        /** first range encloses or Is equal to the second */
        public const int ENCLOSES = 4;

        /**
         * Intersect this range with the specified range.
         * 
         * @param another - the specified range
         * @return code which reflects how the specified range Is related to this range.<br/>
         * Possible return codes are:	
         * 		NO_INTERSECTION - the specified range Is outside of this range;<br/> 
         * 		OVERLAP - both ranges partially overlap;<br/>
         * 		INSIDE - the specified range Is inside of this one<br/>
         * 		ENCLOSES - the specified range encloses (possibly exactly the same as) this range<br/>
         */
        public int intersect(CellRange another)
        {

            int firstRow = another.FirstRow;
            int lastRow = another.LastRow;
            int firstCol = another.FirstColumn;
            int lastCol = another.LastColumn;

            if
            (
                    gt(FirstRow, lastRow) ||
                    lt(LastRow, firstRow) ||
                    gt(FirstColumn, lastCol) ||
                    lt(LastColumn, firstCol)
            )
            {
                return NO_INTERSECTION;
            }
            else if (Contains(another))
            {
                return INSIDE;
            }
            else if (another.Contains(this))
            {
                return ENCLOSES;
            }
            else
            {
                return OVERLAP;
            }

        }

        /**
         * Do all possible cell merges between cells of the list so that:<br/>
         * 	- if a cell range Is completely inside of another cell range, it Gets Removed from the list 
         * 	- if two cells have a shared border, merge them into one bigger cell range
         * @param cellRangeList
         * @return updated List of cell ranges
         */
        public static CellRange[] MergeCellRanges(CellRange[] cellRanges)
        {
            if (cellRanges.Length < 1)
            {
                return cellRanges;
            }
            IList temp = MergeCellRanges(cellRanges);
            return ToArray(temp);
        }
        private static IList MergeCellRanges(IList cellRangeList)
        {

            while (cellRangeList.Count > 1)
            {
                bool somethingGotMerged = false;

                for (int i = 0; i < cellRangeList.Count; i++)
                {
                    CellRange range1 = (CellRange)cellRangeList[i];
                    for (int j = i + 1; j < cellRangeList.Count; j++)
                    {
                        CellRange range2 = (CellRange)cellRangeList[j];

                        CellRange[] mergeResult = MergeRanges(range1, range2);
                        if (mergeResult == null)
                        {
                            continue;
                        }
                        somethingGotMerged = true;
                        // overWrite range1 with first result 
                        cellRangeList[i]= mergeResult[0];
                        // Remove range2
                        cellRangeList.Remove(j--);
                        // Add any extra results beyond the first
                        for (int k = 1; k < mergeResult.Length; k++)
                        {
                            j++;
                            cellRangeList.Insert(j, mergeResult[k]);
                        }
                    }
                }
                if (!somethingGotMerged)
                {
                    break;
                }
            }


            return cellRangeList;
        }

        /**
         * @return the new range(s) to Replace the supplied ones.  <code>null</code> if no merge Is possible
         */
        private static CellRange[] MergeRanges(CellRange range1, CellRange range2)
        {

            int x = range1.intersect(range2);
            switch (x)
            {
                case CellRange.NO_INTERSECTION:
                    if (range1.HasExactSharedBorder(range2))
                    {
                        return new CellRange[] { range1.CreateEnclosingCellRange(range2), };
                    }
                    // else - No intersection and no shared border: do nothing 
                    return null;
                case CellRange.OVERLAP:
                    return resolveRangeOverlap(range1, range2);
                case CellRange.INSIDE:
                    // Remove range2, since it Is completely inside of range1
                    return new CellRange[] { range1, };
                case CellRange.ENCLOSES:
                    // range2 encloses range1, so Replace it with the enclosing one
                    return new CellRange[] { range2, };
            }
            throw new Exception("unexpected intersection result (" + x + ")");
        }

        // TODO - Write junit test for this
        static CellRange[] resolveRangeOverlap(CellRange rangeA, CellRange rangeB)
        {

            if (rangeA.IsFullColumnRange())
            {
                if (rangeB.IsFullRowRange())
                {
                    // Excel seems to leave these Unresolved
                    return null;
                }
                return rangeA.sliceUp(rangeB);
            }
            if (rangeA.IsFullRowRange())
            {
                if (rangeB.IsFullColumnRange())
                {
                    // Excel seems to leave these Unresolved
                    return null;
                }
                return rangeA.sliceUp(rangeB);
            }
            if (rangeB.IsFullColumnRange())
            {
                return rangeB.sliceUp(rangeA);
            }
            if (rangeB.IsFullRowRange())
            {
                return rangeB.sliceUp(rangeA);
            }
            return rangeA.sliceUp(rangeB);
        }

        /**
         * @param range never a full row or full column range
         * @return an array including <b>this</b> <c>CellRange</c> and all parts of <c>range</c> 
         * outside of this range  
         */
        private CellRange[] sliceUp(CellRange range)
        {

            ArrayList temp = new ArrayList();

            // Chop up range horizontally and vertically
            temp.Add(range);
            if (!IsFullColumnRange())
            {
                temp = CutHorizontally(_firstRow, temp);
                temp = CutHorizontally(_lastRow + 1, temp);
            }
            if (!IsFullRowRange())
            {
                temp = CutVertically(_firstColumn, temp);
                temp = CutVertically(_lastColumn + 1, temp);
            }
            CellRange[] crParts = ToArray(temp);

            // form result array
            temp.Clear();
            temp.Add(this);

            for (int i = 0; i < crParts.Length; i++)
            {
                CellRange crPart = crParts[i];
                // only include parts that are not enclosed by this
                if (intersect(crPart) != ENCLOSES)
                {
                    temp.Add(crPart);
                }
            }
            return ToArray(temp);
        }

        private static ArrayList CutHorizontally(int cutRow, IList input)
        {

            ArrayList result = new ArrayList();
            CellRange[] crs = ToArray(input);
            for (int i = 0; i < crs.Length; i++)
            {
                CellRange cr = crs[i];
                if (cr._firstRow < cutRow && cutRow < cr._lastRow)
                {
                    result.Add(new CellRange(cr._firstRow, cutRow, cr._firstColumn, cr._lastColumn));
                    result.Add(new CellRange(cutRow + 1, cr._lastRow, cr._firstColumn, cr._lastColumn));
                }
                else
                {
                    result.Add(cr);
                }
            }
            return result;
        }
        private static ArrayList CutVertically(int cutColumn, ArrayList input)
        {

            ArrayList result = new ArrayList();
            CellRange[] crs = ToArray(input);
            for (int i = 0; i < crs.Length; i++)
            {
                CellRange cr = crs[i];
                if (cr._firstColumn < cutColumn && cutColumn < cr._lastColumn)
                {
                    result.Add(new CellRange(cr._firstRow, cr._lastRow, cr._firstColumn, cutColumn));
                    result.Add(new CellRange(cr._firstRow, cr._lastRow, cutColumn + 1, cr._lastColumn));
                }
                else
                {
                    result.Add(cr);
                }
            }
            return result;
        }


        private static CellRange[] ToArray(IList temp)
        {
            CellRange[] cellranges=new CellRange[temp.Count];

            IEnumerator enumerator = temp.GetEnumerator();
            int i=0;
            while (enumerator.MoveNext())
            {
                cellranges[i] = (CellRange)enumerator.Current;
            }
            return cellranges;
        }

        /**
         * Convert array of regions to a List of CellRange objects
         *  
         * @param regions
         * @return List of CellRange objects
         */
        public static CellRange[] ConvertRegionsToCellRanges(Region[] regions)
        {
            CellRange[] result = new CellRange[regions.Length];
            for (int i = 0; i < regions.Length; i++)
            {
                result[i] = CreateFromRegion(regions[i]);
            }
            return result;
        }

        /**
         * Convert a List of CellRange objects to an array of regions 
         *  
         * @param List of CellRange objects
         * @return regions
         */
        public static Region[] ConvertCellRangesToRegions(CellRange[] cellRanges)
        {
            int size = cellRanges.Length;
            if (size < 1)
            {
                return EMPTY_REGION_ARRAY;
            }

            Region[] result = new Region[size];

            for (int i = 0; i != size; i++)
            {
                result[i] = cellRanges[i].ConvertToRegion();
            }
            return result;
        }



        private Region ConvertToRegion()
        {

            return new Region(_firstRow, (short)_firstColumn, _lastRow, (short)_lastColumn);
        }


        /**
         *  Check if the specified range Is located inside of this cell range.
         *  
         * @param range
         * @return true if this cell range Contains the argument range inside if it's area
         */
        public bool Contains(CellRange range)
        {
            int firstRow = range.FirstRow;
            int lastRow = range.LastRow;
            int firstCol = range.FirstColumn;
            int lastCol = range.LastColumn;
            return le(FirstRow, firstRow) && ge(LastRow, lastRow)
                    && le(FirstColumn, firstCol) && ge(LastColumn, lastCol);
        }

        public bool Contains(int row, short column)
        {
            return le(FirstRow, row) && ge(LastRow, row)
                    && le(FirstColumn, column) && ge(LastColumn, column);
        }

        /**
         * Check if the specified cell range Has a shared border with the current range.
         * 
         * @return <c>true</c> if the ranges have a complete shared border (i.e.
         * the two ranges toGether make a simple rectangular region.
         */
        public bool HasExactSharedBorder(CellRange range)
        {
            int oFirstRow = range._firstRow;
            int oLastRow = range._lastRow;
            int oFirstCol = range._firstColumn;
            int oLastCol = range._lastColumn;

            if (_firstRow > 0 && _firstRow - 1 == oLastRow ||
                oFirstRow > 0 && oFirstRow - 1 == _lastRow)
            {
                // ranges have a horizontal border in common
                // make sure columns are identical:
                return _firstColumn == oFirstCol && _lastColumn == oLastCol;
            }

            if (_firstColumn > 0 && _firstColumn - 1 == oLastCol ||
                oFirstCol > 0 && _lastColumn == oFirstCol - 1)
            {
                // ranges have a vertical border in common
                // make sure rows are identical:
                return _firstRow == oFirstRow && _lastRow == oLastRow;
            }
            return false;
        }

        /**
         * Create an enclosing CellRange for the two cell ranges.
         * 
         * @return enclosing CellRange
         */
        public CellRange CreateEnclosingCellRange(CellRange range)
        {
            if (range == null)
            {
                return CloneCellRange();
            }
            else
            {
                CellRange cellRange =
                    new CellRange(
                        lt(range.FirstRow, FirstRow) ? range.FirstRow : FirstRow,
                        gt(range.LastRow, LastRow) ? range.LastRow : LastRow,
                        lt(range.FirstColumn, FirstColumn) ? range.FirstColumn : FirstColumn,
                        gt(range.LastColumn, LastColumn) ? range.LastColumn : LastColumn
                    );
                return cellRange;
            }
        }

        public CellRange CloneCellRange()
        {
            return new CellRange(FirstRow, LastRow, FirstColumn, LastColumn);
        }

        /**
         * @return true if a &lt; b
         */
        private static bool lt(int a, int b)
        {
            return a == -1 ? false : (b == -1 ? true : a < b);
        }

        /**
         * @return true if a &lt;= b
         */
        private static bool le(int a, int b)
        {
            return a == b || lt(a, b);
        }

        /**
         * @return true if a > b
         */
        private static bool gt(int a, int b)
        {
            return lt(b, a);
        }

        /**
         * @return true if a >= b
         */
        private static bool ge(int a, int b)
        {
            return !lt(a, b);
        }

        public override String ToString()
        {
            return "(" + FirstRow + "," + LastRow + "," + FirstColumn + "," + LastColumn + ")";
        }

    }
}