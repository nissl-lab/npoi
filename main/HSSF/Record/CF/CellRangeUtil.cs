/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.Record.CF
{
    using System;
    using System.Collections;
    using NPOI.SS.Util;
    using System.Collections.Generic;

    /**
     * 
     * @author Dmitriy Kumshayev
     */
    public class CellRangeUtil
    {

        private CellRangeUtil()
        {
            // no instance of this class
        }

        public const int NO_INTERSECTION = 1;
        public const int OVERLAP = 2;
        /** first range is within the second range */
        public const int INSIDE = 3;
        /** first range encloses or is equal to the second */
        public const int ENCLOSES = 4;

        /**
         * Intersect this range with the specified range.
         * 
         * @param crB - the specified range
         * @return code which reflects how the specified range is related to this range.<br/>
         * Possible return codes are:	
         * 		NO_INTERSECTION - the specified range is outside of this range;<br/> 
         * 		OVERLAP - both ranges partially overlap;<br/>
         * 		INSIDE - the specified range is inside of this one<br/>
         * 		ENCLOSES - the specified range encloses (possibly exactly the same as) this range<br/>
         */
        public static int Intersect(CellRangeAddress crA, CellRangeAddress crB)
        {

            int firstRow = crB.FirstRow;
            int lastRow = crB.LastRow;
            int firstCol = crB.FirstColumn;
            int lastCol = crB.LastColumn;

            if
            (
                    gt(crA.FirstRow, lastRow) ||
                    lt(crA.LastRow, firstRow) ||
                    gt(crA.FirstColumn, lastCol) ||
                    lt(crA.LastColumn, firstCol)
            )
            {
                return NO_INTERSECTION;
            }
            else if (Contains(crA, crB))
            {
                return INSIDE;
            }
            else if (Contains(crB, crA))
            {
                return ENCLOSES;
            }
            else
            {
                return OVERLAP;
            }

        }

        /**
         * Do all possible cell merges between cells of the list so that:
         * 	if a cell range is completely inside of another cell range, it s removed from the list 
         * 	if two cells have a shared border, merge them into one bigger cell range
         * @param cellRangeList
         * @return updated List of cell ranges
         */
        public static CellRangeAddress[] MergeCellRanges(CellRangeAddress[] cellRanges)
        {
            if (cellRanges.Length < 1)
            {
                return cellRanges;
            }
            //ArrayList temp = MergeCellRanges(NPOI.Util.Arrays.AsList(cellRanges));
            List<CellRangeAddress> lst = new List<CellRangeAddress>(cellRanges);
            List<CellRangeAddress> temp = MergeCellRanges(lst);
            return temp.ToArray();
        }
        private static List<CellRangeAddress> MergeCellRanges(List<CellRangeAddress> cellRangeList)
        {
            // loop until either only one item is left or we did not merge anything any more
            while (cellRangeList.Count > 1)
            {
                bool somethingGotMerged = false;
                // look at all cell-ranges
                for (int i = 0; i < cellRangeList.Count; i++)
                {
                    CellRangeAddress range1 = cellRangeList[i];
                    // compare each cell range to all other cell-ranges
                    for (int j = i + 1; j < cellRangeList.Count; j++)
                    {
                        CellRangeAddress range2 = cellRangeList[j];

                        CellRangeAddress[] mergeResult = MergeRanges(range1, range2);
                        if (mergeResult == null)
                        {
                            continue;
                        }
                        somethingGotMerged = true;
                        // overwrite range1 with first result 
                        cellRangeList[i] = mergeResult[0];
                        // remove range2
                        cellRangeList.RemoveAt(j--);
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
         * @return the new range(s) to replace the supplied ones.  <c>null</c> if no merge is possible
         */
        private static CellRangeAddress[] MergeRanges(CellRangeAddress range1, CellRangeAddress range2)
        {

            int x = Intersect(range1, range2);
            switch (x)
            {
                // nothing in common: at most they could be adjacent to each other and thus form a single bigger area  
                case CellRangeUtil.NO_INTERSECTION:
                    if (HasExactSharedBorder(range1, range2))
                    {
                        return new CellRangeAddress[] { CreateEnclosingCellRange(range1, range2), };
                    }
                    // else - No intersection and no shared border: do nothing 
                    return null;
                case CellRangeUtil.OVERLAP:
                    // commented out the cells overlap implementation, it caused endless loops, see Bug 55380
                    // disabled for now, the algorithm will not detect some border cases this way currently!
                    //return ResolveRangeOverlap(range1, range2);
                    return null;
                case CellRangeUtil.INSIDE:
                    // Remove range2, since it is completely inside of range1
                    return new CellRangeAddress[] { range1 };
                case CellRangeUtil.ENCLOSES:
                    // range2 encloses range1, so replace it with the enclosing one
                    return new CellRangeAddress[] { range2 };
            }
            throw new InvalidOperationException("unexpected intersection result (" + x + ")");
        }

        //// TODO - write junit test for this
        //static CellRangeAddress[] ResolveRangeOverlap(CellRangeAddress rangeA, CellRangeAddress rangeB)
        //{

        //    if (rangeA.IsFullColumnRange)
        //    {
        //        if (rangeA.IsFullRowRange)
        //        {
        //            // Excel seems to leave these unresolved
        //            return null;
        //        }
        //        return SliceUp(rangeA, rangeB);
        //    }
        //    if (rangeA.IsFullRowRange)
        //    {
        //        if (rangeB.IsFullColumnRange)
        //        {
        //            // Excel seems to leave these unresolved
        //            return null;
        //        }
        //        return SliceUp(rangeA, rangeB);
        //    }
        //    if (rangeB.IsFullColumnRange)
        //    {
        //        return SliceUp(rangeB, rangeA);
        //    }
        //    if (rangeB.IsFullRowRange)
        //    {
        //        return SliceUp(rangeB, rangeA);
        //    }
        //    return SliceUp(rangeA, rangeB);
        //}

        ///**
        // * @param crB never a full row or full column range
        // * @return an array including <b>this</b> <c>CellRange</c> and all parts of <c>range</c> 
        // * outside of this range  
        // */
        //private static CellRangeAddress[] SliceUp(CellRangeAddress crA, CellRangeAddress crB)
        //{

        //    ArrayList temp = new ArrayList();

        //    // Chop up range horizontally and vertically
        //    temp.Add(crB);
        //    if (!crA.IsFullColumnRange)
        //    {
        //        temp = CutHorizontally(crA.FirstRow, temp);
        //        temp = CutHorizontally(crA.LastRow + 1, temp);
        //    }
        //    if (!crA.IsFullRowRange)
        //    {
        //        temp = CutVertically(crA.FirstColumn, temp);
        //        temp = CutVertically(crA.LastColumn + 1, temp);
        //    }
        //    CellRangeAddress[] crParts = ToArray(temp);

        //    // form result array
        //    temp.Clear();
        //    temp.Add(crA);

        //    for (int i = 0; i < crParts.Length; i++)
        //    {
        //        CellRangeAddress crPart = crParts[i];
        //        // only include parts that are not enclosed by this
        //        if (Intersect(crA, crPart) != ENCLOSES)
        //        {
        //            temp.Add(crPart);
        //        }
        //    }
        //    return ToArray(temp);
        //}

        //private static ArrayList CutHorizontally(int cutRow, ArrayList input)
        //{

        //    ArrayList result = new ArrayList();
        //    CellRangeAddress[] crs = ToArray(input);
        //    for (int i = 0; i < crs.Length; i++)
        //    {
        //        CellRangeAddress cr = crs[i];
        //        if (cr.FirstRow < cutRow && cutRow < cr.LastRow)
        //        {
        //            result.Add(new CellRangeAddress(cr.FirstRow, cutRow, cr.FirstColumn, cr.LastColumn));
        //            result.Add(new CellRangeAddress(cutRow + 1, cr.LastRow, cr.FirstColumn, cr.LastColumn));
        //        }
        //        else
        //        {
        //            result.Add(cr);
        //        }
        //    }
        //    return result;
        //}
        //private static ArrayList CutVertically(int cutColumn, ArrayList input)
        //{

        //    ArrayList result = new ArrayList();
        //    CellRangeAddress[] crs = ToArray(input);
        //    for (int i = 0; i < crs.Length; i++)
        //    {
        //        CellRangeAddress cr = crs[i];
        //        if (cr.FirstColumn < cutColumn && cutColumn < cr.LastColumn)
        //        {
        //            result.Add(new CellRangeAddress(cr.FirstRow, cr.LastRow, cr.FirstColumn, cutColumn));
        //            result.Add(new CellRangeAddress(cr.FirstRow, cr.LastRow, cutColumn + 1, cr.LastColumn));
        //        }
        //        else
        //        {
        //            result.Add(cr);
        //        }
        //    }
        //    return result;
        //}


        private static CellRangeAddress[] ToArray(ArrayList temp)
        {
            CellRangeAddress[] result = new CellRangeAddress[temp.Count];
            result = (CellRangeAddress[])temp.ToArray(typeof(CellRangeAddress));
            return result;
        }



        /**
         *  Check if the specified range is located inside of this cell range.
         *  
         * @param crB
         * @return true if this cell range Contains the argument range inside if it's area
         */
        public static bool Contains(CellRangeAddress crA, CellRangeAddress crB)
        {
            int firstRow = crB.FirstRow;
            int lastRow = crB.LastRow;
            int firstCol = crB.FirstColumn;
            int lastCol = crB.LastColumn;
            return le(crA.FirstRow, firstRow) && ge(crA.LastRow, lastRow)
                    && le(crA.FirstColumn, firstCol) && ge(crA.LastColumn, lastCol);
        }

        /**
         * Check if the specified cell range has a shared border with the current range.
         * 
         * @return <c>true</c> if the ranges have a complete shared border (i.e.
         * the two ranges toher make a simple rectangular region.
         */
        public static bool HasExactSharedBorder(CellRangeAddress crA, CellRangeAddress crB)
        {
            int oFirstRow = crB.FirstRow;
            int oLastRow = crB.LastRow;
            int oFirstCol = crB.FirstColumn;
            int oLastCol = crB.LastColumn;

            if (crA.FirstRow > 0 && crA.FirstRow - 1 == oLastRow ||
                oFirstRow > 0 && oFirstRow - 1 == crA.LastRow)
            {
                // ranges have a horizontal border in common
                // make sure columns are identical:
                return crA.FirstColumn == oFirstCol && crA.LastColumn == oLastCol;
            }

            if (crA.FirstColumn > 0 && crA.FirstColumn - 1 == oLastCol ||
                oFirstCol > 0 && crA.LastColumn == oFirstCol - 1)
            {
                // ranges have a vertical border in common
                // make sure rows are identical:
                return crA.FirstRow == oFirstRow && crA.LastRow == oLastRow;
            }
            return false;
        }

        /**
         * Create an enclosing CellRange for the two cell ranges.
         * 
         * @return enclosing CellRange
         */
        public static CellRangeAddress CreateEnclosingCellRange(CellRangeAddress crA, CellRangeAddress crB)
        {
            if (crB == null)
            {
                return crA.Copy();
            }

            return
                new CellRangeAddress(
                    lt(crB.FirstRow, crA.FirstRow) ? crB.FirstRow : crA.FirstRow,
                    gt(crB.LastRow, crA.LastRow) ? crB.LastRow : crA.LastRow,
                    lt(crB.FirstColumn, crA.FirstColumn) ? crB.FirstColumn : crA.FirstColumn,
                    gt(crB.LastColumn, crA.LastColumn) ? crB.LastColumn : crA.LastColumn
                );

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
    }
}