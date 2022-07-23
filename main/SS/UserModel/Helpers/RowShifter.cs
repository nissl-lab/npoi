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

namespace NPOI.SS.UserModel.Helpers
{
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /**
     * Helper for Shifting rows up or down
     * 
     * This abstract class exists to consolidate duplicated code between XSSFRowShifter and HSSFRowShifter (currently methods sprinkled throughout HSSFSheet)
     */
    public abstract class RowShifter
    {
        protected ISheet sheet;

        public RowShifter(ISheet sh)
        {
            sheet = sh;
        }

        /**
         * Shifts, grows, or shrinks the merged regions due to a row Shift.
         * Merged regions that are completely overlaid by Shifting will be deleted.
         *
         * @param startRow the row to start Shifting
         * @param endRow   the row to end Shifting
         * @param n        the number of rows to shift
         * @return an array of affected merged regions, doesn't contain deleted ones
         */
        public List<CellRangeAddress> ShiftMergedRegions(int startRow, int endRow, int n)
        {
            List<CellRangeAddress> ShiftedRegions = new List<CellRangeAddress>();
            ISet<int> removedIndices = new HashSet<int>();
            //move merged regions completely if they fall within the new region boundaries when they are Shifted
            int size = sheet.NumMergedRegions;
            for (int i = 0; i < size; i++)
            {
                CellRangeAddress merged = sheet.GetMergedRegion(i);

                // remove merged region that overlaps Shifting
                var lastCol=sheet.GetRow(startRow) != null ? sheet.GetRow(startRow).LastCellNum : sheet.GetRow(endRow) != null ? sheet.GetRow(endRow).LastCellNum : 0;
                if (removalNeeded(merged, startRow, endRow, n, lastCol ))
                {
                    removedIndices.Add(i);
                    continue;
                }

                bool inStart = (merged.FirstRow >= startRow || merged.LastRow >= startRow);
                bool inEnd = (merged.FirstRow <= endRow || merged.LastRow <= endRow);

                //don't check if it's not within the Shifted area
                if (!inStart || !inEnd)
                {
                    continue;
                }

                //only shift if the region outside the Shifted rows is not merged too
                if (!merged.ContainsRow(startRow - 1) && !merged.ContainsRow(endRow + 1))
                {
                    merged.FirstRow = merged.FirstRow + n;
                    merged.LastRow =merged.LastRow + n;
                    //have to Remove/add it back
                    ShiftedRegions.Add(merged);
                    removedIndices.Add(i);
                }
            }

            if (removedIndices.Count!=0)
            {
                sheet.RemoveMergedRegions(removedIndices.ToList());
            }

            //read so it doesn't Get Shifted again
            foreach (CellRangeAddress region in ShiftedRegions)
            {
                sheet.AddMergedRegion(region);
            }
            return ShiftedRegions;
        }

        // Keep in sync with {@link ColumnShifter#removalNeeded}
        private bool removalNeeded(CellRangeAddress merged, int startRow, int endRow, int n, int lastCol)
        {
            int movedRows = endRow - startRow + 1;

            // build a range of the rows that are overwritten, i.e. the target-area, but without
            // rows that are moved along
             CellRangeAddress overwrite;
            if (n > 0)
            {
                // area is moved down => overwritten area is [endRow + n - movedRows, endRow + n]
                 int firstRow = Math.Max(endRow + 1, endRow + n - movedRows);
                 int lastRow = endRow + n;
                overwrite = new CellRangeAddress(firstRow, lastRow, 0, lastCol);
            }
            else
            {
                // area is moved up => overwritten area is [startRow + n, startRow + n + movedRows]
                 int firstRow = startRow + n;
                 int lastRow = Math.Min(startRow - 1, startRow + n + movedRows);
                overwrite = new CellRangeAddress(firstRow, lastRow, 0, lastCol);
            }

            // if the merged-region and the overwritten area intersect, we need to remove it
            return merged.Intersects(overwrite);
        }

        /**
 * Verify that the given column indices and step denote a valid range of columns to shift
 *
 * @param firstShiftColumnIndex the column to start shifting
 * @param lastShiftColumnIndex the column to end shifting
 * @param step length of the shifting step
 */
        public static void ValidateShiftParameters(int firstShiftColumnIndex, int lastShiftColumnIndex, int step)
        {
            if (step < 0)
            {
                throw new ArgumentException("Shifting step may not be negative, but had " + step);
            }
            if (firstShiftColumnIndex > lastShiftColumnIndex)
            {
                throw new ArgumentException(String.Format("Incorrect shifting range : %d-%d", firstShiftColumnIndex, lastShiftColumnIndex));
            }
        }

        /**
         * Verify that the given column indices and step denote a valid range of columns to shift to the left
         *
         * @param firstShiftColumnIndex the column to start shifting
         * @param lastShiftColumnIndex the column to end shifting
         * @param step length of the shifting step
         */
        public static void ValidateShiftLeftParameters(int firstShiftColumnIndex, int lastShiftColumnIndex, int step)
        {
            ValidateShiftParameters(firstShiftColumnIndex, lastShiftColumnIndex, step);

            if (firstShiftColumnIndex - step < 0)
            {
                throw new InvalidOperationException("Column index less than zero: " + (firstShiftColumnIndex + step));
            }
        }
        /**
         * Updated named ranges
         */
        public abstract void UpdateNamedRanges(FormulaShifter Shifter);

        /**
         * Update formulas.
         */
        public abstract void UpdateFormulas(FormulaShifter Shifter);

        /**
         * Update the formulas in specified row using the formula Shifting policy specified by Shifter
         *
         * @param row the row to update the formulas on
         * @param Shifter the formula Shifting policy
         */

        public abstract void UpdateRowFormulas(IRow row, FormulaShifter Shifter);

        public abstract void UpdateConditionalFormatting(FormulaShifter Shifter);

        /**
         * Shift the Hyperlink anchors (not the hyperlink text, even if the hyperlink
         * is of type LINK_DOCUMENT and refers to a cell that was Shifted). Hyperlinks
         * do not track the content they point to.
         *
         * @param Shifter the formula Shifting policy
         */
        public abstract void UpdateHyperlinks(FormulaShifter Shifter);

    }

}