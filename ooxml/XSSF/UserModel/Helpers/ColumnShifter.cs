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

using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.XSSF.UserModel.Helpers
{
    /// <summary>
    /// Helper for Shifting columns left or right
    /// This abstract class exists to consolidate duplicated code between 
    /// XSSFColumnShifter and HSSFColumnShifter
    /// </summary>
    public abstract class ColumnShifter
    {
        protected XSSFSheet sheet;

        public ColumnShifter(XSSFSheet sh)
        {
            sheet = sh;
        }

        /// <summary>
        /// Shifts, grows, or shrinks the merged regions due to a column Shift.
        /// Merged regions that are completely overlaid by Shifting will 
        /// be deleted.
        /// </summary>
        /// <param name="startColumn">the column to start Shifting</param>
        /// <param name="endColumn">the column to end Shifting</param>
        /// <param name="n">the number of columns to shift</param>
        /// <returns>an array of affected merged regions, doesn't contain 
        /// deleted ones</returns>
        public List<CellRangeAddress> ShiftMergedRegions(
            int startColumn,
            int endColumn,
            int n)
        {
            List<CellRangeAddress> shiftedRegions = new List<CellRangeAddress>();
            ISet<int> removedIndices = new HashSet<int>();
            //move merged regions completely if they fall within the new region
            //boundaries when they are Shifted
            int size = sheet.NumMergedRegions;

            for (int i = 0; i < size; i++)
            {
                CellRangeAddress merged = sheet.GetMergedRegion(i);

                // remove merged region that overlaps Shifting
                if (RemovalNeeded(merged, startColumn, endColumn, n))
                {
                    _ = removedIndices.Add(i);
                    continue;
                }

                bool inStart = merged.FirstColumn >= startColumn
                    || merged.LastColumn >= startColumn;
                bool inEnd = merged.FirstColumn <= endColumn ||
                    merged.LastColumn <= endColumn;

                //don't check if it's not within the Shifted area
                if (!inStart || !inEnd)
                {
                    continue;
                }

                //only shift if the region outside the Shifted columns is not
                //merged too
                if (!merged.ContainsColumn(startColumn - 1)
                    && !merged.ContainsColumn(endColumn + 1))
                {
                    merged.FirstColumn += n;
                    merged.LastColumn += n;
                    //have to Remove/add it back
                    shiftedRegions.Add(merged);
                    _ = removedIndices.Add(i);
                }
            }

            if (removedIndices.Count != 0)
            {
                sheet.RemoveMergedRegions(removedIndices.ToList());
            }

            //read so it doesn't Get Shifted again
            foreach (CellRangeAddress region in shiftedRegions)
            {
                _ = sheet.AddMergedRegion(region);
            }

            return shiftedRegions;
        }

        // Keep in sync with {@link ColumnShifter#removalNeeded}
        private bool RemovalNeeded(
            CellRangeAddress merged,
            int startColumn,
            int endColumn,
            int n)
        {
            int firstColumn = startColumn + n;
            int lastColumn = endColumn + n;
            CellRangeAddress overwrite =
                new CellRangeAddress(0, sheet.LastRowNum, firstColumn, lastColumn);

            // if the merged-region and the overwritten area intersect,
            // we need to remove it
            return merged.Intersects(overwrite);
        }

        /// <summary>
        /// Updated named ranges
        /// </summary>
        /// <param name="Shifter"></param>
        public abstract void UpdateNamedRanges(FormulaShifter Shifter);

        /// <summary>
        /// Update formulas.
        /// </summary>
        /// <param name="Shifter"></param>
        public abstract void UpdateFormulas(FormulaShifter Shifter);

        /// <summary>
        /// Update the formulas in specified column using the formula Shifting 
        /// policy specified by Shifter
        /// </summary>
        /// <param name="column">the column to update the formulas on</param>
        /// <param name="Shifter">the formula Shifting policy</param>
        public abstract void UpdateColumnFormulas(IColumn column, FormulaShifter Shifter);

        public abstract void UpdateConditionalFormatting(FormulaShifter Shifter);

        /// <summary>
        /// Shift the Hyperlink anchors (not the hyperlink text, even if the 
        /// hyperlink is of type LINK_DOCUMENT and refers to a cell that was 
        /// Shifted). Hyperlinks do not track the content they point to.
        /// </summary>
        /// <param name="Shifter">the formula Shifting policy</param>
        public abstract void UpdateHyperlinks(FormulaShifter Shifter);
    }
}
