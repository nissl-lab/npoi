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
            var rangesToShift = new List<CellRangeAddress>();
            var rangesToExpand = new List<CellRangeAddress>();
            var size = sheet.NumMergedRegions;

            for (var i = 0; i < size; i++)
            {
                var merged = sheet.GetMergedRegion(i);
                if (MergedInsideMovingRange(startRow, endRow, n, merged))
                {
                    rangesToShift.Add(merged);
                    continue;
                }

                if (MovingRangeCrossesMerged(startRow, endRow, n, merged))
                {
                    rangesToExpand.Add(merged);
                }
            }

            foreach (var range in rangesToShift)
            {
                range.FirstRow = range.FirstRow + n;
                range.LastRow = range.LastRow + n;
            }

            foreach (var range in rangesToExpand)
            {
                range.LastRow = range.LastRow + n;

            }

            return rangesToShift;
        }

        private bool MergedInsideMovingRange(int startRow, int endRow, int n, CellRangeAddress merged)
        {
            return startRow <= merged.FirstRow && endRow >= merged.LastRow;
        }

        private bool MovingRangeCrossesMerged(int startRow, int endRow, int n, CellRangeAddress merged)
        {
            return merged.FirstRow <= startRow && startRow <= merged.LastRow;
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