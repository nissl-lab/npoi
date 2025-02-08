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
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.SS.UserModel.Helpers
{
    /// <summary>
    /// Helper for Shifting rows up or down
    /// This abstract class exists to consolidate duplicated code between XSSFRowShifter 
    /// and HSSFRowShifter(currently methods sprinkled throughout HSSFSheet)
    /// </summary>
    public abstract class RowShifter
    {
        protected ISheet sheet;

        public RowShifter(ISheet sh)
        {
            sheet = sh;
        }

        /// <summary>
        /// Shifts, grows, or shrinks the merged regions due to a row Shift.
        /// Merged regions that are completely overlaid by Shifting will be deleted.
        /// </summary>
        /// <param name="startRow">the row to start Shifting</param>
        /// <param name="endRow">the row to end Shifting</param>
        /// <param name="n">the number of rows to shift</param>
        /// <returns>an array of affected merged regions, doesn't contain deleted ones</returns>
        public List<CellRangeAddress> ShiftMergedRegions(int startRow, int endRow, int n)
        {
            var ShiftedRegions = new List<CellRangeAddress>();
            ISet<int> removedIndices = new HashSet<int>();
            var size = sheet.NumMergedRegions;

            for (var i = 0; i < size; i++)
            {
                var merged = sheet.GetMergedRegion(i);

                //Shift if the merged region inside the Shifting rows
                if (merged.FirstRow >= startRow && merged.LastRow <= endRow)
                {
                    merged.FirstRow += n;
                    merged.LastRow += n;
                    //have to Remove/add it back
                    ShiftedRegions.Add(merged);
                    removedIndices.Add(i);

                    continue;
                }

                //don't check if it's not within the Shifted area
                if (n > 0)
                {
                    // area is moved down
                    if (merged.LastRow < startRow
                     || merged.FirstRow > endRow + n
                     || merged.FirstRow > endRow && merged.LastRow < startRow + n
                    )
                    {
                        continue;
                    }
                }
                else
                {
                    // area is moved up
                    if (merged.LastRow < startRow + n
                     || merged.FirstRow > endRow
                     || merged.FirstRow > endRow + n && merged.LastRow < startRow
                    )
                    {
                        continue;
                    }
                }

                //remove merged region that overlaps Shifting
                removedIndices.Add(i);
            }

            if (removedIndices.Count != 0)
            {
                sheet.RemoveMergedRegions(removedIndices.ToList());
            }

            //add it which is within the shifted area back
            foreach (var region in ShiftedRegions)
            {
                sheet.AddMergedRegion(region);
            }

            return ShiftedRegions;
        }
        
        /// <summary>
        /// Verify that the given column indices and step denote a valid range of columns to shift
        /// </summary>
        /// <param name="firstShiftColumnIndex">the column to start shifting</param>
        /// <param name="lastShiftColumnIndex">the column to end shifting</param>
        /// <param name="step">length of the shifting step</param>
        /// <exception cref="ArgumentException"></exception>
        public static void ValidateShiftParameters(int firstShiftColumnIndex, int lastShiftColumnIndex, int step)
        {
            if (step < 0)
            {
                throw new ArgumentException("Shifting step may not be negative, but had " + step);
            }

            if (firstShiftColumnIndex > lastShiftColumnIndex)
            {
                throw new ArgumentException(string.Format("Incorrect shifting range : %d-%d", firstShiftColumnIndex, lastShiftColumnIndex));
            }
        }
        
        /// <summary>
        /// Verify that the given column indices and step denote a valid range of columns to shift to the left
        /// </summary>
        /// <param name="firstShiftColumnIndex">the column to start shifting</param>
        /// <param name="lastShiftColumnIndex">the column to end shifting</param>
        /// <param name="step">length of the shifting step</param>
        /// <exception cref="InvalidOperationException"></exception>
        public static void ValidateShiftLeftParameters(int firstShiftColumnIndex, int lastShiftColumnIndex, int step)
        {
            ValidateShiftParameters(firstShiftColumnIndex, lastShiftColumnIndex, step);

            if (firstShiftColumnIndex - step < 0)
            {
                throw new InvalidOperationException("Column index less than zero: " + (firstShiftColumnIndex + step));
            }
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
        /// Update the formulas in specified row using the formula Shifting policy specified by Shifter
        /// </summary>
        /// <param name="row">the row to update the formulas on</param>
        /// <param name="Shifter">the formula Shifting policy</param>
        public abstract void UpdateRowFormulas(IRow row, FormulaShifter Shifter);

        public abstract void UpdateConditionalFormatting(FormulaShifter Shifter);
        
        /// <summary>
        /// Shift the Hyperlink anchors (not the hyperlink text, even if the hyperlink
        /// is of type LINK_DOCUMENT and refers to a cell that was Shifted). Hyperlinks
        /// do not track the content they point to.
        /// </summary>
        /// <param name="Shifter">the formula Shifting policy</param>
        public abstract void UpdateHyperlinks(FormulaShifter Shifter);
    }
}