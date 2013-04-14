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

namespace NPOI.SS.Util.CellWalk
{

    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    /**
     * Traverse cell range.
     *
     * @author Roman Kashitsyn
     */
    public class CellWalk
    {

        private ISheet sheet;
        private CellRangeAddress range;
        private bool traverseEmptyCells;

        public CellWalk(ISheet sheet, CellRangeAddress range)
        {
            this.sheet = sheet;
            this.range = range;
            this.traverseEmptyCells = false;
        }

        /**
         * Should we call handler on empty (blank) cells. Default is
         * false.
         * @return true if handler should be called on empty (blank)
         * cells, false otherwise.
         */
        public bool IsTraverseEmptyCells()
        {
            return traverseEmptyCells;
        }

        /**
         * Sets the traverseEmptyCells property.
         * @param traverseEmptyCells new property value
         */
        public void SetTraverseEmptyCells(bool traverseEmptyCells)
        {
            this.traverseEmptyCells = traverseEmptyCells;
        }

        /**
         * Traverse cell range from top left to bottom right cell.
         * @param handler handler to call on each appropriate cell
         */
        public void Traverse(ICellHandler handler)
        {
            int firstRow = range.FirstRow;
            int lastRow = range.LastRow;
            int firstColumn = range.FirstColumn;
            int lastColumn = range.LastColumn;
            int width = lastColumn - firstColumn + 1;
            SimpleCellWalkContext ctx = new SimpleCellWalkContext();
            IRow currentRow = null;
            ICell currentCell = null;

            for (ctx.rowNumber = firstRow; ctx.rowNumber <= lastRow; ++ctx.rowNumber)
            {
                currentRow = sheet.GetRow(ctx.rowNumber);
                if (currentRow == null)
                {
                    continue;
                }
                for (ctx.colNumber = firstColumn; ctx.colNumber <= lastColumn; ++ctx.colNumber)
                {
                    currentCell = currentRow.GetCell(ctx.colNumber);

                    if (currentCell == null)
                    {
                        continue;
                    }
                    if (IsEmpty(currentCell) && !traverseEmptyCells)
                    {
                        continue;
                    }

                    ctx.ordinalNumber =
                        (ctx.rowNumber - firstRow) * width +
                        (ctx.colNumber - firstColumn + 1);

                    handler.OnCell(currentCell, ctx);
                }
            }
        }

        private bool IsEmpty(ICell cell)
        {
            return (cell.CellType == CellType.Blank);
        }

        /**
         * Inner class to hold walk context.
         * @author Roman Kashitsyn
         */
        private class SimpleCellWalkContext : ICellWalkContext
        {
            public long ordinalNumber = 0;
            public int rowNumber = 0;
            public int colNumber = 0;

            public long OrdinalNumber
            {
                get
                {
                    return ordinalNumber;
                }
            }

            public int RowNumber
            {
                get
                {
                    return rowNumber;
                }
            }

            public int ColumnNumber
            {
                get
                {
                    return colNumber;
                }
            }
        }
    }

}