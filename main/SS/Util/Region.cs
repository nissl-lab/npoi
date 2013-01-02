/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

using System;

namespace NPOI.SS.Util
{

    /**
     * Represents a from/to row/col square.  This is a object primitive
     * that can be used to represent row,col - row,col just as one would use String
     * to represent a string of characters.  Its really only useful for HSSF though.
     *
     * @author  Andrew C. Oliver acoliver at apache dot org
     */
    [Obsolete]
    public class Region
    {
        private int rowFrom;
        private int colFrom;
        private int rowTo;
        private int colTo;

        /**
         * Creates a new instance of Region (0,0 - 0,0)
         */

        public Region()
        {
        }

        public Region(int rowFrom, int colFrom, int rowTo, int colTo)
        {
            this.rowFrom = rowFrom;
            this.rowTo = rowTo;
            this.colFrom = colFrom;
            this.colTo = colTo;
        }

        /**
         * Get the upper left hand corner column number
         *
         * @return column number for the upper left hand corner
         */

        public int ColumnFrom
        {
            get { return colFrom; }
            set { colFrom = value; }
        }

        /**
         * Get the upper left hand corner row number
         *
         * @return row number for the upper left hand corner
         */

        public int RowFrom
        {
            get { return rowFrom; }
            set { rowFrom = value; }
        }

        /**
         * Get the lower right hand corner column number
         *
         * @return column number for the lower right hand corner
         */

        public int ColumnTo
        {
            get { return colTo; }
            set { colTo = value; }
        }

        /**
         * Get the lower right hand corner row number
         *
         * @return row number for the lower right hand corner
         */

        public int RowTo
        {
            get { return rowTo; }
            set { rowTo = value; }
        }
        private static Region ConvertToRegion(CellRangeAddress cr)
        {

            return new Region(cr.FirstRow, cr.FirstColumn, cr.LastRow, cr.LastColumn);
        }
        /**
         * Convert a List of CellRange objects to an array of regions 
         *  
         * @param List of CellRange objects
         * @return regions
         */
        public static Region[] ConvertCellRangesToRegions(CellRangeAddress[] cellRanges)
        {
            int size = cellRanges.Length;
            if (size < 1)
            {
                return new Region[0];
            }

            Region[] result = new Region[size];

            for (int i = 0; i != size; i++)
            {
                result[i] = ConvertToRegion(cellRanges[i]);
            }
            return result;
        }
        public static CellRangeAddress[] ConvertRegionsToCellRanges(Region[] regions)
        {
            int size = regions.Length;
            if (size < 1)
            {
                return new CellRangeAddress[0];
            }

            CellRangeAddress[] result = new CellRangeAddress[size];

            for (int i = 0; i != size; i++)
            {
                result[i] = ConvertToCellRangeAddress(regions[i]);
            }
            return result;
        }
        public static CellRangeAddress ConvertToCellRangeAddress(Region r)
        {
            return new CellRangeAddress(r.RowFrom, r.RowTo, r.ColumnFrom, r.ColumnTo);
        }
    }
}