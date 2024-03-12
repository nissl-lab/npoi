/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.SS.Util
{
    using System;
    using NPOI.SS.UserModel;

    /**
     * Various utility functions that make working with a region of cells easier.
     *
     * @author Eric Pugh epugh@upstate.com
     * @author (secondary) Avinash Kewalramani akewalramani@accelrys.com
     */
    public class RegionUtil
    {

        private RegionUtil()
        {
            // no instances of this class
        }

        /**
         * For setting the same property on many cells to the same value
         */
        private class CellPropertySetter
        {
            private String _propertyName;
            private object _propertyValue;


            public CellPropertySetter(String propertyName, int value)
            {
                _propertyName = propertyName;
                _propertyValue = (short)value;
            }

            public CellPropertySetter(String propertyName, BorderStyle value)
            {
                _propertyName = propertyName;
                _propertyValue = value;
            }
            public void SetProperty(IRow row, int column)
            {
                // create cell if it does not exist
                ICell cell = CellUtil.GetCell(row, column);
                CellUtil.SetCellStyleProperty(cell, _propertyName, _propertyValue);
            }
        }
        /// <summary>
        /// Sets the left border style for a region of cells by manipulating the cell style of the individual cells on the left
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        [Obsolete("use SetBorderLeft(BorderStyle, CellRangeAddress, ISheet) instead")]
        public static void SetBorderLeft(int border, CellRangeAddress region, ISheet sheet)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.FirstColumn;

            CellPropertySetter cps = new CellPropertySetter(CellUtil.BORDER_LEFT, border);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(CellUtil.GetRow(i, sheet), column);
            }
        }

        /// <summary>
        /// Sets the left border style for a region of cells by manipulating the cell style of the individual cells on the left
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        /// <remarks>since POI 3.16 beta 1</remarks>
        public static void SetBorderLeft(BorderStyle border, CellRangeAddress region, ISheet sheet)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.FirstColumn;

            CellPropertySetter cps = new CellPropertySetter(CellUtil.BORDER_LEFT, border);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(CellUtil.GetRow(i, sheet), column);
            }
        }

        /// <summary>
        /// Sets the left border color for a region of cells by manipulating the cell style of the individual cells on the left
        /// </summary>
        /// <param name="color">The color of the border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        public static void SetLeftBorderColor(int color, CellRangeAddress region, ISheet sheet)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.FirstColumn;

            CellPropertySetter cps = new CellPropertySetter(CellUtil.LEFT_BORDER_COLOR,
                    color);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(CellUtil.GetRow(i, sheet), column);
            }
        }
        /// <summary>
        /// Sets the right border for a region of cells by manipulating the cell style of the individual cells on the right
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        [Obsolete("use SetBorderRight(BorderStyle, CellRangeAddress, ISheet) instead")]
        public static void SetBorderRight(int border, CellRangeAddress region, ISheet sheet)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.LastColumn;

            CellPropertySetter cps = new CellPropertySetter(CellUtil.BORDER_RIGHT, border);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(CellUtil.GetRow(i, sheet), column);
            }
        }

        /// <summary>
        /// Sets the right border style for a region of cells by manipulating the cell style of the individual cells on the right
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        public static void SetBorderRight(BorderStyle border, CellRangeAddress region, ISheet sheet)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.LastColumn;

            CellPropertySetter cps = new CellPropertySetter(CellUtil.BORDER_RIGHT, border);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(CellUtil.GetRow(i, sheet), column);
            }
        }

        /// <summary>
        /// Sets the right border color for a region of cells by manipulating the cell style of the individual cells on the right
        /// </summary>
        /// <param name="color">The color of the border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        public static void SetRightBorderColor(int color, CellRangeAddress region, ISheet sheet)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.LastColumn;

            CellPropertySetter cps = new CellPropertySetter(CellUtil.RIGHT_BORDER_COLOR,
                    color);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(CellUtil.GetRow(i, sheet), column);
            }
        }

        /// <summary>
        /// Sets the bottom border for a region of cells by manipulating the cell style of the individual cells on the bottom
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on</param>
        [Obsolete("use SetBorderBottom(BorderStyle, CellRangeAddress, ISheet) instead")]
        public static void SetBorderBottom(int border, CellRangeAddress region, ISheet sheet)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.LastRow;
            CellPropertySetter cps = new CellPropertySetter(CellUtil.BORDER_BOTTOM, border);
            IRow row = CellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }

        /// <summary>
        /// Sets the bottom border style for a region of cells by manipulating the cell style of the individual cells on the bottom
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on</param>
        public static void SetBorderBottom(BorderStyle border, CellRangeAddress region, ISheet sheet)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.LastRow;
            CellPropertySetter cps = new CellPropertySetter(CellUtil.BORDER_BOTTOM, border);
            IRow row = CellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }

        /// <summary>
        /// Sets the bottom border color for a region of cells by manipulating the cell style of the individual cells on the bottom
        /// </summary>
        /// <param name="color">The color of the border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        public static void SetBottomBorderColor(int color, CellRangeAddress region, ISheet sheet)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.LastRow;
            CellPropertySetter cps = new CellPropertySetter(CellUtil.BOTTOM_BORDER_COLOR,
                    color);
            IRow row = CellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }

        /// <summary>
        /// Sets the top border for a region of cells by manipulating the cell style of the individual cells on the top
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        [Obsolete("use SetBorderTop(BorderStyle, CellRangeAddress, ISheet) instead")]
        public static void SetBorderTop(int border, CellRangeAddress region, ISheet sheet)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.FirstRow;
            CellPropertySetter cps = new CellPropertySetter(CellUtil.BORDER_TOP, border);
            IRow row = CellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }

        /// <summary>
        /// Sets the top border for a region of cells by manipulating the cell style of the individual cells on the top
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        public static void SetBorderTop(BorderStyle border, CellRangeAddress region, ISheet sheet)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.FirstRow;
            CellPropertySetter cps = new CellPropertySetter(CellUtil.BORDER_TOP, border);
            IRow row = CellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }

        /// <summary>
        /// Sets the top border color for a region of cells by manipulating the cell style of the individual cells on the top
        /// </summary>
        /// <param name="color">The color of the border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        public static void SetTopBorderColor(int color, CellRangeAddress region, ISheet sheet)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.FirstRow;
            CellPropertySetter cps = new CellPropertySetter(CellUtil.TOP_BORDER_COLOR, color);
            IRow row = CellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }
    }
}