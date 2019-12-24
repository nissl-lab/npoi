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
            private short _propertyValue;


            public CellPropertySetter(String propertyName, int value)
            {
                _propertyName = propertyName;
                _propertyValue = (short)value;
            }


            public void SetProperty(IRow row, int column)
            {
                ICell cell = CellUtil.GetCell(row, column);
                CellUtil.SetCellStyleProperty(cell, _propertyName, _propertyValue);
            }
        }

        /**
         * Sets the left border for a region of cells by manipulating the cell style of the individual
         * cells on the left
         *
         * @param border The new border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
        [Obsolete("deprecated 3.15 beta 2. Use {@link #setBorderLeft(int, CellRangeAddress, Sheet).")]
        public static void SetBorderLeft(int border, CellRangeAddress region, ISheet sheet,
                IWorkbook workbook)
        {
            SetBorderLeft(border, region, sheet);
        }
        /**
         * Sets the left border for a region of cells by manipulating the cell style of the individual
         * cells on the left
         * 
         * @see {@link BorderPropertyTemplate#drawBorders} for improved border functionality
         *
         * @param border The new border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
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

        /**
         * Sets the leftBorderColor attribute of the RegionUtil object
         *
         * @param color The color of the border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
        [Obsolete("deprecated 3.15 beta 2. Use {@link #setLeftBorderColor(int, CellRangeAddress, Sheet).")]
        public static void SetLeftBorderColor(int color, CellRangeAddress region, ISheet sheet,
                IWorkbook workbook)
        {
            SetLeftBorderColor(color, region, sheet);
        }
        /**
         * Sets the left border color for a region of cells by manipulating the cell style of the individual
         * cells on the left
         * 
         * @see {@link BorderPropertyTemplate#drawBorderColors} for improved border functionality
         *
         * @param color The color of the border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
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

        /**
         * Sets the borderRight attribute of the RegionUtil object
         *
         * @param border The new border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
        [Obsolete("deprecated 3.15 beta 2. Use {@link #setBorderRight(int, CellRangeAddress, Sheet).")]
        public static void SetBorderRight(int border, CellRangeAddress region, ISheet sheet,
                IWorkbook workbook)
        {
            SetBorderRight(border, region, sheet);
        }
        /**
         * Sets the right border for a region of cells by manipulating the cell style of the individual
         * cells on the right
         * 
         * @see {@link BorderPropertyTemplate#drawBorders} for improved border functionality
         *
         * @param border The new border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
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

        /**
         * Sets the rightBorderColor attribute of the RegionUtil object
         *
         * @param color The color of the border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
        [Obsolete("deprecated 3.15 beta 2. Use {@link #setRightBorderColor(int, CellRangeAddress, Sheet).")]
        public static void SetRightBorderColor(int color, CellRangeAddress region, ISheet sheet,
                IWorkbook workbook)
        {
            SetRightBorderColor(color, region, sheet);
        }
        /**
         * Sets the right border color for a region of cells by manipulating the cell style of the individual
         * cells on the right
         * 
         * @see {@link BorderPropertyTemplate#drawBorderColors} for improved border functionality
         *
         * @param color The color of the border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
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

        /**
         * Sets the borderBottom attribute of the RegionUtil object
         *
         * @param border The new border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
        [Obsolete("deprecated 3.15 beta 2. Use {@link #setBorderBottom(int, CellRangeAddress, Sheet).")]
        public static void SetBorderBottom(int border, CellRangeAddress region, ISheet sheet,
                IWorkbook workbook)
        {
            SetBorderBottom(border, region, sheet);
        }
        /**
         * Sets the bottom border for a region of cells by manipulating the cell style of the individual
         * cells on the bottom
         * 
         * @see {@link BorderPropertyTemplate#drawBorders} for improved border functionality
         *
         * @param border The new border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
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

        /**
         * Sets the bottomBorderColor attribute of the RegionUtil object
         *
         * @param color The color of the border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
        [Obsolete("deprecated 3.15 beta 2. Use {@link #setBottomBorderColor(int, CellRangeAddress, Sheet).")]
        public static void SetBottomBorderColor(int color, CellRangeAddress region, ISheet sheet,
                IWorkbook workbook)
        {
            SetBottomBorderColor(color, region, sheet);
        }
        /**
         * Sets the bottom border color for a region of cells by manipulating the cell style of the individual
         * cells on the bottom
         * 
         * @see {@link BorderPropertyTemplate#drawBorderColors} for improved border functionality
         *
         * @param color The color of the border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
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

        /**
         * Sets the borderBottom attribute of the RegionUtil object
         *
         * @param border The new border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
         [Obsolete("deprecated 3.15 beta 2. Use {@link #setBorderTop(int, CellRangeAddress, Sheet).")]
        public static void SetBorderTop(int border, CellRangeAddress region, ISheet sheet,
                IWorkbook workbook)
        {
            SetBorderTop(border, region, sheet);
        }
        /**
         * Sets the top border for a region of cells by manipulating the cell style of the individual
         * cells on the top
         * 
         * @see {@link BorderPropertyTemplate#drawBorders} for improved border functionality
         *
         * @param border The new border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
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

        /**
         * Sets the topBorderColor attribute of the RegionUtil object
         *
         * @param color The color of the border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
         [Obsolete("deprecated 3.15 beta 2. Use {@link #setTopBorderColor(int, CellRangeAddress, Sheet)}.")]
        public static void SetTopBorderColor(int color, CellRangeAddress region, ISheet sheet,
                IWorkbook workbook)
        {
            SetTopBorderColor(color, region, sheet);
        }
        /**
         * Sets the top border color for a region of cells by manipulating the cell style of the individual
         * cells on the top
         * 
         * @see {@link BorderPropertyTemplate#drawBorderColors} for improved border functionality
         *
         * @param color The color of the border
         * @param region The region that should have the border
         * @param workbook The workbook that the region is on.
         * @param sheet The sheet that the region is on.
         */
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