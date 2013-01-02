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

namespace NPOI.HSSF.Util
{
    using System;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.Util;

    /// <summary>
    /// Various utility functions that make working with a region of cells easier.
    /// @author Eric Pugh epugh@upstate.com
    /// </summary>
    public class HSSFRegionUtil
    {

        private HSSFRegionUtil()
        {
            // no instances of this class
        }
        /// <summary>
        /// For setting the same property on many cells to the same value
        /// </summary>
        private class CellPropertySetter
        {

            private HSSFWorkbook _workbook;
            private String _propertyName;
            private short _propertyValue;

            public CellPropertySetter(HSSFWorkbook workbook, String propertyName, int value)
            {
                _workbook = workbook;
                _propertyName = propertyName;
                _propertyValue = (short)value;
            }
            public void SetProperty(NPOI.SS.UserModel.IRow row, int column)
            {
                NPOI.SS.UserModel.ICell cell = HSSFCellUtil.GetCell(row, column);
                HSSFCellUtil.SetCellStyleProperty(cell, _workbook, _propertyName, _propertyValue);
            }
        }
        [Obsolete]
        private static CellRangeAddress toCRA(Region region)
        {
            return Region.ConvertToCellRangeAddress(region);
        }

        //[Obsolete]
        //public static void SetBorderLeft(NPOI.SS.UserModel.CellBorderType border, Region region, HSSFSheet sheet,
        //        HSSFWorkbook workbook)
        //{
        //    SetBorderLeft(border, toCRA(region), sheet, workbook);
        //}
        /// <summary>
        /// Sets the left border for a region of cells by manipulating the cell style
        /// of the individual cells on the left
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        /// <param name="workbook">The workbook that the region is on.</param>
        public static void SetBorderLeft(NPOI.SS.UserModel.BorderStyle border, CellRangeAddress region, HSSFSheet sheet,
                HSSFWorkbook workbook)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.FirstColumn;

            CellPropertySetter cps = new CellPropertySetter(workbook, HSSFCellUtil.BORDER_LEFT, (int)border);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(HSSFCellUtil.GetRow(i, sheet), column);
            }
        }

        //[Obsolete]
        //public static void SetLeftBorderColor(short color, Region region, HSSFSheet sheet,
        //        HSSFWorkbook workbook)
        //{
        //    SetLeftBorderColor(color, toCRA(region), sheet, workbook);
        //}
        /// <summary>
        /// Sets the leftBorderColor attribute of the HSSFRegionUtil object
        /// </summary>
        /// <param name="color">The color of the border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        /// <param name="workbook">The workbook that the region is on.</param>
        public static void SetLeftBorderColor(int color, CellRangeAddress region, HSSFSheet sheet,
                HSSFWorkbook workbook)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.FirstColumn;

            CellPropertySetter cps = new CellPropertySetter(workbook, HSSFCellUtil.LEFT_BORDER_COLOR, color);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(HSSFCellUtil.GetRow(i, sheet), column);
            }
        }

        //[Obsolete]
        //public static void SetBorderRight(NPOI.SS.UserModel.CellBorderType border, Region region, HSSFSheet sheet,
        //        HSSFWorkbook workbook)
        //{
        //    SetBorderRight(border, toCRA(region), sheet, workbook);
        //}
        /// <summary>
        /// Sets the borderRight attribute of the HSSFRegionUtil object
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        /// <param name="workbook">The workbook that the region is on.</param>
        public static void SetBorderRight(NPOI.SS.UserModel.BorderStyle border, CellRangeAddress region, HSSFSheet sheet,
                HSSFWorkbook workbook)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.LastColumn;

            CellPropertySetter cps = new CellPropertySetter(workbook, HSSFCellUtil.BORDER_RIGHT, (int)border);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(HSSFCellUtil.GetRow(i, sheet), column);
            }
        }

        //[Obsolete]
        //public static void SetRightBorderColor(short color, Region region, HSSFSheet sheet,
        //        HSSFWorkbook workbook)
        //{
        //    SetRightBorderColor(color, toCRA(region), sheet, workbook);
        //}
        /// <summary>
        /// Sets the rightBorderColor attribute of the HSSFRegionUtil object
        /// </summary>
        /// <param name="color">The color of the border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The workbook that the region is on.</param>
        /// <param name="workbook">The sheet that the region is on.</param>
        public static void SetRightBorderColor(int color, CellRangeAddress region, HSSFSheet sheet,
                HSSFWorkbook workbook)
        {
            int rowStart = region.FirstRow;
            int rowEnd = region.LastRow;
            int column = region.LastColumn;

            CellPropertySetter cps = new CellPropertySetter(workbook, HSSFCellUtil.RIGHT_BORDER_COLOR, color);
            for (int i = rowStart; i <= rowEnd; i++)
            {
                cps.SetProperty(HSSFCellUtil.GetRow(i, sheet), column);
            }
        }

        //[Obsolete]
        //public static void SetBorderBottom(NPOI.SS.UserModel.CellBorderType border, Region region, HSSFSheet sheet,
        //        HSSFWorkbook workbook)
        //{
        //    SetBorderBottom(border, toCRA(region), sheet, workbook);
        //}
        /// <summary>
        /// Sets the borderBottom attribute of the HSSFRegionUtil object
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        /// <param name="workbook">The workbook that the region is on.</param>
        public static void SetBorderBottom(NPOI.SS.UserModel.BorderStyle border, CellRangeAddress region, HSSFSheet sheet,
                HSSFWorkbook workbook)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.LastRow;
            CellPropertySetter cps = new CellPropertySetter(workbook, HSSFCellUtil.BORDER_BOTTOM, (int)border);
            NPOI.SS.UserModel.IRow row = HSSFCellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }

        //[Obsolete]
        //public static void SetBottomBorderColor(short color, Region region, HSSFSheet sheet,
        //        HSSFWorkbook workbook)
        //{
        //    SetBottomBorderColor(color, toCRA(region), sheet, workbook);
        //}
        /// <summary>
        /// Sets the bottomBorderColor attribute of the HSSFRegionUtil object
        /// </summary>
        /// <param name="color">The color of the border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        /// <param name="workbook">The workbook that the region is on.</param>
        public static void SetBottomBorderColor(int color, CellRangeAddress region, HSSFSheet sheet,
                HSSFWorkbook workbook)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.LastRow;
            CellPropertySetter cps = new CellPropertySetter(workbook, HSSFCellUtil.BOTTOM_BORDER_COLOR, color);
            NPOI.SS.UserModel.IRow row = HSSFCellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }

        //[Obsolete]
        //public static void SetBorderTop(NPOI.SS.UserModel.CellBorderType border, Region region, HSSFSheet sheet,
        //        HSSFWorkbook workbook)
        //{
        //    SetBorderTop(border, toCRA(region), sheet, workbook);
        //}
        /// <summary>
        /// Sets the borderBottom attribute of the HSSFRegionUtil object
        /// </summary>
        /// <param name="border">The new border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        /// <param name="workbook">The workbook that the region is on.</param>
        public static void SetBorderTop(NPOI.SS.UserModel.BorderStyle border, CellRangeAddress region, HSSFSheet sheet,
                HSSFWorkbook workbook)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.FirstRow;
            CellPropertySetter cps = new CellPropertySetter(workbook, HSSFCellUtil.BORDER_TOP, (int)border);
            NPOI.SS.UserModel.IRow row = HSSFCellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }

        /// <summary>
        /// Sets the topBorderColor attribute of the HSSFRegionUtil object
        /// </summary>
        /// <param name="color">The color of the border</param>
        /// <param name="region">The region that should have the border</param>
        /// <param name="sheet">The sheet that the region is on.</param>
        /// <param name="workbook">The workbook that the region is on.</param>
        public static void SetTopBorderColor(int color, CellRangeAddress region, HSSFSheet sheet,
                HSSFWorkbook workbook)
        {
            int colStart = region.FirstColumn;
            int colEnd = region.LastColumn;
            int rowIndex = region.FirstRow;
            CellPropertySetter cps = new CellPropertySetter(workbook, HSSFCellUtil.TOP_BORDER_COLOR, color);
            NPOI.SS.UserModel.IRow row = HSSFCellUtil.GetRow(rowIndex, sheet);
            for (int i = colStart; i <= colEnd; i++)
            {
                cps.SetProperty(row, i);
            }
        }
    }
}