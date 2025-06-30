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
            RegionUtil.SetBorderLeft(border, region, sheet);
        }

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
            RegionUtil.SetLeftBorderColor(color, region, sheet);
        }

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
            RegionUtil.SetBorderRight(border, region, sheet);
        }


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
            RegionUtil.SetRightBorderColor(color, region, sheet);
        }

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
            RegionUtil.SetBorderBottom(border, region, sheet);
        }


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
            RegionUtil.SetBottomBorderColor(color, region, sheet);
        }


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
            RegionUtil.SetBorderTop(border, region, sheet);
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
            RegionUtil.SetTopBorderColor(color, region, sheet);
        }
    }
}