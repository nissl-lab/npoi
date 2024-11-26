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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SS.Util
{

    using NPOI.HSSF.UserModel;
    using NPOI.SS;
    using NPOI.SS.UserModel;

    /// <summary>
    /// <para>
    /// A <see cref="PropertyTemplate"/> is a template that can be applied to any sheet in
    /// a project. It contains all the border type and color attributes needed to
    /// Draw all the borders for a single sheet. That template can be applied to any
    /// sheet in any workbook.
    /// </para>
    /// <para>
    /// This class requires the full spreadsheet to be in memory, so
    /// <see cref="NPOI.XSSF.Streaming.SXSSFWorkbook} Spreadsheets are not
    /// supported. The same <see cref="PropertyTemplate"/> can, however, be applied to both
    /// <see cref="HSSFWorkbook"/> and {@link NPOI.xssf.UserModel.XSSFWorkbook" />
    /// objects if necessary. Portions of the border that fall outside the max range
    /// of the <see cref="Workbook"/> sheet are ignored.
    /// </para>
    /// <para>
    /// 
    /// This would replace <see cref="RegionUtil"/>.
    /// </para>
    /// </summary>
    public sealed class PropertyTemplate
    {

        /// <summary>
        /// This is a list of cell properties for one shot application to a range of
        /// cells at a later time.
        /// </summary>
        private Dictionary<CellAddress, Dictionary<String, object>> _propertyTemplate;

        /// <summary>
        /// Create a PropertyTemplate object
        /// </summary>
        public PropertyTemplate()
        {
            _propertyTemplate = new Dictionary<CellAddress, Dictionary<String, object>>();
        }

        /// <summary>
        /// Create a PropertyTemplate object from another PropertyTemplate
        /// </summary>
        /// <param name="template">a PropertyTemplate object</param>
        public PropertyTemplate(PropertyTemplate template) : this()
        {
            foreach(KeyValuePair<CellAddress, Dictionary<String, object>> entry in template.Template)
            {
                _propertyTemplate[new CellAddress(entry.Key)] = cloneCellProperties(entry.Value);
            }
        }

        private Dictionary<CellAddress, Dictionary<String, object>> Template
        {
            get
            {
                return _propertyTemplate;
            }
            
        }

        private static Dictionary<String, object> cloneCellProperties(Dictionary<String, object> properties)
        {
            Dictionary<String, object> newProperties = new Dictionary<String, object>();
            foreach(KeyValuePair<String, object> entry in properties)
            {
                newProperties[entry.Key] = entry.Value;
            }
            return newProperties;
        }

        /// <summary>
        /// Draws a group of cell borders for a cell range. The borders are not
        /// applied to the cells at this time, just the template is Drawn. To apply
        /// the Drawn borders to a sheet, use <see cref="applyBorders"/>.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders to be
        /// applied.
        /// </param>
        public void DrawBorders(CellRangeAddress range, BorderStyle borderType,
                BorderExtent extent)
        {
            switch(extent)
            {
                case BorderExtent.NONE:
                    removeBorders(range);
                    break;
                case BorderExtent.ALL:
                    DrawHorizontalBorders(range, borderType, BorderExtent.ALL);
                    DrawVerticalBorders(range, borderType, BorderExtent.ALL);
                    break;
                case BorderExtent.INSIDE:
                    DrawHorizontalBorders(range, borderType, BorderExtent.INSIDE);
                    DrawVerticalBorders(range, borderType, BorderExtent.INSIDE);
                    break;
                case BorderExtent.OUTSIDE:
                    DrawOutsideBorders(range, borderType, BorderExtent.ALL);
                    break;
                case BorderExtent.TOP:
                    DrawTopBorder(range, borderType);
                    break;
                case BorderExtent.BOTTOM:
                    DrawBottomBorder(range, borderType);
                    break;
                case BorderExtent.LEFT:
                    DrawLeftBorder(range, borderType);
                    break;
                case BorderExtent.RIGHT:
                    DrawRightBorder(range, borderType);
                    break;
                case BorderExtent.HORIZONTAL:
                    DrawHorizontalBorders(range, borderType, BorderExtent.ALL);
                    break;
                case BorderExtent.INSIDE_HORIZONTAL:
                    DrawHorizontalBorders(range, borderType, BorderExtent.INSIDE);
                    break;
                case BorderExtent.OUTSIDE_HORIZONTAL:
                    DrawOutsideBorders(range, borderType, BorderExtent.HORIZONTAL);
                    break;
                case BorderExtent.VERTICAL:
                    DrawVerticalBorders(range, borderType, BorderExtent.ALL);
                    break;
                case BorderExtent.INSIDE_VERTICAL:
                    DrawVerticalBorders(range, borderType, BorderExtent.INSIDE);
                    break;
                case BorderExtent.OUTSIDE_VERTICAL:
                    DrawOutsideBorders(range, borderType, BorderExtent.VERTICAL);
                    break;
            }
        }

        /// <summary>
        /// Draws a group of cell borders for a cell range. The borders are not
        /// applied to the cells at this time, just the template is Drawn. To apply
        /// the Drawn borders to a sheet, use <see cref="applyBorders"/>.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders to be
        /// applied.
        /// </param>
        public void DrawBorders(CellRangeAddress range, BorderStyle borderType,
                short color, BorderExtent extent)
        {
            DrawBorders(range, borderType, extent);
            if(borderType != BorderStyle.None)
            {
                DrawBorderColors(range, color, extent);
            }
        }

        /// <summary>
        /// Draws the top border for a range of cells
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        private void DrawTopBorder(CellRangeAddress range, BorderStyle borderType)
        {
            int row = range.FirstRow;
            int firstCol = range.FirstColumn;
            int lastCol = range.LastColumn;
            for(int i = firstCol; i <= lastCol; i++)
            {
                AddProperty(row, i, CellUtil.BORDER_TOP, borderType);
                if(borderType == BorderStyle.None && row > 0)
                {
                    AddProperty(row - 1, i, CellUtil.BORDER_BOTTOM, borderType);
                }
            }
        }

        /// <summary>
        /// Draws the bottom border for a range of cells
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        private void DrawBottomBorder(CellRangeAddress range,
                BorderStyle borderType)
        {
            int row = range.LastRow;
            int firstCol = range.FirstColumn;
            int lastCol = range.LastColumn;
            for(int i = firstCol; i <= lastCol; i++)
            {
                AddProperty(row, i, CellUtil.BORDER_BOTTOM, borderType);
                if(borderType == BorderStyle.None
                        && row < SpreadsheetVersion.EXCEL2007.MaxRows - 1)
                {
                    AddProperty(row + 1, i, CellUtil.BORDER_TOP, borderType);
                }
            }
        }

        /// <summary>
        /// Draws the left border for a range of cells
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        private void DrawLeftBorder(CellRangeAddress range,
                BorderStyle borderType)
        {
            int firstRow = range.FirstRow;
            int lastRow = range.LastRow;
            int col = range.FirstColumn;
            for(int i = firstRow; i <= lastRow; i++)
            {
                AddProperty(i, col, CellUtil.BORDER_LEFT, borderType);
                if(borderType == BorderStyle.None && col > 0)
                {
                    AddProperty(i, col - 1, CellUtil.BORDER_RIGHT, borderType);
                }
            }
        }

        /// <summary>
        /// Draws the right border for a range of cells
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        private void DrawRightBorder(CellRangeAddress range,
                BorderStyle borderType)
        {
            int firstRow = range.FirstRow;
            int lastRow = range.LastRow;
            int col = range.LastColumn;
            for(int i = firstRow; i <= lastRow; i++)
            {
                AddProperty(i, col, CellUtil.BORDER_RIGHT, borderType);
                if(borderType == BorderStyle.None
                        && col < SpreadsheetVersion.EXCEL2007.MaxColumns - 1)
                {
                    AddProperty(i, col + 1, CellUtil.BORDER_LEFT, borderType);
                }
            }
        }

        /// <summary>
        /// Draws the outside borders for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders to be
        /// applied. Valid Values are:
        /// <list type="bullet">
        /// <item><description>BorderExtent.ALL</li></description></item>
        /// <item><description>BorderExtent.HORIZONTAL</li></description></item>
        /// <item><description>BorderExtent.VERTICAL</li></description></item>
        /// </list>
        /// </param>
        private void DrawOutsideBorders(CellRangeAddress range,
                BorderStyle borderType, BorderExtent extent)
        {
            switch(extent)
            {
                case BorderExtent.ALL:
                case BorderExtent.HORIZONTAL:
                case BorderExtent.VERTICAL:
                    if(extent == BorderExtent.ALL || extent == BorderExtent.HORIZONTAL)
                    {
                        DrawTopBorder(range, borderType);
                        DrawBottomBorder(range, borderType);
                    }
                    if(extent == BorderExtent.ALL || extent == BorderExtent.VERTICAL)
                    {
                        DrawLeftBorder(range, borderType);
                        DrawRightBorder(range, borderType);
                    }
                    break;
                default:
                    throw new ArgumentException(
                            "Unsupported PropertyTemplate.Extent, valid Extents are ALL, HORIZONTAL, and VERTICAL");
            }
        }

        /// <summary>
        /// Draws the horizontal borders for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders to be
        /// applied. Valid Values are:
        /// <list type="bullet">
        /// <item><description>BorderExtent.ALL</li></description></item>
        /// <item><description>BorderExtent.INSIDE</li></description></item>
        /// </list>
        /// </param>
        private void DrawHorizontalBorders(CellRangeAddress range,
                BorderStyle borderType, BorderExtent extent)
        {
            switch(extent)
            {
                case BorderExtent.ALL:
                case BorderExtent.INSIDE:
                    int firstRow = range.FirstRow;
                    int lastRow = range.LastRow;
                    int firstCol = range.FirstColumn;
                    int lastCol = range.LastColumn;
                    for(int i = firstRow; i <= lastRow; i++)
                    {
                        CellRangeAddress row = new CellRangeAddress(i, i, firstCol,
                        lastCol);
                        if(extent == BorderExtent.ALL || i > firstRow)
                        {
                            DrawTopBorder(row, borderType);
                        }
                        if(extent == BorderExtent.ALL || i < lastRow)
                        {
                            DrawBottomBorder(row, borderType);
                        }
                    }
                    break;
                default:
                    throw new ArgumentException(
                            "Unsupported PropertyTemplate.Extent, valid Extents are ALL and INSIDE");
            }
        }

        /// <summary>
        /// Draws the vertical borders for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which borders are
        /// Drawn.
        /// </param>
        /// <param name="borderType">borderType
        /// - Type of border to Draw. <see cref="BorderStyle"/>.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders to be
        /// applied. Valid Values are:
        /// <list type="bullet">
        /// <item><description>BorderExtent.ALL</li></description></item>
        /// <item><description>BorderExtent.INSIDE</li></description></item>
        /// </list>
        /// </param>
        private void DrawVerticalBorders(CellRangeAddress range,
                BorderStyle borderType, BorderExtent extent)
        {
            switch(extent)
            {
                case BorderExtent.ALL:
                case BorderExtent.INSIDE:
                    int firstRow = range.FirstRow;
                    int lastRow = range.LastRow;
                    int firstCol = range.FirstColumn;
                    int lastCol = range.LastColumn;
                    for(int i = firstCol; i <= lastCol; i++)
                    {
                        CellRangeAddress row = new CellRangeAddress(firstRow, lastRow,
                        i, i);
                        if(extent == BorderExtent.ALL || i > firstCol)
                        {
                            DrawLeftBorder(row, borderType);
                        }
                        if(extent == BorderExtent.ALL || i < lastCol)
                        {
                            DrawRightBorder(row, borderType);
                        }
                    }
                    break;
                default:
                    throw new ArgumentException(
                            "Unsupported PropertyTemplate.Extent, valid Extents are ALL and INSIDE");
            }
        }

        /// <summary>
        /// Removes all border properties from this <see cref="PropertyTemplate"/> for the
        /// specified range.
        /// </summary>
        /// @parm range - <see cref="CellRangeAddress"/> range of cells to remove borders.
        private void removeBorders(CellRangeAddress range)
        {
            HashSet<String> properties = new HashSet<String>();
            properties.Add(CellUtil.BORDER_TOP);
            properties.Add(CellUtil.BORDER_BOTTOM);
            properties.Add(CellUtil.BORDER_LEFT);
            properties.Add(CellUtil.BORDER_RIGHT);
            for(int row = range.FirstRow; row <= range.LastRow; row++)
            {
                for(int col = range.FirstColumn; col <= range
                        .LastColumn; col++)
                {
                    removeProperties(row, col, properties);
                }
            }
            removeBorderColors(range);
        }

        /// <summary>
        /// Applies the Drawn borders to a Sheet. The borders that are applied are
        /// the ones that have been Drawn by the <see cref="drawBorders"/> and
        /// <see cref="drawBorderColors"/> methods.
        /// </summary>
        /// <param name="sheet">sheet
        /// - <see cref="Sheet"/> on which to apply borders
        /// </param>
        public void applyBorders(ISheet sheet)
        {
            IWorkbook wb = sheet.Workbook;
            foreach(KeyValuePair<CellAddress, Dictionary<String, object>> entry in _propertyTemplate)
            {
                CellAddress cellAddress = entry.Key;
                if(cellAddress.Row < wb.SpreadsheetVersion.MaxRows
                        && cellAddress.Column < wb.SpreadsheetVersion
                                .MaxColumns)
                {
                    Dictionary<String, object> properties = entry.Value;
                    IRow row = CellUtil.GetRow(cellAddress.Row, sheet);
                    ICell cell = CellUtil.GetCell(row, cellAddress.Column);
                    CellUtil.SetCellStyleProperties(cell, properties);
                }
            }
        }

        /// <summary>
        /// Sets the color for a group of cell borders for a cell range. The borders
        /// are not applied to the cells at this time, just the template is Drawn. If
        /// the borders do not exist, a BORDER_THIN border is used. To apply the
        /// Drawn borders to a sheet, use <see cref="applyBorders"/>.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which colors are
        /// Set.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders for which
        /// colors are Set.
        /// </param>
        public void DrawBorderColors(CellRangeAddress range, short color,
                BorderExtent extent)
        {
            switch(extent)
            {
                case BorderExtent.NONE:
                    removeBorderColors(range);
                    break;
                case BorderExtent.ALL:
                    DrawHorizontalBorderColors(range, color, BorderExtent.ALL);
                    DrawVerticalBorderColors(range, color, BorderExtent.ALL);
                    break;
                case BorderExtent.INSIDE:
                    DrawHorizontalBorderColors(range, color, BorderExtent.INSIDE);
                    DrawVerticalBorderColors(range, color, BorderExtent.INSIDE);
                    break;
                case BorderExtent.OUTSIDE:
                    DrawOutsideBorderColors(range, color, BorderExtent.ALL);
                    break;
                case BorderExtent.TOP:
                    DrawTopBorderColor(range, color);
                    break;
                case BorderExtent.BOTTOM:
                    DrawBottomBorderColor(range, color);
                    break;
                case BorderExtent.LEFT:
                    DrawLeftBorderColor(range, color);
                    break;
                case BorderExtent.RIGHT:
                    DrawRightBorderColor(range, color);
                    break;
                case BorderExtent.HORIZONTAL:
                    DrawHorizontalBorderColors(range, color, BorderExtent.ALL);
                    break;
                case BorderExtent.INSIDE_HORIZONTAL:
                    DrawHorizontalBorderColors(range, color, BorderExtent.INSIDE);
                    break;
                case BorderExtent.OUTSIDE_HORIZONTAL:
                    DrawOutsideBorderColors(range, color, BorderExtent.HORIZONTAL);
                    break;
                case BorderExtent.VERTICAL:
                    DrawVerticalBorderColors(range, color, BorderExtent.ALL);
                    break;
                case BorderExtent.INSIDE_VERTICAL:
                    DrawVerticalBorderColors(range, color, BorderExtent.INSIDE);
                    break;
                case BorderExtent.OUTSIDE_VERTICAL:
                    DrawOutsideBorderColors(range, color, BorderExtent.VERTICAL);
                    break;
            }
        }

        /// <summary>
        /// Sets the color of the top border for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which colors are
        /// Set.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        private void DrawTopBorderColor(CellRangeAddress range, short color)
        {
            int row = range.FirstRow;
            int firstCol = range.FirstColumn;
            int lastCol = range.LastColumn;
            for(int i = firstCol; i <= lastCol; i++)
            {
                if(GetBorderStyle(row, i,
                        CellUtil.BORDER_TOP) == BorderStyle.None)
                {
                    DrawTopBorder(new CellRangeAddress(row, row, i, i),
                            BorderStyle.Thin);
                }
                addProperty(row, i, CellUtil.TOP_BORDER_COLOR, color);
            }
        }

        /// <summary>
        /// Sets the color of the bottom border for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which colors are
        /// Set.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        private void DrawBottomBorderColor(CellRangeAddress range, short color)
        {
            int row = range.LastRow;
            int firstCol = range.FirstColumn;
            int lastCol = range.LastColumn;
            for(int i = firstCol; i <= lastCol; i++)
            {
                if(GetBorderStyle(row, i,
                        CellUtil.BORDER_BOTTOM) == BorderStyle.None)
                {
                    DrawBottomBorder(new CellRangeAddress(row, row, i, i),
                            BorderStyle.Thin);
                }
                addProperty(row, i, CellUtil.BOTTOM_BORDER_COLOR, color);
            }
        }

        /// <summary>
        /// Sets the color of the left border for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which colors are
        /// Set.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        private void DrawLeftBorderColor(CellRangeAddress range, short color)
        {
            int firstRow = range.FirstRow;
            int lastRow = range.LastRow;
            int col = range.FirstColumn;
            for(int i = firstRow; i <= lastRow; i++)
            {
                if(GetBorderStyle(i, col,
                        CellUtil.BORDER_LEFT) == BorderStyle.None)
                {
                    DrawLeftBorder(new CellRangeAddress(i, i, col, col),
                            BorderStyle.Thin);
                }
                addProperty(i, col, CellUtil.LEFT_BORDER_COLOR, color);
            }
        }

        /// <summary>
        /// Sets the color of the right border for a range of cells. If the border is
        /// not Drawn, it defaults to BORDER_THIN
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which colors are
        /// Set.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        private void DrawRightBorderColor(CellRangeAddress range, short color)
        {
            int firstRow = range.FirstRow;
            int lastRow = range.LastRow;
            int col = range.LastColumn;
            for(int i = firstRow; i <= lastRow; i++)
            {
                if(GetBorderStyle(i, col,
                        CellUtil.BORDER_RIGHT) == BorderStyle.None)
                {
                    DrawRightBorder(new CellRangeAddress(i, i, col, col),
                            BorderStyle.Thin);
                }
                addProperty(i, col, CellUtil.RIGHT_BORDER_COLOR, color);
            }
        }

        /// <summary>
        /// Sets the color of the outside borders for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which colors are
        /// Set.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders for which
        /// colors are Set. Valid Values are:
        /// <list type="bullet">
        /// <item><description>BorderExtent.ALL</li></description></item>
        /// <item><description>BorderExtent.HORIZONTAL</li></description></item>
        /// <item><description>BorderExtent.VERTICAL</li></description></item>
        /// </list>
        /// </param>
        private void DrawOutsideBorderColors(CellRangeAddress range, short color,
                BorderExtent extent)
        {
            switch(extent)
            {
                case BorderExtent.ALL:
                case BorderExtent.HORIZONTAL:
                case BorderExtent.VERTICAL:
                    if(extent == BorderExtent.ALL || extent == BorderExtent.HORIZONTAL)
                    {
                        DrawTopBorderColor(range, color);
                        DrawBottomBorderColor(range, color);
                    }
                    if(extent == BorderExtent.ALL || extent == BorderExtent.VERTICAL)
                    {
                        DrawLeftBorderColor(range, color);
                        DrawRightBorderColor(range, color);
                    }
                    break;
                default:
                    throw new ArgumentException(
                            "Unsupported PropertyTemplate.Extent, valid Extents are ALL, HORIZONTAL, and VERTICAL");
            }
        }

        /// <summary>
        /// Sets the color of the horizontal borders for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which colors are
        /// Set.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders for which
        /// colors are Set. Valid Values are:
        /// <list type="bullet">
        /// <item><description>BorderExtent.ALL</li></description></item>
        /// <item><description>BorderExtent.INSIDE</li></description></item>
        /// </list>
        /// </param>
        private void DrawHorizontalBorderColors(CellRangeAddress range, short color,
                BorderExtent extent)
        {
            switch(extent)
            {
                case BorderExtent.ALL:
                case BorderExtent.INSIDE:
                    int firstRow = range.FirstRow;
                    int lastRow = range.LastRow;
                    int firstCol = range.FirstColumn;
                    int lastCol = range.LastColumn;
                    for(int i = firstRow; i <= lastRow; i++)
                    {
                        CellRangeAddress row = new CellRangeAddress(i, i, firstCol,
                        lastCol);
                        if(extent == BorderExtent.ALL || i > firstRow)
                        {
                            DrawTopBorderColor(row, color);
                        }
                        if(extent == BorderExtent.ALL || i < lastRow)
                        {
                            DrawBottomBorderColor(row, color);
                        }
                    }
                    break;
                default:
                    throw new ArgumentException(
                            "Unsupported PropertyTemplate.Extent, valid Extents are ALL and INSIDE");
            }
        }

        /// <summary>
        /// Sets the color of the vertical borders for a range of cells.
        /// </summary>
        /// <param name="range">range
        /// - <see cref="CellRangeAddress"/> range of cells on which colors are
        /// Set.
        /// </param>
        /// <param name="color">color
        /// - Color index from <see cref="IndexedColors"/> used to Draw the
        /// borders.
        /// </param>
        /// <param name="extent">extent
        /// - <see cref="BorderExtent"/> of the borders for which
        /// colors are Set. Valid Values are:
        /// <list type="bullet">
        /// <item><description>BorderExtent.ALL</li></description></item>
        /// <item><description>BorderExtent.INSIDE</li></description></item>
        /// </list>
        /// </param>
        private void DrawVerticalBorderColors(CellRangeAddress range, short color,
                BorderExtent extent)
        {
            switch(extent)
            {
                case BorderExtent.ALL:
                case BorderExtent.INSIDE:
                    int firstRow = range.FirstRow;
                    int lastRow = range.LastRow;
                    int firstCol = range.FirstColumn;
                    int lastCol = range.LastColumn;
                    for(int i = firstCol; i <= lastCol; i++)
                    {
                        CellRangeAddress row = new CellRangeAddress(firstRow, lastRow,
                        i, i);
                        if(extent == BorderExtent.ALL || i > firstCol)
                        {
                            DrawLeftBorderColor(row, color);
                        }
                        if(extent == BorderExtent.ALL || i < lastCol)
                        {
                            DrawRightBorderColor(row, color);
                        }
                    }
                    break;
                default:
                    throw new ArgumentException(
                            "Unsupported PropertyTemplate.Extent, valid Extents are ALL and INSIDE");
            }
        }

        /// <summary>
        /// Removes all border properties from this <see cref="PropertyTemplate"/> for the
        /// specified range.
        /// </summary>
        /// @parm range - <see cref="CellRangeAddress"/> range of cells to remove borders.
        private void removeBorderColors(CellRangeAddress range)
        {
            HashSet<String> properties = new HashSet<String>();
            properties.Add(CellUtil.TOP_BORDER_COLOR);
            properties.Add(CellUtil.BOTTOM_BORDER_COLOR);
            properties.Add(CellUtil.LEFT_BORDER_COLOR);
            properties.Add(CellUtil.RIGHT_BORDER_COLOR);
            for(int row = range.FirstRow; row <= range.LastRow; row++)
            {
                for(int col = range.FirstColumn; col <= range
                        .LastColumn; col++)
                {
                    removeProperties(row, col, properties);
                }
            }
        }

        /// <summary>
        /// Adds a property to this <see cref="PropertyTemplate"/> for a given cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <param name="property">property</param>
        /// <param name="value">value</param>
        private void addProperty(int row, int col, String property, short value)
        {
            AddProperty(row, col, property, (object)value);
        }

        /// <summary>
        /// Adds a property to this <see cref="PropertyTemplate"/> for a given cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <param name="property">property</param>
        /// <param name="value">value</param>
        private void AddProperty(int row, int col, String property, object value)
        {
            CellAddress cell = new CellAddress(row, col);
            Dictionary<String, object> cellProperties = _propertyTemplate.ContainsKey(cell) ? _propertyTemplate[cell] : null;
            if(cellProperties == null)
            {
                cellProperties = new Dictionary<String, object>();
            }
            cellProperties[property] = value;
            _propertyTemplate[cell] = cellProperties;
        }

        /// <summary>
        /// Removes a Set of properties from this <see cref="PropertyTemplate"/> for a
        /// given cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <param name="properties">properties</param>
        private void removeProperties(int row, int col, HashSet<String> properties)
        {
            CellAddress cell = new CellAddress(row, col);
            Dictionary<String, object> cellProperties = _propertyTemplate.ContainsKey(cell) ? _propertyTemplate[cell] : null;
            if(cellProperties != null)
            {
                //cellProperties.keySet().removeAll(properties);
                foreach(string p in properties)
                {
                    cellProperties.Remove(p);
                }
                
                
                if(cellProperties.Count == 0)
                {
                    _propertyTemplate.Remove(cell);
                }
                else
                {
                    _propertyTemplate[cell] = cellProperties;
                }
            }
        }

        /// <summary>
        /// Retrieves the number of borders assigned to a cell
        /// </summary>
        /// <param name="cell">cell</param>
        public int GetNumBorders(CellAddress cell)
        {
            Dictionary<String, object> cellProperties = _propertyTemplate.ContainsKey(cell) ? _propertyTemplate[cell] : null;
            if(cellProperties == null)
            {
                return 0;
            }

            int count = 0;
            foreach(String property in cellProperties.Keys)
            {
                if(property.Equals(CellUtil.BORDER_TOP))
                    count += 1;
                if(property.Equals(CellUtil.BORDER_BOTTOM))
                    count += 1;
                if(property.Equals(CellUtil.BORDER_LEFT))
                    count += 1;
                if(property.Equals(CellUtil.BORDER_RIGHT))
                    count += 1;
            }
            return count;
        }

        /// <summary>
        /// Retrieves the number of borders assigned to a cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        public int GetNumBorders(int row, int col)
        {
            return GetNumBorders(new CellAddress(row, col));
        }

        /// <summary>
        /// Retrieves the number of border colors assigned to a cell
        /// </summary>
        /// <param name="cell">cell</param>
        public int GetNumBorderColors(CellAddress cell)
        {
            Dictionary<String, object> cellProperties = _propertyTemplate.ContainsKey(cell) ? _propertyTemplate[cell] : null;
            if(cellProperties == null)
            {
                return 0;
            }

            int count = 0;
            foreach(String property in cellProperties.Keys)
            {
                if(property.Equals(CellUtil.TOP_BORDER_COLOR))
                    count += 1;
                if(property.Equals(CellUtil.BOTTOM_BORDER_COLOR))
                    count += 1;
                if(property.Equals(CellUtil.LEFT_BORDER_COLOR))
                    count += 1;
                if(property.Equals(CellUtil.RIGHT_BORDER_COLOR))
                    count += 1;
            }
            return count;
        }

        /// <summary>
        /// Retrieves the number of border colors assigned to a cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        public int GetNumBorderColors(int row, int col)
        {
            return GetNumBorderColors(new CellAddress(row, col));
        }

        /// <summary>
        /// Retrieves the border style for a given cell
        /// </summary>
        /// <param name="cell">cell</param>
        /// <param name="property">property</param>
        public BorderStyle GetBorderStyle(CellAddress cell, String property)
        {
            BorderStyle value = BorderStyle.None;
            Dictionary<String, object> cellProperties = _propertyTemplate.ContainsKey(cell) ? _propertyTemplate[cell] : null;
            if(cellProperties != null)
            {
                object obj = cellProperties.ContainsKey(property) ? cellProperties[property] : null;
                if(obj is BorderStyle)
                {
                    value = (BorderStyle) obj;
                }
            }
            return value;
        }

        /// <summary>
        /// Retrieves the border style for a given cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <param name="property">property</param>
        public BorderStyle GetBorderStyle(int row, int col, String property)
        {
            return GetBorderStyle(new CellAddress(row, col), property);
        }

        /// <summary>
        /// Retrieves the border style for a given cell
        /// </summary>
        /// <param name="cell">cell</param>
        /// <param name="property">property</param>
        public short GetTemplateProperty(CellAddress cell, String property)
        {
            short value = 0;
            Dictionary<String, object> cellProperties = _propertyTemplate.ContainsKey(cell) ? _propertyTemplate[cell] : null;
            if(cellProperties != null)
            {
                object obj = cellProperties.ContainsKey(property) ? cellProperties[property] : null;
                if(obj != null)
                {
                    value = Getshort(obj);
                }
            }
            return value;
        }

        /// <summary>
        /// Retrieves the border style for a given cell
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        /// <param name="property">property</param>
        public short GetTemplateProperty(int row, int col, String property)
        {
            return GetTemplateProperty(new CellAddress(row, col), property);
        }

        /// <summary>
        /// Converts a short object to a short value or 0 if the object is not a
        /// short
        /// </summary>
        /// <param name="value">Potentially short value to convert</param>
        /// <return>short value, or 0 if not a short</return>
        private static short Getshort(object value)
        {
            if(value is short)
            {
                return (short) value;
            }
            return 0;
        }
    }
}

