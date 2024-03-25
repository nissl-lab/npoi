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
    using System.Collections.Generic;
    using NPOI.SS.UserModel;
    using NPOI.Util;


    /// <summary>
    /// Various utility functions that make working with a cells and rows easier. The various methods
    /// that deal with style's allow you to create your CellStyles as you need them. When you apply a
    /// style change to a cell, the code will attempt to see if a style already exists that meets your
    /// needs. If not, then it will create a new style. This is to prevent creating too many styles.
    /// there is an upper limit in Excel on the number of styles that can be supported.
    /// </summary>
    /// @author Eric Pugh epugh@upstate.com
    /// @author (secondary) Avinash Kewalramani akewalramani@accelrys.com
    public class CellUtil
    {
        public const string ALIGNMENT = "alignment";
        public const string BORDER_BOTTOM = "borderBottom";
        public const string BORDER_DIAGONAL = "borderDiagonal";
        public const string BORDER_LEFT = "borderLeft";
        public const string BORDER_RIGHT = "borderRight";
        public const string BORDER_TOP = "borderTop";
        public const string BOTTOM_BORDER_COLOR = "bottomBorderColor";
        public const string DATA_FORMAT = "dataFormat";
        public const string FILL_BACKGROUND_COLOR = "fillBackgroundColor";
        public const string FILL_FOREGROUND_COLOR = "fillForegroundColor";
        public const string FILL_PATTERN = "fillPattern";
        public const string FONT = "font";
        public const string HIDDEN = "hidden";
        public const string INDENTION = "indention";
        public const string LEFT_BORDER_COLOR = "leftBorderColor";
        public const string LOCKED = "locked";
        public const string RIGHT_BORDER_COLOR = "rightBorderColor";
        public const string ROTATION = "rotation";
        public const string SHRINK_TO_FIT = "shrinkToFit";
        public const string TOP_BORDER_COLOR = "topBorderColor";
        public const string VERTICAL_ALIGNMENT = "verticalAlignment";
        public const string WRAP_TEXT = "wrapText";

        private static ISet<String> shortValues = new HashSet<string>(new string[]{
                    BOTTOM_BORDER_COLOR,
                    LEFT_BORDER_COLOR,
                    RIGHT_BORDER_COLOR,
                    TOP_BORDER_COLOR,
                    FILL_FOREGROUND_COLOR,
                    FILL_BACKGROUND_COLOR,
                    INDENTION,
                    DATA_FORMAT,
                    ROTATION
            });
        private static ISet<String> intValues = new HashSet<string>(new string[]{
                        FONT
        });
        private static ISet<String> booleanValues = new HashSet<string>(new string[]{
                        LOCKED,
                        HIDDEN,
                        WRAP_TEXT
        });
        private static ISet<String> borderTypeValues = new HashSet<string>(new string[]{
                        BORDER_BOTTOM,
                        BORDER_LEFT,
                        BORDER_RIGHT,
                        BORDER_TOP
        });


        private static UnicodeMapping[] unicodeMappings;

        private class UnicodeMapping
        {

            public String entityName;
            public String resolvedValue;

            public UnicodeMapping(String pEntityName, String pResolvedValue)
            {
                entityName = "&" + pEntityName + ";";
                resolvedValue = pResolvedValue;
            }
        }

        private CellUtil()
        {
            // no instances of this class
        }

        public static ICell CopyCell(IRow row, int sourceIndex, int targetIndex)
        {
            // Grab a copy of the old/new cell
            ICell oldCell = row.GetCell(sourceIndex);

            // If the old cell is null jump to next cell
            if(oldCell == null)
            {
                return null;
            }

            ICell newCell = row.GetCell(targetIndex);
            if(newCell == null) //not exist
            {
                newCell = row.CreateCell(targetIndex);
            }
            else
            {
                //TODO:shift cells                
            }

            return CopyCell(oldCell, newCell, sourceIndex, targetIndex);
        }

        private static ICell CopyCell(ICell oldCell, ICell newCell, int sourceIndex, int targetIndex)
        {
            if(sourceIndex == targetIndex)
                throw new ArgumentException("sourceIndex and targetIndex cannot be same");

            // Copy style from old cell and apply to new cell
            if(oldCell.CellStyle != null)
            {
                newCell.CellStyle = oldCell.CellStyle;
            }
            // If there is a cell comment, copy
            if(oldCell.CellComment != null)
            {
                newCell.CellComment = oldCell.CellComment;
            }

            // If there is a cell hyperlink, copy
            if(oldCell.Hyperlink != null)
            {
                newCell.Hyperlink = oldCell.Hyperlink;
            }

            // Set the cell data type
            newCell.SetCellType(oldCell.CellType);

            // Set the cell data value
            switch(oldCell.CellType)
            {
                case CellType.Blank:
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;
                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;
                case CellType.Numeric:
                    newCell.SetCellValue(oldCell.NumericCellValue);
                    break;
                case CellType.String:
                    newCell.SetCellValue(oldCell.RichStringCellValue);
                    break;
            }
            return newCell;
        }

        /// <summary>
        /// Get a row from the spreadsheet, and create it if it doesn't exist.
        /// </summary>
        /// <param name="rowIndex">The 0 based row number</param>
        /// <param name="sheet">The sheet that the row is part of.</param>
        /// <return>The row indicated by the rowCounter</return>
        public static IRow GetRow(int rowIndex, ISheet sheet)
        {
            IRow row = sheet.GetRow(rowIndex);
            if(row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }
            return row;
        }

        /// <summary>
        /// Get a specific cell from a row. If the cell doesn't exist, then create it.
        /// </summary>
        /// <param name="row">The row that the cell is part of</param>
        /// <param name="columnIndex">The column index that the cell is in.</param>
        /// <return>The cell indicated by the column.</return>
        public static ICell GetCell(IRow row, int columnIndex)
        {
            ICell cell = row.GetCell(columnIndex);

            if(cell == null)
            {
                cell = row.CreateCell(columnIndex);
            }
            return cell;
        }

        /// <summary>
        /// Creates a cell, gives it a value, and applies a style if provided
        /// </summary>
        /// <param name="row">    the row to create the cell in</param>
        /// <param name="column"> the column index to create the cell in</param>
        /// <param name="value">  The value of the cell</param>
        /// <param name="style">  If the style is not null, then set</param>
        /// <return>A new Cell</return>
        public static ICell CreateCell(IRow row, int column, String value, ICellStyle style)
        {
            ICell cell = GetCell(row, column);

            cell.SetCellValue(cell.Row.Sheet.Workbook.GetCreationHelper()
                    .CreateRichTextString(value));
            if(style != null)
            {
                cell.CellStyle = style;
            }
            return cell;
        }

        /// <summary>
        /// Create a cell, and give it a value.
        /// </summary>
        /// <param name="row">    the row to create the cell in</param>
        /// <param name="column"> the column index to create the cell in</param>
        /// <param name="value">  The value of the cell</param>
        /// <return>A new Cell.</return>
        public static ICell CreateCell(IRow row, int column, String value)
        {
            return CreateCell(row, column, value, null);
        }

        /// <summary>
        /// Take a cell, and align it.
        /// </summary>
        /// <param name="cell">the cell to set the alignment for</param>
        /// <param name="workbook">The workbook that is being worked with.</param>
        /// <param name="align">the column alignment to use.</param>
        /// 
        /// <see cref="CellStyle for alignment options" />
        [Obsolete("deprecated 3.15-beta2. Use {@link #SetAlignment(ICell, HorizontalAlignment)} instead.")]
        public static void SetAlignment(ICell cell, IWorkbook workbook, short align)
        {
            SetCellStyleProperty(cell, workbook, ALIGNMENT, align);
        }
        /// <summary>
        /// <para>
        /// Take a cell, and align it.
        /// </para>
        /// <para>
        /// This is superior to cell.getCellStyle().setAlignment(align) because
        /// this method will not modify the CellStyle object that may be referenced
        /// by multiple cells. Instead, this method will search for existing CellStyles
        /// that match the desired CellStyle, creating a new CellStyle with the desired
        /// style if no match exists.
        /// </para>
        /// </summary>
        /// <param name="cell">the cell to set the alignment for</param>
        /// <param name="align">the horizontal alignment to use.</param>
        /// 
        /// <see cref="HorizontalAlignment for alignment options" />
        /// @since POI 3.15 beta 3
        public static void SetAlignment(ICell cell, HorizontalAlignment align)
        {
            SetCellStyleProperty(cell, ALIGNMENT, align);
        }

        /// <summary>
        /// <para>
        /// Take a cell, and vertically align it.
        /// </para>
        /// <para>
        /// This is superior to cell.getCellStyle().setVerticalAlignment(align) because
        /// this method will not modify the CellStyle object that may be referenced
        /// by multiple cells. Instead, this method will search for existing CellStyles
        /// that match the desired CellStyle, creating a new CellStyle with the desired
        /// style if no match exists.
        /// </para>
        /// </summary>
        /// <param name="cell">the cell to set the alignment for</param>
        /// <param name="align">the vertical alignment to use.</param>
        /// 
        /// <see cref="VerticalAlignment" /> for alignment options
        /// @since POI 3.15 beta 3
        public static void SetVerticalAlignment(ICell cell, VerticalAlignment align)
        {
            SetCellStyleProperty(cell, VERTICAL_ALIGNMENT, align);
        }

        /// <summary>
        /// Take a cell, and apply a font to it
        /// </summary>
        /// <param name="cell">the cell to set the alignment for</param>
        /// <param name="workbook">The workbook that is being worked with.</param>
        /// <param name="font">The Font that you want to set...</param>
        [Obsolete("deprecated 3.15-beta2. Use {@link #SetFont(ICell, IFont)} instead.")]
        public static void SetFont(ICell cell, IWorkbook workbook, IFont font)
        {
            // Check if font belongs to workbook
            short fontIndex = font.Index;
            if(!workbook.GetFontAt(fontIndex).Equals(font))
            {
                throw new ArgumentException("Font does not belong to this workbook");
            }

            // Check if cell belongs to workbook
            // (checked in setCellStyleProperty)

            SetCellStyleProperty(cell, workbook, FONT, fontIndex);
        }
        /// <summary>
        /// Take a cell, and apply a font to it
        /// </summary>
        /// <param name="cell">the cell to set the alignment for</param>
        /// <param name="font">The Font that you want to set.</param>
        /// <exception cref="ArgumentException">if <tt>font</tt> and <tt>cell</tt> do not belong to the same workbook</exception>
        public static void SetFont(ICell cell, IFont font)
        {
            // Check if font belongs to workbook
            IWorkbook wb = cell.Sheet.Workbook;
            short fontIndex = font.Index;
            if(!wb.GetFontAt(fontIndex).Equals(font))
            {
                throw new ArgumentException("Font does not belong to this workbook");
            }

            // Check if cell belongs to workbook
            // (checked in setCellStyleProperty)

            SetCellStyleProperty(cell, FONT, fontIndex);
        }

        /// <summary>
        /// <para>
        /// This method attempts to find an existing CellStyle that matches the <c>cell</c>'s
        /// current style plus styles properties in <c>properties</c>. A new style is created if the
        /// workbook does not contain a matching style.
        /// </para>
        /// <para>
        /// Modifies the cell style of <c>cell</c> without affecting other cells that use the
        /// same style.
        /// </para>
        /// <para>
        /// This is necessary because Excel has an upper limit on the number of styles that it supports.
        /// </para>
        /// <para>
        /// This function is more efficient than multiple calls to
        /// <see cref="SetCellStyleProperty(ICell, IWorkbook, String, Object)" />
        /// if adding multiple cell styles.
        /// </para>
        /// <para>
        /// For performance reasons, if this is the only cell in a workbook that uses a cell style,
        /// this method does NOT remove the old style from the workbook.
        /// <!-- NOT IMPLEMENTED: Unused styles should be
        /// pruned from the workbook with [@link #removeUnusedCellStyles(Workbook)] or
        /// [@link #removeStyleFromWorkbookIfUnused(CellStyle, Workbook)]. -->
        /// </para>
        /// </summary>
        /// <param name="cell">The cell to change the style of</param>
        /// <param name="properties">The properties to be added to a cell style, as {propertyName: propertyValue}.</param>
        /// @since POI 3.14 beta 2
        public static void SetCellStyleProperties(ICell cell, Dictionary<String, Object> properties, bool cloneExistingStyles = false)
        {
            IWorkbook workbook = cell.Sheet.Workbook;
            ICellStyle originalStyle = cell.CellStyle;
            ICellStyle newStyle = null;
            Dictionary<string, object> values = GetFormatProperties(originalStyle);
            PutAll(properties, values);

            // index seems like what index the cellstyle is in the list of styles for a workbook.
            // not good to compare on!
            int numberCellStyles = workbook.NumCellStyles;

            for(int i = 0; i < numberCellStyles; i++)
            {
                ICellStyle wbStyle = workbook.GetCellStyleAt(i);
                Dictionary<string, object> wbStyleMap = GetFormatProperties(wbStyle);

                // the desired style already exists in the workbook. Use the existing style.
                if(DictionaryEqual(wbStyleMap, values, null))
                {
                    newStyle = wbStyle;
                    break;
                }
            }

            // the desired style does not exist in the workbook. Create a new style with desired properties.
            if(newStyle == null)
            {
                newStyle = workbook.CreateCellStyle();
                if(cloneExistingStyles)
                {
                    newStyle.CloneStyleFrom(originalStyle);
                }
                SetFormatProperties(newStyle, workbook, values);
            }

            cell.CellStyle = newStyle;
        }
        public static bool DictionaryEqual<TKey, TValue>(IDictionary<TKey, TValue> first,
            IDictionary<TKey, TValue> second, IEqualityComparer<TValue> valueComparer)
        {
            if(first == second)
                return true;
            if((first == null) || (second == null))
                return false;
            if(first.Count != second.Count)
                return false;

            valueComparer = valueComparer ?? EqualityComparer<TValue>.Default;

            foreach(var kvp in first)
            {
                TValue secondValue;
                if(!second.TryGetValue(kvp.Key, out secondValue))
                    return false;
                if(!valueComparer.Equals(kvp.Value, secondValue))
                    return false;
            }
            return true;
        }
        /// <summary>
        /// <para>
        /// This method attempts to find an existing CellStyle that matches the <c>cell</c>'s
        /// current style plus a single style property <c>propertyName</c> with value
        /// <c>propertyValue</c>.
        /// A new style is created if the workbook does not contain a matching style.
        /// </para>
        /// <para>
        /// 
        /// Modifies the cell style of <c>cell</c> without affecting other cells that use the
        /// same style.
        /// </para>
        /// <para>
        /// 
        /// If setting more than one cell style property on a cell, use
        /// <see cref="SetCellStyleProperties(ICell, Dictionary<string, object>)" />,
        /// which is faster and does not add unnecessary intermediate CellStyles to the workbook.
        /// </para>
        /// </summary>
        /// <param name="cell">The cell that is to be changed.</param>
        /// <param name="propertyName">The name of the property that is to be changed.</param>
        /// <param name="propertyValue">The value of the property that is to be changed.</param>
        public static void SetCellStyleProperty(ICell cell, String propertyName, Object propertyValue)
        {
            Dictionary<String, Object> values = new Dictionary<string, object>() { { propertyName, propertyValue } };
            SetCellStyleProperties(cell, values);
        }
        /// <summary>
        /// <para>
        /// This method attempts to find an existing CellStyle that matches the <c>cell</c>'s
        /// current style plus a single style property <c>propertyName</c> with value
        /// <c>propertyValue</c>.
        /// A new style is created if the workbook does not contain a matching style.
        /// </para>
        /// <para>
        /// 
        /// Modifies the cell style of <c>cell</c> without affecting other cells that use the
        /// same style.
        /// </para>
        /// <para>
        /// 
        /// If setting more than one cell style property on a cell, use
        /// <see cref="SetCellStyleProperties(ICell, Dictionary<string, object>)" />,
        /// which is faster and does not add unnecessary intermediate CellStyles to the workbook.
        /// </para>
        /// </summary>
        /// <param name="workbook">The workbook that is being worked with.</param>
        /// <param name="propertyName">The name of the property that is to be changed.</param>
        /// <param name="propertyValue">The value of the property that is to be changed.</param>
        /// <param name="cell">The cell that needs it's style changes</param>
        [Obsolete("deprecated 3.15-beta2. Use {@link #setCellStyleProperty(Cell, String, Object)} instead.")]
        public static void SetCellStyleProperty(ICell cell, IWorkbook workbook, String propertyName,
               Object propertyValue)
        {
            if(cell.Sheet.Workbook != workbook)
            {
                throw new ArgumentException("Cannot set cell style property. Cell does not belong to workbook.");
            }

            Dictionary<String, Object> values = new Dictionary<string, object>() { { propertyName, propertyValue } };
            SetCellStyleProperties(cell, values);
        }

        /// <summary>
        /// Copies the entries in src to dest, using the preferential data type
        /// so that maps can be compared for equality
        /// </summary>
        /// <param name="src">the property map to copy from (read-only)</param>
        /// <param name="dest">the property map to copy into</param>
        /// @since POI 3.15 beta 3
        private static void PutAll(Dictionary<String, Object> src, Dictionary<String, Object> dest)
        {
            foreach(String key in src.Keys)
            {
                if(shortValues.Contains(key))
                {
                    dest[key] = GetShort(src, key);
                }
                else if(intValues.Contains(key))
                {
                    dest[key] = GetInt(src, key);
                }
                else if(booleanValues.Contains(key))
                {
                    dest[key] = GetBoolean(src, key);
                }
                else if(borderTypeValues.Contains(key))
                {
                    dest[key] = GetBorderStyle(src, key);
                }
                else if(ALIGNMENT.Equals(key))
                {
                    dest[key] = GetHorizontalAlignment(src, key);
                }
                else if(VERTICAL_ALIGNMENT.Equals(key))
                {
                    dest[key] = GetVerticalAlignment(src, key);
                }
                else if(FILL_PATTERN.Equals(key))
                {
                    dest[key] = GetFillPattern(src, key);
                }
                else
                {
                    //if (log.check(POILogger.INFO))
                    //{
                    //    log.log(POILogger.INFO, "Ignoring unrecognized CellUtil format properties key: " + key);
                    //}
                }
            }
        }

        /// <summary>
        /// Returns a map containing the format properties of the given cell style.
        /// The returned map is not tied to <c>style</c>, so subsequent changes
        /// to <c>style</c> will not modify the map, and changes to the returned
        /// map will not modify the cell style. The returned map is mutable.
        /// </summary>
        /// <param name="style">cell style</param>
        /// <return>map of format properties (String -> Object)</return>
        /// <see cref="SetFormatProperties(ICellStyle, IWorkbook, IDictionary)" />
        private static Dictionary<String, Object> GetFormatProperties(ICellStyle style)
        {
            Dictionary<String, Object> properties = new Dictionary<String, Object>();
            Put(properties, ALIGNMENT, style.Alignment);
            Put(properties, VERTICAL_ALIGNMENT, style.VerticalAlignment);
            Put(properties, BORDER_BOTTOM, style.BorderBottom);
            Put(properties, BORDER_LEFT, style.BorderLeft);
            Put(properties, BORDER_RIGHT, style.BorderRight);
            Put(properties, BORDER_TOP, style.BorderTop);
            Put(properties, BOTTOM_BORDER_COLOR, style.BottomBorderColor);
            Put(properties, DATA_FORMAT, style.DataFormat);
            Put(properties, FILL_PATTERN, style.FillPattern);
            Put(properties, FILL_FOREGROUND_COLOR, style.FillForegroundColor);
            Put(properties, FILL_BACKGROUND_COLOR, style.FillBackgroundColor);
            Put(properties, FONT, (int) style.FontIndex);
            Put(properties, HIDDEN, style.IsHidden);
            Put(properties, INDENTION, style.Indention);
            Put(properties, LEFT_BORDER_COLOR, style.LeftBorderColor);
            Put(properties, LOCKED, style.IsLocked);
            Put(properties, RIGHT_BORDER_COLOR, style.RightBorderColor);
            Put(properties, ROTATION, style.Rotation);
            //Put(properties, SHRINK_TO_FIT, style.ShrinkToFit);
            Put(properties, TOP_BORDER_COLOR, style.TopBorderColor);
            Put(properties, WRAP_TEXT, style.WrapText);
            return properties;
        }

        /// <summary>
        /// Sets the format properties of the given style based on the given map.
        /// </summary>
        /// <param name="style">cell style</param>
        /// <param name="workbook">parent workbook</param>
        /// <param name="properties">map of format properties (String -> Object)</param>
        /// <see cref="GetFormatProperties(ICellStyle)" />
        private static void SetFormatProperties(ICellStyle style, IWorkbook workbook, Dictionary<String, Object> properties)
        {
            style.Alignment = GetHorizontalAlignment(properties, ALIGNMENT);
            style.VerticalAlignment = GetVerticalAlignment(properties, VERTICAL_ALIGNMENT);
            style.BorderBottom = GetBorderStyle(properties, BORDER_BOTTOM);
            style.BorderLeft = GetBorderStyle(properties, BORDER_LEFT);
            style.BorderRight = GetBorderStyle(properties, BORDER_RIGHT);
            style.BorderTop = GetBorderStyle(properties, BORDER_TOP);
            style.BottomBorderColor = GetShort(properties, BOTTOM_BORDER_COLOR);
            style.DataFormat = GetShort(properties, DATA_FORMAT);
            style.FillPattern = GetFillPattern(properties, FILL_PATTERN);
            style.FillForegroundColor = GetShort(properties, FILL_FOREGROUND_COLOR);
            style.FillBackgroundColor = GetShort(properties, FILL_BACKGROUND_COLOR);
            style.SetFont(workbook.GetFontAt(GetShort(properties, FONT)));
            style.IsHidden = GetBoolean(properties, HIDDEN);
            style.Indention = GetShort(properties, INDENTION);
            style.LeftBorderColor = GetShort(properties, LEFT_BORDER_COLOR);
            style.IsLocked = GetBoolean(properties, LOCKED);
            style.RightBorderColor = GetShort(properties, RIGHT_BORDER_COLOR);
            style.Rotation = GetShort(properties, ROTATION);
            //style.ShrinkToFit = GetBoolean(properties, SHRINK_TO_FIT);
            style.TopBorderColor = GetShort(properties, TOP_BORDER_COLOR);
            style.WrapText = GetBoolean(properties, WRAP_TEXT);
        }

        /// <summary>
        /// Utility method that returns the named short value form the given map.
        /// </summary>
        /// <param name="properties">map of named properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <return>zero if the property does not exist, or is not a <see cref="short"/>.</return>
        private static short GetShort(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            short result = 0;
            if(short.TryParse(value.ToString(), out result))
                return result;
            return 0;
        }

        /// <summary>
        /// Utility method that returns the named int value from the given map.
        /// </summary>
        /// <param name="properties">map of named properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <return>zero if the property does not exist, or is not a <see cref="Integer"/>
        /// otherwise the property value
        /// </return>
        private static int GetInt(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            if(Number.IsNumber(value))
            {
                return int.Parse(value.ToString());
            }
            return 0;
        }

        /// <summary>
        /// Utility method that returns the named BorderStyle value form the given map.
        /// </summary>
        /// <param name="properties">map of named properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <return>Border style if set, otherwise <see cref="BorderStyle.None" /></return>
        private static BorderStyle GetBorderStyle(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            BorderStyle border;
            if(value is BorderStyle)
            {
                border = (BorderStyle) value;
            }
            // @deprecated 3.15 beta 2. getBorderStyle will only work on BorderStyle enums instead of codes in the future.
            else if(value is short || value is int)
            {
                //if (log.check(POILogger.WARN))
                //{
                //    log.log(POILogger.WARN, "Deprecation warning: CellUtil properties map uses Short values for "
                //            + name + ". Should use BorderStyle enums instead.");
                //}
                short code = short.Parse(value.ToString());
                border = (BorderStyle) code;
            }
            else if(value == null)
            {
                border = BorderStyle.None;
            }
            else
            {
                throw new RuntimeException("Unexpected border style class. Must be BorderStyle or Short (deprecated).");
            }
            return border;
        }

        /// <summary>
        /// Utility method that returns the named FillPattern value from the given map.
        /// </summary>
        /// <param name="properties">map of named properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <return>FillPattern style if set, otherwise <see cref="FillPattern.NoFill" /></return>
        /// @since POI 3.15 beta 3
        private static FillPattern GetFillPattern(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            FillPattern pattern;
            if(value is FillPattern)
            {
                pattern = (FillPattern) value;
            }
            // @deprecated 3.15 beta 2. getFillPattern will only work on FillPattern enums instead of codes in the future.
            else if(value is short)
            {
                //if (log.check(POILogger.WARN))
                //{
                //    log.log(POILogger.WARN, "Deprecation warning: CellUtil properties map uses Short values for "
                //            + name + ". Should use FillPattern enums instead.");
                //}
                short code = (short)value;
                pattern = (FillPattern) code;
            }
            else if(value == null)
            {
                pattern = FillPattern.NoFill;
            }
            else
            {
                throw new RuntimeException("Unexpected fill pattern style class. Must be FillPattern or Short (deprecated).");
            }
            return pattern;
        }

        /// <summary>
        /// Utility method that returns the named HorizontalAlignment value from the given map.
        /// </summary>
        /// <param name="properties">map of named properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <return>HorizontalAlignment style if set, otherwise <see cref="HorizontalAlignment.General" /></return>
        /// @since POI 3.15 beta 3
        private static HorizontalAlignment GetHorizontalAlignment(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            HorizontalAlignment align;
            if(value is HorizontalAlignment)
            {
                align = (HorizontalAlignment) value;
            }
            // @deprecated 3.15 beta 2. getHorizontalAlignment will only work on HorizontalAlignment enums instead of codes in the future.
            else if(value is short)
            {
                //if (log.check(POILogger.WARN))
                //{
                //    log.log(POILogger.WARN, "Deprecation warning: CellUtil properties map used a Short value for "
                //            + name + ". Should use HorizontalAlignment enums instead.");
                //}
                short code = (short)value;
                align = (HorizontalAlignment) code;
            }
            else if(value == null)
            {
                align = HorizontalAlignment.General;
            }
            else
            {
                throw new RuntimeException("Unexpected horizontal alignment style class. Must be HorizontalAlignment or Short (deprecated).");
            }
            return align;
        }

        /// <summary>
        /// Utility method that returns the named VerticalAlignment value from the given map.
        /// </summary>
        /// <param name="properties">map of named properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <return>VerticalAlignment style if set, otherwise <see cref="VerticalAlignment.Bottom" /></return>
        /// @since POI 3.15 beta 3
        private static VerticalAlignment GetVerticalAlignment(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            VerticalAlignment align;
            if(value is VerticalAlignment)
            {
                align = (VerticalAlignment) value;
            }
            // @deprecated 3.15 beta 2. getVerticalAlignment will only work on VerticalAlignment enums instead of codes in the future.
            else if(value is short)
            {
                //if (log.check(POILogger.WARN))
                //{
                //    log.log(POILogger.WARN, "Deprecation warning: CellUtil properties map used a Short value for "
                //            + name + ". Should use VerticalAlignment enums instead.");
                //}
                short code = (short)value;
                align = (VerticalAlignment) code;
            }
            else if(value == null)
            {
                align = VerticalAlignment.Bottom;
            }
            else
            {
                throw new RuntimeException("Unexpected vertical alignment style class. Must be VerticalAlignment or Short (deprecated).");
            }
            return align;
        }

        /// <summary>
        /// Utility method that returns the named boolean value form the given map.
        /// </summary>
        /// <param name="properties">map of properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <return>false if the property does not exist, or is not a <see cref="Boolean"/>.</return>
        private static bool GetBoolean(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            bool result = false;
            if(bool.TryParse(value.ToString(), out result))
                return result;

            return false;
        }
        /// <summary>
        /// Utility method that puts the given value to the given map.
        /// </summary>
        /// <param name="properties">map of properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <param name="value">property value</param>
        private static void Put(Dictionary<String, Object> properties, String name, Object value)
        {
            properties[name] = value;
        }
        /// <summary>
        /// Utility method that puts the named short value to the given map.
        /// </summary>
        /// <param name="properties">map of properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <param name="value">property value</param>
        private static void PutShort(Dictionary<String, Object> properties, String name, short value)
        {
            properties[name] = value;
        }
        /// <summary>
        /// Utility method that puts the named short value to the given map.
        /// </summary>
        /// <param name="properties">map of properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <param name="value">property value</param>
        private static void PutEnum(Dictionary<String, Object> properties, String name, Enum value)
        {
            properties[name] = value;
        }
        /// <summary>
        /// Utility method that puts the named boolean value to the given map.
        /// </summary>
        /// <param name="properties">map of properties (String -> Object)</param>
        /// <param name="name">property name</param>
        /// <param name="value">property value</param>
        private static void PutBoolean(Dictionary<String, Object> properties, String name, bool value)
        {
            properties[name] = value;
        }

        /// <summary>
        ///  Looks for text in the cell that should be unicode, like an alpha and provides the
        ///  unicode version of it.
        /// </summary>
        /// <param name="cell"> The cell to check for unicode values</param>
        /// <return>translated to unicode</return>
        public static ICell TranslateUnicodeValues(ICell cell)
        {
            String s = cell.RichStringCellValue.String;
            bool foundUnicode = false;
            String lowerCaseStr = s.ToLower();

            foreach(UnicodeMapping entry in unicodeMappings)
            {
                String key = entry.entityName;
                if(lowerCaseStr.Contains(key))
                {
                    s = s.Replace(key, entry.resolvedValue);
                    foundUnicode = true;
                }
            }
            if(foundUnicode)
            {
                cell.SetCellValue(cell.Row.Sheet.Workbook.GetCreationHelper()
                        .CreateRichTextString(s));
            }
            return cell;
        }

        static CellUtil()
        {
            unicodeMappings = new UnicodeMapping[] {
            um("alpha",   "\u03B1" ),
            um("beta",    "\u03B2" ),
            um("gamma",   "\u03B3" ),
            um("delta",   "\u03B4" ),
            um("epsilon", "\u03B5" ),
            um("zeta",    "\u03B6" ),
            um("eta",     "\u03B7" ),
            um("theta",   "\u03B8" ),
            um("iota",    "\u03B9" ),
            um("kappa",   "\u03BA" ),
            um("lambda",  "\u03BB" ),
            um("mu",      "\u03BC" ),
            um("nu",      "\u03BD" ),
            um("xi",      "\u03BE" ),
            um("omicron", "\u03BF" ),
        };
        }

        private static UnicodeMapping um(String entityName, String resolvedValue)
        {
            return new UnicodeMapping(entityName, resolvedValue);
        }
    }
}
