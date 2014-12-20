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


    /**
     * Various utility functions that make working with a cells and rows easier. The various methods
     * that deal with style's allow you to create your CellStyles as you need them. When you apply a
     * style change to a cell, the code will attempt to see if a style already exists that meets your
     * needs. If not, then it will create a new style. This is to prevent creating too many styles.
     * there is an upper limit in Excel on the number of styles that can be supported.
     *
     *@author Eric Pugh epugh@upstate.com
     *@author (secondary) Avinash Kewalramani akewalramani@accelrys.com
     */
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
        public const string DIAGONAL_BORDER_COLOR = "diagonalBorderColor";
        public const string DIAGONAL_BORDER_LINE_STYLE = "diagonalBorderLineStyle";
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
            if (sourceIndex == targetIndex)
                throw new ArgumentException("sourceIndex and targetIndex cannot be same");
            // Grab a copy of the old/new cell
            ICell oldCell = row.GetCell(sourceIndex);

            // If the old cell is null jump to next cell
            if (oldCell == null)
            {
                return null;
            }

            ICell newCell = row.GetCell(targetIndex);
            if (newCell == null) //not exist
            {
                newCell = row.CreateCell(targetIndex);
            }
            else
            {
                //TODO:shift cells                
            }
            // Copy style from old cell and apply to new cell
            if (oldCell.CellStyle != null)
            {
                newCell.CellStyle = oldCell.CellStyle;
            }
            // If there is a cell comment, copy
            if (oldCell.CellComment != null)
            {
                newCell.CellComment = oldCell.CellComment;
            }

            // If there is a cell hyperlink, copy
            if (oldCell.Hyperlink != null)
            {
                newCell.Hyperlink = oldCell.Hyperlink;
            }

            // Set the cell data type
            newCell.SetCellType(oldCell.CellType);

            // Set the cell data value
            switch (oldCell.CellType)
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
        /**
         * Get a row from the spreadsheet, and create it if it doesn't exist.
         *
         *@param rowIndex The 0 based row number
         *@param sheet The sheet that the row is part of.
         *@return The row indicated by the rowCounter
         */
        public static IRow GetRow(int rowIndex, ISheet sheet)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }
            return row;
        }

        /**
         * Get a specific cell from a row. If the cell doesn't exist, then create it.
         *
         *@param row The row that the cell is part of
         *@param columnIndex The column index that the cell is in.
         *@return The cell indicated by the column.
         */
        public static ICell GetCell(IRow row, int columnIndex)
        {
            ICell cell = row.GetCell(columnIndex);

            if (cell == null)
            {
                cell = row.CreateCell(columnIndex);
            }
            return cell;
        }

        /**
         * Creates a cell, gives it a value, and applies a style if provided
         *
         * @param  row     the row to create the cell in
         * @param  column  the column index to create the cell in
         * @param  value   The value of the cell
         * @param  style   If the style is not null, then set
         * @return         A new Cell
         */
        public static ICell CreateCell(IRow row, int column, String value, ICellStyle style)
        {
            ICell cell = GetCell(row, column);

            cell.SetCellValue(cell.Row.Sheet.Workbook.GetCreationHelper()
                    .CreateRichTextString(value));
            if (style != null)
            {
                cell.CellStyle = style;
            }
            return cell;
        }

        /**
         * Create a cell, and give it a value.
         *
         *@param  row     the row to create the cell in
         *@param  column  the column index to create the cell in
         *@param  value   The value of the cell
         *@return         A new Cell.
         */
        public static ICell CreateCell(IRow row, int column, String value)
        {
            return CreateCell(row, column, value, null);
        }

        /**
         * Take a cell, and align it.
         *
         *@param cell the cell to set the alignment for
         *@param workbook The workbook that is being worked with.
         *@param align the column alignment to use.
         *
         * @see CellStyle for alignment options
         */
        public static void SetAlignment(ICell cell, IWorkbook workbook, short align)
        {
            SetCellStyleProperty(cell, workbook, ALIGNMENT, align);
        }

        /**
         * Take a cell, and apply a font to it
         *
         *@param cell the cell to set the alignment for
         *@param workbook The workbook that is being worked with.
         *@param font The Font that you want to set...
         */
        public static void SetFont(ICell cell, IWorkbook workbook, IFont font)
        {
            SetCellStyleProperty(cell, workbook, FONT, font.Index);
        }

        /**
         * This method attempt to find an already existing CellStyle that matches what you want the
         * style to be. If it does not find the style, then it creates a new one. If it does create a
         * new one, then it applies the propertyName and propertyValue to the style. This is necessary
         * because Excel has an upper limit on the number of Styles that it supports.
         *
         *@param workbook The workbook that is being worked with.
         *@param propertyName The name of the property that is to be changed.
         *@param propertyValue The value of the property that is to be changed.
         *@param cell The cell that needs it's style changes
         */
        public static void SetCellStyleProperty(ICell cell, IWorkbook workbook, String propertyName, Object propertyValue)
        {
            ICellStyle originalStyle = cell.CellStyle;
            ICellStyle newStyle = null;
            Dictionary<String, Object> values = GetFormatProperties(originalStyle);
            if (values.ContainsKey(propertyName))
                values[propertyName] = propertyValue;
            else
                values.Add(propertyName, propertyValue);

            // index seems like what index the cellstyle is in the list of styles for a workbook.
            // not good to compare on!
            short numberCellStyles = workbook.NumCellStyles;

            for (short i = 0; i < numberCellStyles; i++)
            {
                ICellStyle wbStyle = workbook.GetCellStyleAt(i);

                Dictionary<String, Object> wbStyleMap = GetFormatProperties(wbStyle);

                if (values.Keys.Count != wbStyleMap.Keys.Count) continue;

                bool found = true;
                
                foreach (string key in values.Keys)
                {
                    if (!wbStyleMap.ContainsKey(key))
                    {
                        found = false;
                        break;
                    }

                    if (values[key].Equals(wbStyleMap[key])) continue;

                    found = false;
                    break;
                }

                if (found)
                {
                    newStyle = wbStyle;
                    break;
                }
            }

            if (newStyle == null)
            {
                newStyle = workbook.CreateCellStyle();
                SetFormatProperties(newStyle, workbook, values);
            }

            cell.CellStyle = newStyle;
        }

        /**
         * Returns a map containing the format properties of the given cell style.
         *
         * @param style cell style
         * @return map of format properties (String -> Object)
         * @see #setFormatProperties(org.apache.poi.ss.usermodel.CellStyle, org.apache.poi.ss.usermodel.Workbook, java.util.Map)
         */
        private static Dictionary<String, Object> GetFormatProperties(ICellStyle style)
        {
            Dictionary<String, Object> properties = new Dictionary<String, Object>();
            PutShort(properties, ALIGNMENT, (short)style.Alignment);
            PutShort(properties, BORDER_BOTTOM, (short)style.BorderBottom);
            PutShort(properties, BORDER_DIAGONAL, (short)style.BorderDiagonal);
            PutShort(properties, BORDER_LEFT, (short)style.BorderLeft);
            PutShort(properties, BORDER_RIGHT, (short)style.BorderRight);
            PutShort(properties, BORDER_TOP, (short)style.BorderTop);
            PutShort(properties, BOTTOM_BORDER_COLOR, style.BottomBorderColor);
            PutShort(properties, DATA_FORMAT, style.DataFormat);
            PutShort(properties, DIAGONAL_BORDER_COLOR, style.BorderDiagonalColor);
            PutShort(properties, DIAGONAL_BORDER_LINE_STYLE, (short)style.BorderDiagonalLineStyle);
            PutShort(properties, FILL_BACKGROUND_COLOR, style.FillBackgroundColor);
            PutShort(properties, FILL_FOREGROUND_COLOR, style.FillForegroundColor);
            PutShort(properties, FILL_PATTERN, (short)style.FillPattern);
            PutShort(properties, FONT, style.FontIndex);
            PutBoolean(properties, HIDDEN, style.IsHidden);
            PutShort(properties, INDENTION, style.Indention);
            PutShort(properties, LEFT_BORDER_COLOR, style.LeftBorderColor);
            PutBoolean(properties, LOCKED, style.IsLocked);
            PutShort(properties, RIGHT_BORDER_COLOR, style.RightBorderColor);
            PutShort(properties, ROTATION, style.Rotation);
            PutBoolean(properties, SHRINK_TO_FIT, style.ShrinkToFit);
            PutShort(properties, TOP_BORDER_COLOR, style.TopBorderColor);
            PutShort(properties, VERTICAL_ALIGNMENT, (short)style.VerticalAlignment);
            PutBoolean(properties, WRAP_TEXT, style.WrapText);
            return properties;
        }

        /**
         * Sets the format properties of the given style based on the given map.
         *
         * @param style cell style
         * @param workbook parent workbook
         * @param properties map of format properties (String -> Object)
         * @see #getFormatProperties(CellStyle)
         */
        private static void SetFormatProperties(ICellStyle style, IWorkbook workbook, Dictionary<String, Object> properties)
        {
            style.Alignment = (HorizontalAlignment)GetShort(properties, ALIGNMENT);
            style.BorderBottom = (BorderStyle)GetShort(properties, BORDER_BOTTOM);
            style.BorderDiagonalColor = GetShort(properties, DIAGONAL_BORDER_COLOR);
            style.BorderDiagonal = (BorderDiagonal)GetShort(properties, BORDER_DIAGONAL);
            style.BorderDiagonalLineStyle = (BorderStyle)GetShort(properties, DIAGONAL_BORDER_LINE_STYLE);
            style.BorderLeft = (BorderStyle)GetShort(properties, BORDER_LEFT);
            style.BorderRight = (BorderStyle)GetShort(properties, BORDER_RIGHT);
            style.BorderTop = (BorderStyle)GetShort(properties, BORDER_TOP);
            style.BottomBorderColor = GetShort(properties, BOTTOM_BORDER_COLOR);
            style.DataFormat =GetShort(properties, DATA_FORMAT);
            style.FillBackgroundColor = GetShort(properties, FILL_BACKGROUND_COLOR);
            style.FillForegroundColor = GetShort(properties, FILL_FOREGROUND_COLOR);
            style.FillPattern = (FillPattern)GetShort(properties, FILL_PATTERN);
            style.SetFont(workbook.GetFontAt(GetShort(properties, FONT)));
            style.IsHidden = GetBoolean(properties, HIDDEN);
            style.Indention = GetShort(properties, INDENTION);
            style.LeftBorderColor = GetShort(properties, LEFT_BORDER_COLOR);
            style.IsLocked = GetBoolean(properties, LOCKED);
            style.RightBorderColor = GetShort(properties, RIGHT_BORDER_COLOR);
            style.Rotation = GetShort(properties, ROTATION);
            style.ShrinkToFit = GetBoolean(properties, SHRINK_TO_FIT);
            style.TopBorderColor = GetShort(properties, TOP_BORDER_COLOR);
            style.VerticalAlignment = (VerticalAlignment)GetShort(properties, VERTICAL_ALIGNMENT);
            style.WrapText = GetBoolean(properties, WRAP_TEXT);
        }

        /**
         * Utility method that returns the named short value form the given map.
         * @return zero if the property does not exist, or is not a {@link Short}.
         *
         * @param properties map of named properties (String -> Object)
         * @param name property name
         * @return property value, or zero
         */
        private static short GetShort(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            short result = 0;
            if (short.TryParse(value.ToString(), out result))
                return result;
            return 0;
        }

        /**
         * Utility method that returns the named boolean value form the given map.
         * @return false if the property does not exist, or is not a {@link Boolean}.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @return property value, or false
         */
        private static bool GetBoolean(Dictionary<String, Object> properties, String name)
        {
            Object value = properties[name];
            bool result = false;
            if (bool.TryParse(value.ToString(), out result))
                return result;

            return false;
        }

        /**
         * Utility method that puts the named short value to the given map.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @param value property value
         */
        private static void PutShort(Dictionary<String, Object> properties, String name, short value)
        {
            if (properties.ContainsKey(name))
                properties[name] = value;
            else
                properties.Add(name, value);
        }

        /**
         * Utility method that puts the named boolean value to the given map.
         *
         * @param properties map of properties (String -> Object)
         * @param name property name
         * @param value property value
         */
        private static void PutBoolean(Dictionary<String, Object> properties, String name, bool value)
        {
            if (properties.ContainsKey(name))
                properties[name] = value;
            else
                properties.Add(name, value);
        }

        /**
         *  Looks for text in the cell that should be unicode, like an alpha and provides the
         *  unicode version of it.
         *
         *@param  cell  The cell to check for unicode values
         *@return       translated to unicode
         */
        public static ICell TranslateUnicodeValues(ICell cell)
        {
            String s = cell.RichStringCellValue.String;
            bool foundUnicode = false;
            String lowerCaseStr = s.ToLower();

            for (int i = 0; i < unicodeMappings.Length; i++)
            {
                UnicodeMapping entry = unicodeMappings[i];
                String key = entry.entityName;
                if (lowerCaseStr.IndexOf(key, StringComparison.Ordinal) != -1)
                {
                    s = s.Replace(key, entry.resolvedValue);
                    foundUnicode = true;
                }
            }
            if (foundUnicode)
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
